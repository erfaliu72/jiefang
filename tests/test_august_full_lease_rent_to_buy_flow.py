import json
import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta

import app as app_module
import database


class AugustFullLeaseRentToBuyFlowTestCase(unittest.TestCase):
    """2026-08 验收：仅覆盖租赁与以租代售的业务闭环和关键门禁。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-august-full-flow-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "acceptance.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.vehicle_seq = 0

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def login(self, username):
        response = self.client.post(
            "/api/auth/login",
            json={"username": username, "password": "123456"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def db_row(self, sql, params=()):
        conn = database.get_db()
        try:
            row = conn.execute(sql, params).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def db_value(self, sql, params=()):
        row = self.db_row(sql, params)
        return next(iter(row.values())) if row else None

    def create_vehicle(self, car_type, status="在库", box_type="厢货", condition="新车"):
        self.vehicle_seq += 1
        vin = f"V{self.vehicle_seq:016d}"
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles
                    (vin, plate_number, car_type, condition, status, box_type,
                     vehicle_box_type, validation_status, lock_status)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'valid', '未锁')
                """,
                (vin, f"陕验{self.vehicle_seq:04d}", car_type, condition, status, box_type, box_type),
            )
            vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return vehicle_id, vin
        finally:
            conn.close()

    def save_lease_guidance(self, car_type, deposit=2000, monthly=3000):
        self.login("boss")
        response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "is_new": "新车",
                "lease_deposit_guidance": deposit,
                "box_standard_price": monthly,
                "tail_plate_price": 300,
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def create_finance_plan(self, car_type, down_payment=18000, period_price=3600, periods=3, **extra):
        self.login("boss")
        payload = {
            "car_type": car_type,
            "plan_name": f"{car_type}-{periods}期",
            "down_payment": down_payment,
            "period_price": period_price,
            "periods": periods,
            "condition": "新车",
            "box_type": "厢货",
        }
        payload.update(extra)
        response = self.client.post("/api/finance-plans", json=payload)
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()["id"]

    def create_order(self, vin, car_type, sales_mode="租赁", **overrides):
        payload = {
            "payment_date": "2026-08-01",
            "customer_name": "八月全链路客户",
            "customer_phone": "13800000000",
            "sales_mode": sales_mode,
            "vin": vin,
            "car_type": car_type,
            "vehicle_box_type": "厢货",
            "lease_term": "3期",
            "deposit_amount": 2000,
            "vehicle_rent_amount": 3000,
            "customer_screenshot_path": "/uploads/customer-first-payment.jpg",
            "first_payment_received_amount": 5000,
        }
        payload.update(overrides)
        self.login("sales")
        response = self.client.post("/api/sales-orders", json=payload)
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()["id"]

    def approve_pending_flows(self, ref_type, ref_id):
        role_to_user = {"老板": "boss", "财务": "fin", "运营": "ops", "车管": "fleet", "销售": "sales"}
        while True:
            flow = self.db_row(
                """
                SELECT id, required_role
                FROM approval_flows
                WHERE ref_type=? AND ref_id=? AND status='待审批'
                ORDER BY step_order, id
                LIMIT 1
                """,
                (ref_type, ref_id),
            )
            if not flow:
                return
            self.login(role_to_user[flow["required_role"]])
            response = self.client.post(
                f"/api/approvals/{flow['id']}/approve",
                json={"comment": "八月验收审批通过"},
            )
            self.assertEqual(response.status_code, 200, response.get_json())

    def activate_upload_contract_and_deliver(self, order_id, vehicle_id, contract_type):
        self.login("fin")
        activation = self.client.post(
            f"/api/sales-orders/{order_id}/activate",
            json={"bank_serial": f"ACT{order_id:08d}", "bank_receipt_path": "/uploads/company-receipt.pdf"},
        )
        self.assertEqual(activation.status_code, 200, activation.get_json())

        self.login("ops")
        upload = self.client.post(
            "/api/contracts",
            json={
                "sales_order_id": order_id,
                "vehicle_id": vehicle_id,
                "contract_type": contract_type,
                "customer_name": "八月全链路客户",
                "customer_phone": "13800000000",
                "contract_file": "/uploads/signed-contract.pdf",
            },
        )
        self.assertEqual(upload.status_code, 200, upload.get_json())
        contract_id = upload.get_json()["id"]

        self.login("fleet")
        blocked = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(blocked.status_code, 400, blocked.get_json())
        saved = self.client.post(
            f"/api/contracts/{contract_id}/delivery-files",
            json={
                "delivery_photo_path": "/uploads/delivery-photo.jpg",
                "delivery_document_path": "/uploads/delivery-document.jpg",
            },
        )
        self.assertEqual(saved.status_code, 200, saved.get_json())
        delivered = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(delivered.status_code, 200, delivered.get_json())
        return contract_id

    def reconcile(self, repayment_id, received_amount, serial, extra_alloc_periods=None):
        self.login("ops")
        screenshot = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": f"/uploads/{serial}-customer.jpg", "reported_amount": received_amount},
        )
        self.assertEqual(screenshot.status_code, 200, screenshot.get_json())
        self.login("fin")
        receipt = self.client.post(
            f"/api/reconciliation/{repayment_id}/receipt",
            json={"bank_receipt_path": f"/uploads/{serial}-bank.pdf"},
        )
        self.assertEqual(receipt.status_code, 200, receipt.get_json())
        payload = {"bank_serial": serial, "received_amount": received_amount}
        if extra_alloc_periods is not None:
            payload["extra_alloc_periods"] = extra_alloc_periods
        verified = self.client.post(f"/api/reconciliation/{repayment_id}/verify", json=payload)
        return verified

    def test_inventory_vin_guidance_and_chassis_order_gate(self):
        priced_type = "八月入库车型"
        self.save_lease_guidance(priced_type)
        self.login("fleet")
        vin = "INV20260811000001"
        created = self.client.post(
            "/api/vehicles",
            json={
                "vin": vin,
                "plate_number": "陕A验001",
                "car_type": priced_type,
                "condition": "新车",
                "box_type": "厢货",
                "vehicle_box_type": "厢货",
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        duplicate = self.client.post(
            "/api/vehicles",
            json={"vin": vin, "car_type": priced_type, "condition": "新车", "box_type": "厢货"},
        )
        self.assertEqual(duplicate.status_code, 400, duplicate.get_json())

        unpriced = self.client.post(
            "/api/vehicles",
            json={
                "vin": "INV20260811000002",
                "plate_number": "陕A验002",
                "car_type": "八月未维护车型",
                "condition": "新车",
                "box_type": "厢货",
            },
        )
        self.assertEqual(unpriced.status_code, 200, unpriced.get_json())
        self.login("sales")
        blocked_unpriced = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-01",
                "customer_name": "未维护车型客户",
                "customer_phone": "13800000001",
                "sales_mode": "租赁",
                "vin": "INV20260811000002",
                "car_type": "八月未维护车型",
                "vehicle_box_type": "厢货",
                "lease_term": "3期",
                "deposit_amount": 2000,
                "vehicle_rent_amount": 3000,
                "customer_screenshot_path": "/uploads/unpriced.jpg",
                "first_payment_received_amount": 5000,
            },
        )
        self.assertEqual(blocked_unpriced.status_code, 400, blocked_unpriced.get_json())

        chassis_id, chassis_vin = self.create_vehicle(priced_type, box_type="底盘")
        self.login("sales")
        blocked_chassis = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-01",
                "customer_name": "底盘车客户",
                "customer_phone": "13800000002",
                "sales_mode": "租赁",
                "vin": chassis_vin,
                "car_type": priced_type,
                "vehicle_box_type": "厢货",
                "lease_term": "3期",
                "deposit_amount": 2000,
                "vehicle_rent_amount": 3000,
                "customer_screenshot_path": "/uploads/chassis.jpg",
                "first_payment_received_amount": 5000,
            },
        )
        self.assertEqual(chassis_id > 0, True)
        self.assertEqual(blocked_chassis.status_code, 400, blocked_chassis.get_json())

    def test_lease_price_exception_approval_finance_receipt_and_delivery(self):
        car_type = "八月租赁特批车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin,
            car_type,
            deposit_amount=1900,
            first_payment_received_amount=4900,
        )
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "待老板审批",
        )
        self.login("fin")
        blocked_activation = self.client.post(
            f"/api/sales-orders/{order_id}/activate",
            json={"bank_serial": "PREAPPROVE"},
        )
        self.assertEqual(blocked_activation.status_code, 400, blocked_activation.get_json())
        self.approve_pending_flows("order_exception", order_id)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "租赁中",
        )
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "已出库",
        )
        self.assertEqual(
            self.db_value("SELECT finance_bank_receipt_path FROM sales_orders WHERE id=?", (order_id,)),
            "/uploads/company-receipt.pdf",
        )

    def test_rent_to_buy_plan_dates_snapshot_shortage_and_ownership_transfer(self):
        car_type = "八月以租代售闭环车型"
        expired_plan = self.create_finance_plan(
            car_type,
            plan_name="已失效方案",
            effective_date="2026-01-01",
            expiry_date="2026-08-10",
        )
        active_plan = self.create_finance_plan(
            car_type,
            plan_name="验收有效方案",
            down_payment=18000,
            period_price=3600,
            periods=3,
            effective_date="2026-08-01",
            expiry_date="2026-12-31",
        )
        vehicle_id, vin = self.create_vehicle(car_type)
        self.login("sales")
        expired = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-01",
                "customer_name": "失效方案客户",
                "customer_phone": "13800000003",
                "sales_mode": "以租代售",
                "vin": vin,
                "car_type": car_type,
                "finance_plan_id": expired_plan,
                "customer_screenshot_path": "/uploads/expired-plan.jpg",
                "first_payment_received_amount": 18000,
            },
        )
        self.assertEqual(expired.status_code, 400, expired.get_json())

        order_id = self.create_order(
            vin,
            car_type,
            sales_mode="以租代售",
            finance_plan_id=active_plan,
            deposit_amount=18000,
            vehicle_rent_amount=1,
            lease_term="1期",
            first_payment_received_amount=17000,
            first_payment_shortage_reason="客户到账延迟",
            first_payment_promised_date="2026-08-20",
        )
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "待老板审批",
        )
        self.approve_pending_flows("order_exception", order_id)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "以租代售")
        snapshot = json.loads(
            self.db_value("SELECT snapshot_finance_plan FROM sales_orders WHERE id=?", (order_id,))
        )
        self.assertEqual(snapshot["id"], active_plan)
        self.assertEqual(snapshot["period_price"], 3600)
        self.assertEqual(
            self.db_value("SELECT COUNT(*) FROM repayments WHERE contract_id=? AND period>=1", (contract_id,)),
            3,
        )
        self.assertEqual(
            self.db_value(
                "SELECT status FROM repayments WHERE contract_id=? AND period=1",
                (contract_id,),
            ),
            "待还款",
        )
        receivable = self.db_row(
            "SELECT amount, status FROM receivables WHERE sales_order_id=? AND receivable_type='initial_payment_shortfall'",
            (order_id,),
        )
        self.assertEqual(receivable["amount"], 1000)
        self.assertEqual(receivable["status"], "待归还")

        conn = database.get_db()
        try:
            conn.execute(
                "UPDATE repayments SET status='已还款', paid_amount=amount, verified_amount=amount "
                "WHERE contract_id=? AND period>=1",
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()
        self.login("ops")
        transfer = self.client.post(
            f"/api/contracts/{contract_id}/ownership-transfer",
            json={"settle_type": "natural_settle", "idempotency_key": f"august-transfer-{contract_id}"},
        )
        self.assertEqual(transfer.status_code, 200, transfer.get_json())
        completed = self.client.post(
            f"/api/ownership-transfers/{transfer.get_json()['id']}/complete",
            json={"transfer_date": "2026-08-11", "new_owner_name": "以租代售客户"},
        )
        self.assertEqual(completed.status_code, 200, completed.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "已过户",
        )

    def test_reconciliation_shortfall_fee_priority_and_future_offset(self):
        car_type = "八月回款核销车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        # 标准交付链路的计划激活单独由下方门禁用例验收。这里将账单置为已激活，
        # 以隔离验证回款三段式对账与顺序扣款本身。
        conn = database.get_db()
        try:
            conn.execute(
                "UPDATE repayments SET status='待还款' WHERE contract_id=? AND period>=1",
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()
        period1 = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1", (contract_id,)
        )
        period2 = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2", (contract_id,)
        )
        self.login("ops")
        fee = self.client.post(
            f"/api/contracts/{contract_id}/fee-items",
            json={"fee_type": "insurance_fee", "amount_due": 500, "description": "保险服务费"},
        )
        self.assertEqual(fee.status_code, 200, fee.get_json())

        shortfall = self.reconcile(period1, 2500, "SHORTFALL001")
        self.assertEqual(shortfall.status_code, 200, shortfall.get_json())
        allocation = shortfall.get_json()["allocation"]
        self.assertEqual(allocation["rent_allocated"], 2000)
        self.assertEqual(allocation["shortfall_amount"], 1000)
        self.assertEqual(
            self.db_value(
                "SELECT status FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (period1,),
            ),
            "逾期应收",
        )

        settle_period1 = self.reconcile(period1, 1000, "SHORTFALL002")
        self.assertEqual(settle_period1.status_code, 200, settle_period1.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (period1,)),
            "已还款",
        )

        offset = self.reconcile(period2, 6000, "OVERPAY001", extra_alloc_periods=[3])
        self.assertEqual(offset.status_code, 200, offset.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=3", (contract_id,)),
            "已还款",
        )

    def test_delivery_should_activate_customer_repayment_plan(self):
        car_type = "八月计划激活车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        self.assertEqual(
            self.db_value(
                "SELECT status FROM repayments WHERE contract_id=? AND period=1",
                (contract_id,),
            ),
            "待还款",
        )

    def test_overdue_collection_lock_unlock_and_return_repair_cycle(self):
        car_type = "八月逾期退车车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1", (contract_id,)
        )
        overdue_date = (datetime.now().date() - timedelta(days=8)).strftime("%Y-%m-%d")
        conn = database.get_db()
        try:
            conn.execute(
                "UPDATE repayments SET due_date=?, status='待还款', paid_amount=0, verified_amount=0 WHERE id=?",
                (overdue_date, repayment_id),
            )
            conn.commit()
        finally:
            conn.close()
        app_module.run_daily_collect(force=True)

        self.login("ops")
        urge3 = self.client.post(
            f"/api/repayments/{repayment_id}/urge",
            json={"evidence_path": "/uploads/t3.jpg", "result": "运营已联系"},
        )
        self.assertEqual(urge3.status_code, 200, urge3.get_json())
        self.login("sales")
        urge7 = self.client.post(
            f"/api/repayments/{repayment_id}/urge",
            json={"evidence_path": "/uploads/t7.jpg", "result": "销售已联系"},
        )
        self.assertEqual(urge7.status_code, 200, urge7.get_json())
        lock = self.client.post(
            "/api/lock-requests",
            json={"repayment_id": repayment_id, "reason": "逾期八天锁车"},
        )
        self.assertEqual(lock.status_code, 200, lock.get_json())
        lock_id = lock.get_json()["lock_request_id"]
        self.login("ops")
        self.assertEqual(
            self.client.post(f"/api/lock-requests/{lock_id}/ops-review", json={"decision": "approve"}).status_code,
            200,
        )
        self.login("boss")
        self.assertEqual(
            self.client.post(f"/api/lock-requests/{lock_id}/boss-approve", json={"decision": "approve"}).status_code,
            200,
        )
        self.login("sales")
        self.assertEqual(self.client.post(f"/api/lock-requests/{lock_id}/confirm", json={}).status_code, 200)
        self.assertEqual(
            self.db_value("SELECT lock_status FROM vehicles WHERE id=?", (vehicle_id,)),
            "车辆已锁",
        )

        unlock = self.client.post("/api/unlock-requests", json={"vehicle_id": vehicle_id, "reason": "客户已补款"})
        self.assertEqual(unlock.status_code, 200, unlock.get_json())
        unlock_id = unlock.get_json()["lock_request_id"]
        self.login("ops")
        self.assertEqual(
            self.client.post(f"/api/lock-requests/{unlock_id}/ops-review", json={"decision": "approve"}).status_code,
            200,
        )
        self.login("boss")
        self.assertEqual(
            self.client.post(f"/api/lock-requests/{unlock_id}/boss-approve", json={"decision": "approve"}).status_code,
            200,
        )
        self.login("sales")
        self.assertEqual(self.client.post(f"/api/lock-requests/{unlock_id}/confirm", json={}).status_code, 200)

        return_created = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "到期退车", "remark": "八月验收退车"},
        )
        self.assertEqual(return_created.status_code, 200, return_created.get_json())
        return_id = return_created.get_json()["id"]
        self.login("fleet")
        self.assertEqual(
            self.client.post(
                f"/api/return-inspections/{return_id}/fleet",
                json={"needs_repair": True, "repair_reason": "轮胎磨损"},
            ).status_code,
            200,
        )
        self.login("ops")
        self.assertEqual(
            self.client.post(
                f"/api/return-inspections/{return_id}/operator",
                json={"deposit_paid": 2000, "total_deduction": 500, "actual_refund": 1500},
            ).status_code,
            200,
        )
        self.login("fin")
        self.assertEqual(
            self.client.post(
                f"/api/return-inspections/{return_id}/finance",
                json={"refund_bank_name": "", "refund_bank_card_no": ""},
            ).status_code,
            200,
        )
        self.login("boss")
        self.assertEqual(
            self.client.post(f"/api/return-inspections/{return_id}/boss-approve", json={}).status_code,
            200,
        )
        self.login("fin")
        paid = self.client.post(
            f"/api/return-inspections/{return_id}/pay",
            json={"refund_serial": "RETURN001", "refund_paid_amount": 1500},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "待维修",
        )
        self.login("fleet")
        self.assertEqual(self.client.post(f"/api/vehicles/{vehicle_id}/repair/start", json={}).status_code, 200)
        repaired = self.client.post(f"/api/vehicles/{vehicle_id}/repair/complete", json={})
        self.assertEqual(repaired.status_code, 200, repaired.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "在库",
        )


if __name__ == "__main__":
    unittest.main()
