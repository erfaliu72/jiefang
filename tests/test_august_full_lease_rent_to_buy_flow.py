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
        blocked = self.client.post(
            f"/api/contracts/{contract_id}/ownership-transfer",
            json={"settle_type": "natural_settle", "idempotency_key": f"august-transfer-{contract_id}"},
        )
        self.assertEqual(blocked.status_code, 400, blocked.get_json())
        self.assertIn("首付款", blocked.get_json()["message"])

        receivable_id = self.db_value(
            "SELECT id FROM receivables WHERE sales_order_id=? AND receivable_type='initial_payment_shortfall'",
            (order_id,),
        )
        screenshot = self.client.post(
            f"/api/receivables/{receivable_id}/screenshot",
            json={"screenshot_path": "/uploads/rent-to-buy-shortfall.jpg"},
        )
        self.assertEqual(screenshot.status_code, 200, screenshot.get_json())
        self.login("fin")
        settled = self.client.post(
            f"/api/receivables/{receivable_id}/settle",
            json={"bank_serial": "DOWNPAY001", "received_amount": 1000},
        )
        self.assertEqual(settled.status_code, 200, settled.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=0", (contract_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value("SELECT down_payment_status FROM contracts WHERE id=?", (contract_id,)),
            "已收",
        )

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
        transfer_id = transfer.get_json()["id"]
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "已过户",
        )
        history = self.client.get("/api/completion-history?source_type=ownership_transfer")
        self.assertEqual(history.status_code, 200, history.get_json())
        record = next((row for row in history.get_json() if row["source_id"] == transfer_id), None)
        self.assertIsNotNone(record)
        self.assertEqual(record["contract_id"], contract_id)
        self.assertEqual(record["completion_type"], "以租代售过户")
        self.assertEqual(record["completion_result"], "车辆已过户")
        self.assertEqual(record["handled_by"], "李运营")
        self.assertTrue(record["completed_at"])
        keyword_history = self.client.get(
            f"/api/completion-history?source_type=ownership_transfer&keyword={vin}"
        )
        self.assertEqual(keyword_history.status_code, 200, keyword_history.get_json())
        self.assertTrue(any(row["source_id"] == transfer_id for row in keyword_history.get_json()))

    def test_reconciliation_shortfall_fee_priority_and_future_offset(self):
        car_type = "八月回款核销车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type, lease_term="4期")
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        conn = database.get_db()
        try:
            conn.execute(
                """
                UPDATE repayments
                SET due_date='2026-08-01', status='待还款', paid_amount=0, verified_amount=0
                WHERE contract_id=? AND period=2
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()
        period2 = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2", (contract_id,)
        )
        period3 = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=3", (contract_id,)
        )
        self.login("ops")
        fee = self.client.post(
            f"/api/contracts/{contract_id}/fee-items",
            json={"fee_type": "insurance_fee", "amount_due": 500, "description": "保险服务费"},
        )
        self.assertEqual(fee.status_code, 200, fee.get_json())

        shortfall = self.reconcile(period2, 2500, "SHORTFALL001")
        self.assertEqual(shortfall.status_code, 200, shortfall.get_json())
        allocation = shortfall.get_json()["allocation"]
        self.assertEqual(allocation["rent_allocated"], 2000)
        self.assertEqual(allocation["shortfall_amount"], 1000)
        self.assertEqual(
            self.db_value(
                "SELECT status FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (period2,),
            ),
            "逾期应收",
        )

        settle_period2 = self.reconcile(period2, 1000, "SHORTFALL002")
        self.assertEqual(settle_period2.status_code, 200, settle_period2.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (period2,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value(
                "SELECT status FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (period2,),
            ),
            "已结清",
        )

        offset = self.reconcile(period3, 6000, "OVERPAY001", extra_alloc_periods=[4])
        self.assertEqual(offset.status_code, 200, offset.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=4", (contract_id,)),
            "已还款",
        )

    def test_reconciliation_partial_payment_requires_new_voucher_before_supplement(self):
        car_type = "八月补款累计核销车型"
        self.save_lease_guidance(car_type, monthly=5000)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin,
            car_type,
            lease_term="6期",
            vehicle_rent_amount=5000,
            first_payment_received_amount=7000,
        )
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2",
            (contract_id,),
        )

        self.login("ops")
        first_submission = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/period-2-first-4500.jpg", "reported_amount": 4500},
        )
        self.assertEqual(first_submission.status_code, 200, first_submission.get_json())
        overwritten = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/period-2-supplement-500.jpg", "reported_amount": 500},
        )
        self.assertEqual(overwritten.status_code, 400, overwritten.get_json())
        self.assertIn("等待财务核销", overwritten.get_json()["message"])
        pending = self.db_row(
            "SELECT screenshot_path, reported_amount FROM repayments WHERE id=?",
            (repayment_id,),
        )
        self.assertEqual(pending["screenshot_path"], "/uploads/period-2-first-4500.jpg")
        self.assertEqual(pending["reported_amount"], 4500)

        self.login("fin")
        first_verified = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "P2FIRST4500", "received_amount": 4500},
        )
        self.assertEqual(first_verified.status_code, 200, first_verified.get_json())
        self.assertEqual(
            self.db_value("SELECT paid_amount FROM repayments WHERE id=?", (repayment_id,)),
            4500,
        )
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (repayment_id,)),
            "部分核销",
        )
        self.assertEqual(
            self.db_value(
                "SELECT amount FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (repayment_id,),
            ),
            500,
        )
        duplicate_verify = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "P2DUP4500", "received_amount": 500},
        )
        self.assertEqual(duplicate_verify.status_code, 400, duplicate_verify.get_json())
        self.assertIn("上传补款凭证", duplicate_verify.get_json()["message"])

        self.login("ops")
        supplement_submission = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/period-2-supplement-500.jpg", "reported_amount": 500},
        )
        self.assertEqual(supplement_submission.status_code, 200, supplement_submission.get_json())
        self.login("fin")
        supplement_verified = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "P2SUPP0500", "received_amount": 500},
        )
        self.assertEqual(supplement_verified.status_code, 200, supplement_verified.get_json())
        repayment = self.db_row(
            "SELECT paid_amount, verified_amount, status FROM repayments WHERE id=?",
            (repayment_id,),
        )
        self.assertEqual(repayment["paid_amount"], 5000)
        self.assertEqual(repayment["verified_amount"], 5000)
        self.assertEqual(repayment["status"], "已还款")
        self.assertEqual(
            self.db_value(
                "SELECT status FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (repayment_id,),
            ),
            "已结清",
        )

    def test_rent_waiver_closes_the_linked_shortfall_receivable(self):
        vehicle_id, _ = self.create_vehicle("八月租金减免同步车型")
        conn = database.get_db()
        try:
            conn.execute("INSERT INTO customers (name, phone) VALUES ('减免客户', '13800000021')")
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, contract_status, delivery_status)
                VALUES (?, ?, '租赁', '执行中', '已出库')
                """,
                (vehicle_id, customer_id),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO repayments
                    (contract_id, period, due_date, amount, paid_amount, verified_amount, status)
                VALUES (?, 1, '2026-08-01', 3000, 2000, 2000, '部分核销')
                """,
                (contract_id,),
            )
            repayment_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO receivables
                    (contract_id, repayment_id, receivable_type, source_period, amount, status)
                VALUES (?, ?, 'period_shortfall', 1, 1000, '逾期应收')
                """,
                (contract_id, repayment_id),
            )
            conn.execute(
                """
                INSERT INTO waivers
                    (contract_id, waiver_kind, target_period_list, waive_amount, reason, status)
                VALUES (?, 'rent', '[1]', 1000, '到期减免', '已通过')
                """,
                (contract_id,),
            )
            waiver_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
        finally:
            conn.close()

        self.login("fin")
        response = self.client.post(f"/api/waivers/{waiver_id}/execute", json={})
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (repayment_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value(
                "SELECT status FROM receivables WHERE repayment_id=? AND receivable_type='period_shortfall'",
                (repayment_id,),
            ),
            "已取消",
        )

    def test_early_settlement_includes_initial_payment_shortfall_and_closes_it(self):
        vehicle_id, _ = self.create_vehicle("八月提前结清首付车型", status="以租代售")
        conn = database.get_db()
        try:
            conn.execute("INSERT INTO customers (name, phone) VALUES ('提前结清客户', '13800000022')")
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, contract_status, delivery_status, down_payment)
                VALUES (?, ?, '以租代售', '执行中', '已出库', 1000)
                """,
                (vehicle_id, customer_id),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.executemany(
                """
                INSERT INTO repayments (contract_id, period, due_date, amount, paid_amount, status)
                VALUES (?, ?, '2026-08-01', ?, 0, '待还款')
                """,
                [(contract_id, 0, 1000), (contract_id, 1, 2000)],
            )
            conn.execute(
                """
                INSERT INTO receivables
                    (contract_id, receivable_type, source_period, amount, status)
                VALUES (?, 'initial_payment_shortfall', 0, 1000, '待归还')
                """,
                (contract_id,),
            )
            receivable_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
        finally:
            conn.close()

        self.login("ops")
        quote = self.client.get(f"/api/contracts/{contract_id}/early-settlement")
        self.assertEqual(quote.status_code, 200, quote.get_json())
        self.assertEqual(quote.get_json()["outstanding_installments"], 3000)
        settled = self.client.post(
            f"/api/contracts/{contract_id}/early-settlement",
            json={"bank_serial": "EARLY20260821", "received_amount": 3000},
        )
        self.assertEqual(settled.status_code, 200, settled.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=0", (contract_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM receivables WHERE id=?", (receivable_id,)),
            "已结清",
        )

    def test_ownership_transfer_blocks_unsettled_contract_fees(self):
        vehicle_id, _ = self.create_vehicle("八月过户费用门禁车型", status="以租代售")
        conn = database.get_db()
        try:
            conn.execute("INSERT INTO customers (name, phone) VALUES ('过户费用客户', '13800000023')")
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, contract_status, delivery_status)
                VALUES (?, ?, '以租代售', '执行中', '已出库')
                """,
                (vehicle_id, customer_id),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO repayments (contract_id, period, due_date, amount, paid_amount, status)
                VALUES (?, 0, '2026-08-01', 1000, 1000, '已还款')
                """,
                (contract_id,),
            )
            conn.execute(
                """
                INSERT INTO contract_fee_items
                    (contract_id, fee_type, description, amount_due, amount_paid, status)
                VALUES (?, 'service_fee', '未结清服务费', 300, 0, '待支付')
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()

        self.login("ops")
        response = self.client.post(
            f"/api/contracts/{contract_id}/ownership-transfer",
            json={"idempotency_key": f"fee-transfer-{contract_id}"},
        )
        self.assertEqual(response.status_code, 400, response.get_json())
        self.assertIn("合同费用", response.get_json()["message"])

    def test_pending_ownership_transfer_todo_only_lists_settled_contracts(self):
        vehicle_id, vin = self.create_vehicle("八月待办过户车型", status="以租代售")
        conn = database.get_db()
        try:
            conn.execute("INSERT INTO customers (name, phone) VALUES ('待办过户客户', '13800000025')")
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, contract_status, delivery_status)
                VALUES (?, ?, '以租代售', '执行中', '已出库')
                """,
                (vehicle_id, customer_id),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO repayments (contract_id, period, due_date, amount, paid_amount, status)
                VALUES (?, 0, '2026-08-01', 1000, 1000, '已还款')
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()

        self.login("ops")
        todo = self.client.get("/api/ownership-transfers/pending")
        self.assertEqual(todo.status_code, 200, todo.get_json())
        row = next((item for item in todo.get_json()["items"] if item["contract_id"] == contract_id), None)
        self.assertIsNotNone(row)
        self.assertEqual(row["vin"], vin)
        self.assertEqual(row["pending_kind"], "待发起")
        self.assertIsNone(row["transfer_id"])

        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO contract_fee_items
                    (contract_id, fee_type, description, amount_due, amount_paid, status)
                VALUES (?, 'service_fee', '待收服务费', 300, 0, '待支付')
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()
        blocked_todo = self.client.get("/api/ownership-transfers/pending")
        self.assertEqual(blocked_todo.status_code, 200, blocked_todo.get_json())
        self.assertFalse(any(item["contract_id"] == contract_id for item in blocked_todo.get_json()["items"]))

        conn = database.get_db()
        try:
            conn.execute(
                """
                UPDATE contract_fee_items
                SET amount_paid=amount_due, status='已支付'
                WHERE contract_id=?
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()
        created = self.client.post(
            f"/api/contracts/{contract_id}/ownership-transfer",
            json={"idempotency_key": f"pending-transfer-{contract_id}"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        pending_created = self.client.get("/api/ownership-transfers/pending")
        self.assertEqual(pending_created.status_code, 200, pending_created.get_json())
        created_row = next(
            (item for item in pending_created.get_json()["items"] if item["contract_id"] == contract_id),
            None,
        )
        self.assertIsNotNone(created_row)
        self.assertEqual(created_row["transfer_id"], created.get_json()["id"])
        self.assertEqual(created_row["pending_kind"], "待完成")

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
            "已还款",
        )
        self.assertEqual(
            self.db_value(
                "SELECT status FROM repayments WHERE contract_id=? AND period=2",
                (contract_id,),
            ),
            "待还款",
        )

    def test_rental_initial_shortfall_settlement_syncs_first_installment(self):
        car_type = "八月租赁首付差额车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin,
            car_type,
            first_payment_received_amount=3500,
            first_payment_shortage_reason="首月租金差额",
            first_payment_promised_date="2026-08-25",
        )
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "待老板审批",
        )
        self.approve_pending_flows("order_exception", order_id)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        receivable_id = self.db_value(
            "SELECT id FROM receivables WHERE sales_order_id=? AND receivable_type='initial_payment_shortfall'",
            (order_id,),
        )
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=0", (contract_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value("SELECT paid_amount FROM repayments WHERE contract_id=? AND period=1", (contract_id,)),
            1500,
        )

        self.login("ops")
        screenshot = self.client.post(
            f"/api/receivables/{receivable_id}/screenshot",
            json={"screenshot_path": "/uploads/lease-shortfall.jpg"},
        )
        self.assertEqual(screenshot.status_code, 200, screenshot.get_json())
        self.login("fin")
        partial = self.client.post(
            f"/api/receivables/{receivable_id}/settle",
            json={"bank_serial": "LEASESHORT001", "received_amount": 500},
        )
        self.assertEqual(partial.status_code, 200, partial.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM receivables WHERE id=?", (receivable_id,)),
            "部分归还",
        )
        self.assertEqual(
            self.db_value("SELECT paid_amount FROM repayments WHERE contract_id=? AND period=1", (contract_id,)),
            2000,
        )
        settled = self.client.post(
            f"/api/receivables/{receivable_id}/settle",
            json={"bank_serial": "LEASESHORT002", "received_amount": 1000},
        )
        self.assertEqual(settled.status_code, 200, settled.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM receivables WHERE id=?", (receivable_id,)),
            "已结清",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE contract_id=? AND period=1", (contract_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value("SELECT collected_rent FROM contracts WHERE id=?", (contract_id,)),
            3000,
        )

    def test_reconciliation_closes_initial_shortfall_when_first_rent_is_covered(self):
        car_type = "八月首期核销关闭挂账车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin,
            car_type,
            first_payment_received_amount=3500,
            first_payment_shortage_reason="首月租金差额",
            first_payment_promised_date="2026-08-25",
        )
        self.approve_pending_flows("order_exception", order_id)
        contract_id = self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        receivable_id = self.db_value(
            "SELECT id FROM receivables WHERE sales_order_id=? AND receivable_type='initial_payment_shortfall'",
            (order_id,),
        )
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1",
            (contract_id,),
        )

        self.login("ops")
        initiated = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/first-rent-shortfall.jpg", "reported_amount": 1500},
        )
        self.assertEqual(initiated.status_code, 200, initiated.get_json())
        self.login("fin")
        verified = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "FIRSTRENT500", "received_amount": 1500},
        )
        self.assertEqual(verified.status_code, 200, verified.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (repayment_id,)),
            "已还款",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM receivables WHERE id=?", (receivable_id,)),
            "已结清",
        )
        self.assertEqual(
            self.db_value("SELECT paid_amount FROM receivables WHERE id=?", (receivable_id,)),
            1500,
        )

    def test_delivered_contract_plan_cannot_be_regenerated(self):
        car_type = "八月计划防覆盖车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        self.activate_upload_contract_and_deliver(order_id, vehicle_id, "租赁")
        self.login("ops")
        response = self.client.put(
            f"/api/sales-orders/{order_id}/planning-contract",
            json={"loan_periods": 6, "rent": 9999},
        )
        self.assertEqual(response.status_code, 400, response.get_json())
        self.assertIn("不能重新生成客户还款计划", response.get_json()["message"])

    def test_initial_rent_correction_reverses_only_the_wrong_late_fee(self):
        vehicle_id, _ = self.create_vehicle("八月滞纳金冲回车型", status="租赁中")
        conn = database.get_db()
        try:
            conn.execute("INSERT INTO customers (name, phone) VALUES ('滞纳金客户', '13800000024')")
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute("""
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, contract_status, delivery_status)
                VALUES (?, ?, '租赁', '执行中', '已出库')
            """, (vehicle_id, customer_id))
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.executemany("""
                INSERT INTO repayments (contract_id, period, due_date, amount, status)
                VALUES (?, ?, ?, 4000, '待还款')
            """, [
                (contract_id, 1, '2026-08-14'),
                (contract_id, 2, '2026-08-14'),
            ])
            period1 = conn.execute(
                "SELECT id FROM repayments WHERE contract_id=? AND period=1", (contract_id,)
            ).fetchone()[0]
            period2 = conn.execute(
                "SELECT id FROM repayments WHERE contract_id=? AND period=2", (contract_id,)
            ).fetchone()[0]
            conn.executemany("""
                INSERT INTO late_fee_ledger
                    (repayment_id, contract_id, period, accrued_date, outstanding, daily_amount, cumulative_amount)
                VALUES (?, ?, ?, ?, 4000, 2, ?)
            """, [
                (period1, contract_id, 1, '2026-08-15', 2),
                (period1, contract_id, 1, '2026-08-16', 4),
                (period2, contract_id, 2, '2026-08-15', 2),
                (period2, contract_id, 2, '2026-08-16', 4),
            ])
            conn.execute("""
                INSERT INTO contract_fee_items
                    (contract_id, fee_type, description, amount_due, amount_paid, status)
                VALUES (?, 'late_fee', '逾期滞纳金(系统计提)', 8, 0, '待支付')
            """, (contract_id,))
            conn.commit()

            reversed_count = app_module.reverse_late_fee_accruals_from_payment_date(
                conn, period1, '2026-08-14', '报单首期租金已收'
            )
            conn.commit()
        finally:
            conn.close()

        self.assertEqual(reversed_count, 2)
        self.assertEqual(
            self.db_value("SELECT SUM(waived_amount) FROM late_fee_ledger WHERE repayment_id=?", (period1,)),
            4,
        )
        self.assertEqual(
            self.db_value(
                "SELECT amount_due FROM contract_fee_items WHERE contract_id=? AND fee_type='late_fee'",
                (contract_id,),
            ),
            4,
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

        conn = database.get_db()
        try:
            conn.execute(
                "UPDATE repayments SET paid_amount=amount, verified_amount=amount, status='已还款' WHERE id=?",
                (repayment_id,),
            )
            conn.execute("""
                UPDATE contract_fee_items
                SET amount_paid=amount_due, status='已支付'
                WHERE contract_id=? AND amount_due>COALESCE(amount_paid, 0)
            """, (contract_id,))
            conn.commit()
        finally:
            conn.close()
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
                json={
                    "needs_repair": True, "repair_reason": "轮胎磨损",
                    "mileage": "30000", "body_tire_clean": "已清理", "accident_info": "无出险",
                    "insurance_surcharge": "无", "violation_info": "无违章", "etc_info": "已注销",
                    "maintenance_info": "正常",
                },
            ).status_code,
            200,
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "退车中",
        )
        self.login("ops")
        self.assertEqual(
            self.client.post(
                f"/api/return-inspections/{return_id}/operator",
                json={
                    "rent_late_fee": 0, "return_late_fee": 0, "deposit_rent_receivable": 0,
                    "deposit_paid": 2000, "total_deduction": 500, "actual_refund": 1500,
                },
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
        wrong_amount = self.client.post(
            f"/api/return-inspections/{return_id}/pay",
            json={"refund_serial": "RETURN001", "refund_paid_amount": 1400},
        )
        self.assertEqual(wrong_amount.status_code, 400, wrong_amount.get_json())
        paid = self.client.post(
            f"/api/return-inspections/{return_id}/pay",
            json={"refund_serial": "RETURN001", "refund_paid_amount": 1500},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM repayments WHERE contract_id=? AND period>=1 AND status='已取消'",
                (contract_id,),
            ),
            2,
        )
        self.login("fin")
        cancelled_reconciliation = self.client.post(
            f"/api/repayments/{self.db_value('SELECT id FROM repayments WHERE contract_id=? AND period=2', (contract_id,))}/confirm",
            json={"received_amount": 2000},
        )
        self.assertEqual(cancelled_reconciliation.status_code, 400, cancelled_reconciliation.get_json())
        self.assertIn("退车结算取消", cancelled_reconciliation.get_json()["message"])
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "待维修",
        )
        self.assertEqual(
            self.db_value("SELECT condition FROM vehicles WHERE id=?", (vehicle_id,)),
            "二手车",
        )
        self.login("fleet")
        missing_repair_note = self.client.post(f"/api/vehicles/{vehicle_id}/repair/complete", json={})
        self.assertEqual(missing_repair_note.status_code, 400, missing_repair_note.get_json())
        self.assertIn("维修完成情况", missing_repair_note.get_json()["message"])
        repaired = self.client.post(
            f"/api/vehicles/{vehicle_id}/repair/complete",
            json={
                "repair_completed_at": "2026-08-22",
                "repair_completion_note": "已更换磨损轮胎并完成路试",
                "repair_cost": 500,
            },
        )
        self.assertEqual(repaired.status_code, 200, repaired.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "在库",
        )
        self.assertEqual(
            self.db_value("SELECT condition FROM vehicles WHERE id=?", (vehicle_id,)),
            "二手车",
        )
        self.assertEqual(
            self.db_value("SELECT repair_completion_note FROM return_inspections WHERE id=?", (return_id,)),
            "已更换磨损轮胎并完成路试",
        )

    def test_init_db_backfills_completed_returned_vehicle_as_used(self):
        vehicle_id, _ = self.create_vehicle("八月历史退车回填车型", condition="新车")
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO return_inspections (vehicle_id, status, paid_out)
                VALUES (?, '已完成', 1)
                """,
                (vehicle_id,),
            )
            conn.commit()
        finally:
            conn.close()

        database.init_db()

        self.assertEqual(
            self.db_value("SELECT condition FROM vehicles WHERE id=?", (vehicle_id,)),
            "二手车",
        )


if __name__ == "__main__":
    unittest.main()
