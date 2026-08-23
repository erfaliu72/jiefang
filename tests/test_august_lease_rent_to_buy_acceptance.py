import json
import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta

import app as app_module
import database


class AugustLeaseRentToBuyAcceptanceTestCase(unittest.TestCase):
    """2026-08 最新流程验收：只覆盖租赁与以租代售。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-august-acceptance-")
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

    def create_vehicle(self, car_type, status="在库"):
        self.vehicle_seq += 1
        vin = f"V{self.vehicle_seq:016d}"
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles
                    (vin, plate_number, car_type, condition, status, validation_status)
                VALUES (?, ?, ?, '新车', ?, 'valid')
                """,
                (vin, f"陕验{self.vehicle_seq:04d}", car_type, status),
            )
            vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return vehicle_id, vin
        finally:
            conn.close()

    def create_active_rental_contract(self, needs_repair=False):
        vehicle_id, _ = self.create_vehicle("八月租赁退车车型", "租赁中")
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO customers (name, phone)
                VALUES ('八月验收客户', '13800000000')
                """
            )
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, start_date, end_date, rent,
                     loan_periods, deposit, contract_file, contract_status, delivery_status,
                     lease_bank_name, lease_bank_card_no)
                VALUES (?, ?, '租赁', '2026-01-01', '2026-12-31', 3000, 12, 2000,
                        '/uploads/accepted-contract.pdf', '执行中', '已出库',
                        '测试银行', '6222000012345678')
                """,
                (vehicle_id, customer_id),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return vehicle_id, contract_id, needs_repair
        finally:
            conn.close()

    def test_guidance_unique_manual_contract_and_finance_receipt(self):
        car_type = "八月租赁指导价车型"
        base_car_type = app_module.normalize_base_car_type(car_type)
        self.login("boss")
        for deposit, monthly in ((2000, 3000), (2200, 3200)):
            response = self.client.post(
                "/api/model-guidance-prices",
                json={
                    "car_type": car_type,
                    "is_new": "新车",
                    "lease_deposit_guidance": deposit,
                    "box_standard_price": monthly,
                },
            )
            self.assertEqual(response.status_code, 200, response.get_json())
        response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "is_new": "二手车",
                "lease_deposit_guidance": 1500,
                "box_standard_price": 2500,
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM model_guidance_prices WHERE car_type=?",
                (base_car_type,),
            ),
            2,
        )
        self.assertEqual(
            self.db_value(
                "SELECT lease_deposit_guidance FROM model_guidance_prices WHERE car_type=? AND is_new='新车'",
                (base_car_type,),
            ),
            2200,
        )

        manual_vehicle_id, _ = self.create_vehicle("八月手工合同车型")
        self.login("ops")
        manual_contract = self.client.post(
            "/api/contracts",
            json={
                "vehicle_id": manual_vehicle_id,
                "customer_name": "手工合同客户",
                "customer_phone": "13800000001",
                "contract_type": "租赁",
                "start_date": "2026-08-01",
                "loan_periods": 2,
                "rent": 3000,
                "deposit": 2000,
                "contract_file": "/uploads/manual-contract.pdf",
            },
        )
        self.assertEqual(manual_contract.status_code, 200, manual_contract.get_json())

        order_vehicle_id, order_vin = self.create_vehicle(car_type)
        self.login("sales")
        order = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-01",
                "customer_name": "租赁报单客户",
                "customer_phone": "13800000002",
                "sales_mode": "租赁",
                "vin": order_vin,
                "car_type": car_type,
                "vehicle_box_type": "厢货",
                "lease_term": "12期",
                "deposit_amount": 2200,
                "vehicle_rent_amount": 3200,
                "customer_screenshot_path": "/uploads/lease-first-payment.jpg",
                "first_payment_received_amount": 5400,
            },
        )
        self.assertEqual(order.status_code, 200, order.get_json())
        order_id = order.get_json()["id"]
        self.login("fin")
        activation = self.client.post(
            f"/api/sales-orders/{order_id}/activate",
            json={
                "bank_serial": "LEASE202608001",
                "bank_receipt_path": "/uploads/lease-bank-receipt.pdf",
            },
        )
        self.assertEqual(activation.status_code, 200, activation.get_json())
        persisted = self.db_row(
            """
            SELECT order_status, finance_bank_serial, finance_bank_receipt_path
            FROM sales_orders WHERE id=?
            """,
            (order_id,),
        )
        self.assertEqual(persisted["order_status"], "已激活")
        self.assertEqual(persisted["finance_bank_serial"], "LEASE202608001")
        self.assertEqual(persisted["finance_bank_receipt_path"], "/uploads/lease-bank-receipt.pdf")
        self.assertEqual(order_vehicle_id > 0, True)

    def test_approval_center_finance_confirm_persists_sales_order_receipt(self):
        car_type = "审批中心财务确认车型"
        self.login("boss")
        guidance = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "is_new": "新车",
                "lease_deposit_guidance": 2000,
                "box_standard_price": 3000,
            },
        )
        self.assertEqual(guidance.status_code, 200, guidance.get_json())
        _, vin = self.create_vehicle(car_type)
        self.login("sales")
        order = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-22",
                "customer_name": "审批中心财务客户",
                "customer_phone": "13800000008",
                "sales_mode": "租赁",
                "vin": vin,
                "car_type": car_type,
                "vehicle_box_type": "厢货",
                "lease_term": "3期",
                "deposit_amount": 2000,
                "vehicle_rent_amount": 3000,
                "customer_screenshot_path": "/uploads/approval-center-payment.jpg",
                "first_payment_received_amount": 5000,
            },
        )
        self.assertEqual(order.status_code, 200, order.get_json())
        order_id = order.get_json()["id"]
        flow = self.db_row(
            """
            SELECT id FROM approval_flows
            WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
            """,
            (order_id,),
        )
        self.assertIsNotNone(flow)

        self.login("fin")
        approved = self.client.post(
            f"/api/approvals/{flow['id']}/approve",
            json={
                "comment": "审批中心确认到账",
                "bank_serial": "APPROVAL20260822",
                "bank_receipt_path": "/uploads/approval-center-receipt.pdf",
            },
        )
        self.assertEqual(approved.status_code, 200, approved.get_json())
        order_row = self.db_row(
            """
            SELECT order_status, finance_bank_serial, finance_bank_receipt_path
            FROM sales_orders WHERE id=?
            """,
            (order_id,),
        )
        self.assertEqual(order_row["order_status"], "已激活")
        self.assertEqual(order_row["finance_bank_serial"], "APPROVAL20260822")
        self.assertEqual(order_row["finance_bank_receipt_path"], "/uploads/approval-center-receipt.pdf")
        self.assertEqual(
            self.db_value("SELECT status FROM approval_flows WHERE id=?", (flow["id"],)),
            "已通过",
        )

    def test_manual_contract_requires_initial_payment_before_delivery(self):
        vehicle_id, _ = self.create_vehicle("八月手工首次付款车型")
        self.login("ops")
        created = self.client.post(
            "/api/contracts",
            json={
                "vehicle_id": vehicle_id,
                "customer_name": "手工合同首付客户",
                "customer_phone": "13800000009",
                "contract_type": "租赁",
                "start_date": "2026-08-12",
                "loan_periods": 2,
                "rent": 3000,
                "deposit": 2000,
                "contract_file": "/uploads/manual-payment-contract.pdf",
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        contract_id = created.get_json()["id"]
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待首付款",
        )

        self.login("fleet")
        blocked = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(blocked.status_code, 400, blocked.get_json())

        self.login("ops")
        initiated = self.client.post(
            f"/api/contracts/{contract_id}/initial-payment",
            json={
                "customer_screenshot_path": "/uploads/manual-customer-payment.jpg",
                "remark": "八月手工合同首付",
            },
        )
        self.assertEqual(initiated.status_code, 200, initiated.get_json())
        payment_id = initiated.get_json()["id"]

        self.login("fin")
        receipt = self.client.post(
            f"/api/initial-payments/{payment_id}/receipt",
            json={
                "bank_receipt_path": "/uploads/manual-bank-receipt.pdf",
                "bank_serial": "MANUAL20260812",
                "received_amount": 5000,
            },
        )
        self.assertEqual(receipt.status_code, 200, receipt.get_json())
        flow = self.db_row(
            """
            SELECT id FROM approval_flows
            WHERE ref_type='initial_payment' AND ref_id=? AND status='待审批'
            ORDER BY id DESC LIMIT 1
            """,
            (payment_id,),
        )
        approved = self.client.post(
            f"/api/approvals/{flow['id']}/approve",
            json={"comment": "手工合同首付核验通过"},
        )
        self.assertEqual(approved.status_code, 200, approved.get_json())
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待出库",
        )

    def test_rent_to_buy_plan_snapshot_generates_contract_plan(self):
        car_type = "八月以租代售车型"
        self.login("boss")
        plan_response = self.client.post(
            "/api/finance-plans",
            json={
                "car_type": car_type,
                "plan_name": "八月36期方案",
                "down_payment": 18000,
                "period_price": 3600,
                "periods": 36,
                "condition": "新车",
                "box_type": "厢货",
            },
        )
        self.assertEqual(plan_response.status_code, 200, plan_response.get_json())
        plan_id = plan_response.get_json()["id"]
        _, vin = self.create_vehicle(car_type)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-01",
                "customer_name": "以租代售客户",
                "customer_phone": "13800000003",
                "sales_mode": "以租代售",
                "vin": vin,
                "car_type": car_type,
                "finance_plan_id": plan_id,
                "deposit_amount": 18000,
                "vehicle_rent_amount": 1,
                "lease_term": "1期",
                "customer_screenshot_path": "/uploads/rent-to-buy-first-payment.jpg",
                "first_payment_received_amount": 18000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        order = self.db_row(
            "SELECT contract_id, snapshot_finance_plan FROM sales_orders WHERE id=?",
            (order_id,),
        )
        snapshot = json.loads(order["snapshot_finance_plan"])
        self.assertEqual(snapshot["periods"], 36)
        self.assertEqual(snapshot["period_price"], 3600)
        contract = self.db_row(
            "SELECT loan_periods, rent, down_payment FROM contracts WHERE id=?",
            (order["contract_id"],),
        )
        self.assertEqual(contract["loan_periods"], 36)
        self.assertEqual(contract["rent"], 3600)
        self.assertEqual(contract["down_payment"], 18000)

        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM repayments WHERE contract_id=? AND period>=1",
                (order["contract_id"],),
            ),
            36,
        )

        self.login("fin")
        activated = self.client.post(
            f"/api/sales-orders/{order_id}/activate",
            json={
                "bank_serial": "RTB202608001",
                "bank_receipt_path": "/uploads/rent-to-buy-bank-receipt.pdf",
            },
        )
        self.assertEqual(activated.status_code, 200, activated.get_json())
        self.login("ops")
        uploaded = self.client.post(
            "/api/contracts",
            json={
                "sales_order_id": order_id,
                "vehicle_id": self.db_value(
                    "SELECT vehicle_id FROM sales_orders WHERE id=?",
                    (order_id,),
                ),
                "contract_type": "以租代售",
                "customer_name": "以租代售客户",
                "customer_phone": "13800000003",
                "contract_file": "/uploads/rent-to-buy-contract.pdf",
                "loan_periods": 1,
                "rent": 1,
                "down_payment": 1,
            },
        )
        self.assertEqual(uploaded.status_code, 200, uploaded.get_json())
        uploaded_contract = self.db_row(
            "SELECT loan_periods, rent, down_payment FROM contracts WHERE id=?",
            (order["contract_id"],),
        )
        self.assertEqual(uploaded_contract["loan_periods"], 36)
        self.assertEqual(uploaded_contract["rent"], 3600)
        self.assertEqual(uploaded_contract["down_payment"], 18000)

    def test_finance_plan_only_vehicle_can_submit_rent_to_buy_order(self):
        car_type = "八月仅方案车型"
        self.login("boss")
        plan_response = self.client.post(
            "/api/finance-plans",
            json={
                "car_type": car_type,
                "plan_name": "八月冷藏方案",
                "down_payment": 18000,
                "period_price": 3600,
                "periods": 36,
                "condition": "新车",
                "box_type": "冷藏",
            },
        )
        self.assertEqual(plan_response.status_code, 200, plan_response.get_json())
        plan_id = plan_response.get_json()["id"]

        vehicle_id, vin = self.create_vehicle(car_type)
        conn = database.get_db()
        try:
            vehicle = {
                "car_type": car_type,
                "condition": "新车",
                "box_type": "冷藏",
                "vehicle_box_type": "冷藏",
            }
            validation_status, validation_message = app_module.validate_vehicle_dict(
                conn, car_type, vehicle
            )
            self.assertEqual(validation_status, "warning")
            self.assertIn("仅支持按已维护的以租代售方案报单", validation_message)
            conn.execute(
                """UPDATE vehicles
                   SET box_type='冷藏', vehicle_box_type='冷藏',
                       validation_status=?, validation_message=?
                   WHERE id=?""",
                (validation_status, validation_message, vehicle_id),
            )
            no_plan_status, _ = app_module.validate_vehicle_dict(
                conn,
                "八月无方案车型",
                {
                    "car_type": "八月无方案车型",
                    "condition": "新车",
                    "box_type": "冷藏",
                    "vehicle_box_type": "冷藏",
                },
            )
            self.assertEqual(no_plan_status, "invalid")
            conn.commit()
        finally:
            conn.close()

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-22",
                "customer_name": "仅方案以租代售客户",
                "customer_phone": "13800000023",
                "sales_mode": "以租代售",
                "vin": vin,
                "car_type": car_type,
                "finance_plan_id": plan_id,
                "vehicle_box_type": "冷藏",
                "lease_term": "36期",
                "deposit_amount": 18000,
                "vehicle_rent_amount": 3600,
                "customer_screenshot_path": "/uploads/plan-only-first-payment.jpg",
                "first_payment_received_amount": 18000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())

    def test_t3_due_soon_reminder_text_generated(self):
        """流程图 4.1：T-3 还款日前 3 天，系统生成黄色预警提醒文案（幂等）。"""
        vehicle_id, contract_id, _ = self.create_active_rental_contract()
        due_date = (datetime.now().date() + timedelta(days=2)).strftime("%Y-%m-%d")
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO repayments (contract_id, period, due_date, amount, status)
                VALUES (?, 1, ?, 3000, '待还款')
                """,
                (contract_id, due_date),
            )
            repayment_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
        finally:
            conn.close()

        app_module.run_daily_collect(force=True)
        self.assertEqual(
            self.db_value("SELECT status FROM repayments WHERE id=?", (repayment_id,)),
            "临近还款",
        )
        log = self.db_row(
            "SELECT detail FROM audit_logs WHERE action='T-3还款提醒' AND target_type='repayment' AND target_id=?",
            (repayment_id,),
        )
        self.assertIsNotNone(log)
        self.assertIn("请您知悉", log["detail"])
        self.assertIn(due_date, log["detail"])

        # 幂等：重跑不重复生成提醒
        app_module.run_daily_collect(force=True)
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM audit_logs WHERE action='T-3还款提醒' AND target_type='repayment' AND target_id=?",
                (repayment_id,),
            ),
            1,
        )

        # 账单接口为「临近还款」账单附带提醒文案，供前端 T-3 黄色预警展示
        self.login("ops")
        bills = self.client.get("/api/bills/pending").get_json()
        bill = next((b for b in bills if b["id"] == repayment_id), None)
        self.assertIsNotNone(bill)
        self.assertEqual(bill["status"], "临近还款")
        self.assertIn("请您知悉", bill["t3_reminder_text"])

    def test_collection_tasks_require_evidence_before_lock_request(self):
        vehicle_id, contract_id, _ = self.create_active_rental_contract()
        due_date = (datetime.now().date() - timedelta(days=7)).strftime("%Y-%m-%d")
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO repayments (contract_id, period, due_date, amount, status)
                VALUES (?, 1, ?, 3000, '待还款')
                """,
                (contract_id, due_date),
            )
            repayment_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
        finally:
            conn.close()

        app_module.run_daily_collect(force=True)
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM urge_records WHERE repayment_id=? AND urge_day IN (3, 7)",
                (repayment_id,),
            ),
            2,
        )
        self.login("sales")
        blocked = self.client.post(
            "/api/lock-requests",
            json={"repayment_id": repayment_id, "reason": "八月验收锁车"},
        )
        self.assertEqual(blocked.status_code, 400, blocked.get_json())

        self.login("ops")
        t3 = self.client.post(
            f"/api/repayments/{repayment_id}/urge",
            json={"evidence_path": "/uploads/t3-evidence.jpg", "result": "已联系"},
        )
        self.assertEqual(t3.status_code, 200, t3.get_json())
        self.login("sales")
        t7 = self.client.post(
            f"/api/repayments/{repayment_id}/urge",
            json={"evidence_path": "/uploads/t7-evidence.jpg", "result": "已联系"},
        )
        self.assertEqual(t7.status_code, 200, t7.get_json())
        lock = self.client.post(
            "/api/lock-requests",
            json={"repayment_id": repayment_id, "reason": "八月验收锁车"},
        )
        self.assertEqual(lock.status_code, 200, lock.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT status FROM lock_requests WHERE repayment_id=?",
                (repayment_id,),
            ),
            "待运营审核",
        )
        self.assertEqual(
            self.db_value(
                "SELECT status FROM vehicles WHERE id=?",
                (vehicle_id,),
            ),
            "租赁中",
        )

    def test_return_reject_resubmit_and_refund_returns_vehicle_to_stock_pool(self):
        vehicle_id, _, _ = self.create_active_rental_contract(needs_repair=True)
        self.login("sales")
        created = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "到期退车", "remark": "初次提交"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        return_id = created.get_json()["id"]

        self.login("fleet")
        empty_fleet = self.client.post(
            f"/api/return-inspections/{return_id}/fleet",
            json={"needs_repair": True, "repair_reason": "轮胎磨损", "tool_triangle": True},
        )
        self.assertEqual(empty_fleet.status_code, 400, empty_fleet.get_json())
        self.assertIn("公里数记录", empty_fleet.get_json()["message"])
        fleet = self.client.post(
            f"/api/return-inspections/{return_id}/fleet",
            json={
                "needs_repair": True, "repair_reason": "轮胎磨损", "tool_triangle": True,
                "mileage": "30000", "body_tire_clean": "已清理", "accident_info": "无出险",
                "insurance_surcharge": "无", "violation_info": "无违章", "etc_info": "已注销",
                "maintenance_info": "正常",
            },
        )
        self.assertEqual(fleet.status_code, 200, fleet.get_json())
        queued_returns = self.client.get("/api/return-inspections")
        self.assertEqual(queued_returns.status_code, 200, queued_returns.get_json())
        queued_row = next(row for row in queued_returns.get_json() if row["id"] == return_id)
        self.assertEqual(queued_row["vehicle_status"], "退车中")
        self.assertEqual(queued_row["repair_queue_status"], "待运营填写")
        self.assertFalse(queued_row["repair_queue_actionable"])
        self.login("ops")
        operator = self.client.post(
            f"/api/return-inspections/{return_id}/operator",
            json={
                "rent_late_fee": 0, "return_late_fee": 0, "deposit_rent_receivable": 0,
                "deposit_paid": 2000, "total_deduction": 500, "actual_refund": 1500,
            },
        )
        self.assertEqual(operator.status_code, 200, operator.get_json())
        self.login("fin")
        finance = self.client.post(
            f"/api/return-inspections/{return_id}/finance",
            json={"refund_bank_name": "测试银行", "refund_bank_card_no": "6222000012345678"},
        )
        self.assertEqual(finance.status_code, 200, finance.get_json())

        self.login("boss")
        rejected = self.client.post(
            f"/api/return-inspections/{return_id}/boss-reject",
            json={"reason": "请补充退车原因"},
        )
        self.assertEqual(rejected.status_code, 200, rejected.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM return_inspections WHERE id=?", (return_id,)),
            "已驳回待销售修改",
        )

        self.login("sales")
        edited = self.client.put(
            f"/api/return-inspections/{return_id}",
            json={"return_reason": "客户提前退车", "remark": "已补充退车说明"},
        )
        self.assertEqual(edited.status_code, 200, edited.get_json())
        resubmitted = self.client.post(
            f"/api/return-inspections/{return_id}/resubmit",
            json={"resubmit_note": "已补充客户提前退车说明"},
        )
        self.assertEqual(resubmitted.status_code, 200, resubmitted.get_json())

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
                json={"refund_bank_name": "测试银行", "refund_bank_card_no": "6222000012345678"},
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
            json={"refund_serial": "REFUND202608001", "refund_paid_amount": 1500},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "待维修",
        )
        self.assertEqual(
            self.db_value("SELECT condition FROM vehicles WHERE id=?", (vehicle_id,)),
            "二手车",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM return_inspections WHERE id=?", (return_id,)),
            "已完成",
        )
        history = self.client.get("/api/completion-history?source_type=rental_return")
        self.assertEqual(history.status_code, 200, history.get_json())
        record = next((row for row in history.get_json() if row["source_id"] == return_id), None)
        self.assertIsNotNone(record)
        self.assertEqual(record["completion_type"], "租赁退车入库")
        self.assertIn("二手车", record["completion_result"])
        self.assertEqual(record["amount_value"], 1500)
        self.assertEqual(record["bank_serial"], "REFUND202608001")
        self.assertEqual(record["handled_by"], "张财务")


if __name__ == "__main__":
    unittest.main()
