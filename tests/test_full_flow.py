import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta

import app as app_module
import database


class FullFlowTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-test-")
        self.original_database = database.DATABASE
        self.original_upload_dir = app_module.UPLOAD_DIR

        database.DATABASE = os.path.join(self.temp_dir, "jinjuyuan-test.db")
        app_module.UPLOAD_DIR = os.path.join(self.temp_dir, "uploads")
        os.makedirs(app_module.UPLOAD_DIR, exist_ok=True)

        database.init_db()
        database.seed_data()

        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.vin_index = 0

    def tearDown(self):
        database.DATABASE = self.original_database
        app_module.UPLOAD_DIR = self.original_upload_dir
        shutil.rmtree(self.temp_dir)

    def login(self, username, password="123456"):
        response = self.client.post(
            "/api/auth/login",
            json={"username": username, "password": password},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()["user"]

    def logout(self):
        response = self.client.post("/api/auth/logout")
        self.assertEqual(response.status_code, 200)

    def db_value(self, sql, params=()):
        conn = database.get_db()
        try:
            row = conn.execute(sql, params).fetchone()
            if row is None:
                return None
            if len(row.keys()) == 1:
                return row[0]
            return dict(row)
        finally:
            conn.close()

    def db_execute(self, sql, params=()):
        conn = database.get_db()
        try:
            conn.execute(sql, params)
            conn.commit()
        finally:
            conn.close()

    def next_vehicle_payload(self):
        self.vin_index += 1
        suffix = f"{self.vin_index:016d}"
        return {
            "vin": f"T{suffix}",
            "plate_number": f"陕测{self.vin_index:04d}",
            "company": "陕西金聚源汽车服务有限公司",
            "car_type": "解放轻卡4米2-领途150马力",
            "is_new": "新车",
            "invoice_date": "2026-05-01",
            "invoice_price": 120000,
            "purchase_price": 118000,
            "tax_rate": 0.13,
            "estimated_residual_value": 90000,
            "guidance_price": 98000,
            "status": "在库",
        }

    def create_vehicle(self):
        self.login("fleet")
        response = self.client.post("/api/vehicles", json=self.next_vehicle_payload())
        self.assertEqual(response.status_code, 200, response.get_json())
        vehicle_id = response.get_json()["id"]
        self.logout()
        return vehicle_id

    def create_contract(self, vehicle_id, contract_type="租赁", business_mode="转租"):
        self.login("sales")
        payload = {
            "vehicle_id": vehicle_id,
            "customer_name": f"测试客户{vehicle_id}",
            "customer_phone": f"1380000{vehicle_id:04d}",
            "contract_type": contract_type,
            "business_mode": business_mode,
            "rental_method": "经营租赁",
            "repayment_day": 15,
            "start_date": "2026-01-01",
            "loan_periods": 12,
            "rent": 3000,
            "monthly_payment": 2500,
            "deposit": 2000,
            "down_payment": 1000,
            "total_price": 120000,
            "loan_amount": 90000,
        }
        if contract_type == "销售":
            payload.update({
                "loan_periods": 0,
                "rent": 0,
                "monthly_payment": 0,
                "deposit": 0,
                "down_payment": 30000,
                "business_mode": "卖车",
            })
        response = self.client.post("/api/contracts", json=payload)
        self.assertEqual(response.status_code, 200, response.get_json())
        contract_id = response.get_json()["id"]
        self.logout()
        return contract_id

    def get_approval_steps(self, ref_type, ref_id):
        conn = database.get_db()
        try:
            latest_batch_row = conn.execute(
                """
                SELECT batch_no
                FROM approval_flows
                WHERE ref_type=? AND ref_id=?
                ORDER BY id DESC
                LIMIT 1
                """,
                (ref_type, ref_id),
            ).fetchone()
            if not latest_batch_row:
                return []
            latest_batch = latest_batch_row["batch_no"]
            return conn.execute(
                """
                SELECT id, step_order, required_role, status
                FROM approval_flows
                WHERE ref_type=? AND ref_id=? AND batch_no=?
                ORDER BY step_order ASC
                """,
                (ref_type, ref_id, latest_batch),
            ).fetchall()
        finally:
            conn.close()

    def approve_latest_flow(self, ref_type, ref_id):
        role_to_user = {"运营": "ops", "法务": "legal", "财务": "fin", "老板": "boss"}
        while True:
            pending_steps = [step for step in self.get_approval_steps(ref_type, ref_id) if step["status"] == "待审批"]
            if not pending_steps:
                break
            step = pending_steps[0]
            self.login(role_to_user[step["required_role"]])
            response = self.client.post(
                f"/api/approvals/{step['id']}/approve",
                json={"comment": f"{step['required_role']}审批通过"},
            )
            self.assertEqual(response.status_code, 200, response.get_json())
            self.logout()

    def initiate_initial_payment(self, contract_id, amount=None):
        self.login("sales")
        payload = {
            "customer_screenshot_path": "/uploads/initial-customer-payment.png",
            "remark": "自动化测试首次付款",
        }
        if amount is not None:
            payload["amount"] = amount
        response = self.client.post(
            f"/api/contracts/{contract_id}/initial-payment",
            json=payload,
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        payment_id = response.get_json()["id"]
        self.logout()
        return payment_id

    def upload_initial_payment_receipt(self, payment_id):
        self.login("fin")
        response = self.client.post(
            f"/api/initial-payments/{payment_id}/receipt",
            json={"bank_receipt_path": "/uploads/initial-bank-receipt.png"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.logout()

    def complete_initial_payment(self, contract_id, amount=None):
        payment_id = self.initiate_initial_payment(contract_id, amount=amount)
        self.upload_initial_payment_receipt(payment_id)
        self.approve_latest_flow("initial_payment", payment_id)
        return payment_id

    def deliver_vehicle(self, vehicle_id):
        self.login("fleet")
        response = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(response.status_code, 200, response.get_json())
        self.logout()

    def submit_return_inspection(self, vehicle_id):
        self.login("sales")
        response = self.client.post(
            "/api/return-inspections",
            json={
                "vehicle_id": vehicle_id,
                "return_reason": "到期退车",
                "sales_advisor": "周销售",
                "remark": "自动化测试退车验收",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        inspection_id = response.get_json()["id"]

        submit_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/submit-approval",
            json={},
        )
        self.assertEqual(submit_response.status_code, 200, submit_response.get_json())
        self.logout()
        return inspection_id

    def create_return_inspection_only(self, vehicle_id):
        self.login("sales")
        response = self.client.post(
            "/api/return-inspections",
            json={
                "vehicle_id": vehicle_id,
                "return_reason": "到期退车",
                "sales_advisor": "周销售",
                "remark": "自动化测试退车验收",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        inspection_id = response.get_json()["id"]
        self.logout()
        return inspection_id

    def test_rental_delivery_reject_resubmit_and_reconciliation_flow(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)

        contract_roles = [step["required_role"] for step in self.get_approval_steps("contract_delivery", contract_id)]
        self.assertEqual(contract_roles, ["运营", "财务", "法务", "老板"])

        first_step = self.get_approval_steps("contract_delivery", contract_id)[0]
        self.login("ops")
        reject_response = self.client.post(
            f"/api/approvals/{first_step['id']}/reject",
            json={"comment": "资料不完整"},
        )
        self.assertEqual(reject_response.status_code, 200, reject_response.get_json())
        self.logout()

        self.login("sales")
        resubmit_response = self.client.post(
            f"/api/approvals/{contract_id}/resubmit",
            json={"ref_type": "contract_delivery"},
        )
        self.assertEqual(resubmit_response.status_code, 200, resubmit_response.get_json())
        self.logout()

        self.approve_latest_flow("contract_delivery", contract_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待首付款",
        )

        self.login("fleet")
        premature_delivery = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(premature_delivery.status_code, 400, premature_delivery.get_json())
        self.logout()

        payment_id = self.initiate_initial_payment(contract_id)
        self.login("fin")
        finance_approve_without_receipt = self.client.post(
            f"/api/approvals/{self.get_approval_steps('initial_payment', payment_id)[0]['id']}/approve",
            json={"comment": "缺少回单时不允许通过"},
        )
        self.assertEqual(finance_approve_without_receipt.status_code, 400, finance_approve_without_receipt.get_json())
        self.logout()
        self.upload_initial_payment_receipt(payment_id)
        self.approve_latest_flow("initial_payment", payment_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待出库",
        )

        self.deliver_vehicle(vehicle_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "租赁中",
        )

        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )
        initial_repayment_state = self.db_value(
            "SELECT status, verified_by FROM repayments WHERE id=?",
            (repayment_id,),
        )
        self.assertEqual(initial_repayment_state["status"], "已还款")
        self.assertEqual(initial_repayment_state["verified_by"], "王老板")

        recon_repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1 OFFSET 1",
            (contract_id,),
        )
        factory_repayment_id = self.db_value(
            "SELECT id FROM factory_repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )

        self.login("sales")
        screenshot_response = self.client.post(
            f"/api/reconciliation/{recon_repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/customer-payment.png"},
        )
        self.assertEqual(screenshot_response.status_code, 200, screenshot_response.get_json())
        self.logout()

        self.login("fin")
        receipt_response = self.client.post(
            f"/api/reconciliation/{recon_repayment_id}/receipt",
            json={"bank_receipt_path": "/uploads/bank-receipt.png"},
        )
        verify_response = self.client.post(
            f"/api/reconciliation/{recon_repayment_id}/verify",
            json={"bank_serial": "5678"},
        )
        factory_confirm_response = self.client.post(
            f"/api/factory-repayments/{factory_repayment_id}/confirm",
            json={},
        )
        self.assertEqual(receipt_response.status_code, 200, receipt_response.get_json())
        self.assertEqual(verify_response.status_code, 200, verify_response.get_json())
        self.assertEqual(factory_confirm_response.status_code, 200, factory_confirm_response.get_json())

        verify_repeat_response = self.client.post(
            f"/api/reconciliation/{recon_repayment_id}/verify",
            json={"bank_serial": "5678"},
        )
        self.assertEqual(verify_repeat_response.status_code, 400, verify_repeat_response.get_json())

        factory_repeat_response = self.client.post(
            f"/api/factory-repayments/{factory_repayment_id}/confirm",
            json={},
        )
        self.assertEqual(factory_repeat_response.status_code, 400, factory_repeat_response.get_json())
        self.logout()

        contract_state = self.db_value(
            """
            SELECT collected_rent, collected_deposit, deposit_status, down_payment_status, paid_principal
            FROM contracts
            WHERE id=?
            """,
            (contract_id,),
        )
        reconciliation_state = self.db_value(
            "SELECT verified_by, status FROM repayments WHERE id=?",
            (recon_repayment_id,),
        )
        self.assertEqual(contract_state["deposit_status"], "已收")
        self.assertEqual(contract_state["down_payment_status"], "已收")
        self.assertEqual(contract_state["collected_deposit"], 2000.0)
        self.assertEqual(contract_state["collected_rent"], 6000.0)
        self.assertEqual(contract_state["paid_principal"], 2500.0)
        self.assertEqual(reconciliation_state["verified_by"], "张财务")
        self.assertEqual(reconciliation_state["status"], "已还款")

        third_repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1 OFFSET 2",
            (contract_id,),
        )
        self.login("fin")
        direct_confirm_response = self.client.post(f"/api/repayments/{third_repayment_id}/confirm", json={})
        direct_confirm_repeat_response = self.client.post(f"/api/repayments/{third_repayment_id}/confirm", json={})
        self.assertEqual(direct_confirm_response.status_code, 200, direct_confirm_response.get_json())
        self.assertEqual(direct_confirm_repeat_response.status_code, 400, direct_confirm_repeat_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT collected_rent FROM contracts WHERE id=?", (contract_id,)),
            9000.0,
        )

    def test_sale_contract_payment_check_flow(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id, contract_type="销售", business_mode="卖车")

        self.approve_latest_flow("contract_delivery", contract_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待首付款",
        )

        self.assertEqual(self.get_approval_steps("sale_payment", contract_id), [])

        payment_id = self.complete_initial_payment(contract_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待出库",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM contract_initial_payments WHERE id=?", (payment_id,)),
            "已通过",
        )

        self.deliver_vehicle(vehicle_id)

        contract_status = self.db_value(
            "SELECT contract_status FROM contracts WHERE id=?",
            (contract_id,),
        )
        vehicle_status = self.db_value(
            "SELECT status FROM vehicles WHERE id=?",
            (vehicle_id,),
        )
        self.assertEqual(contract_status, "已结清")
        self.assertEqual(vehicle_status, "已售")

    def test_lease_to_sale_contract_requires_initial_payment_before_delivery(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id, contract_type="以租代售", business_mode="以租代售")

        self.approve_latest_flow("contract_delivery", contract_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待首付款",
        )

        self.login("fleet")
        blocked_delivery = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(blocked_delivery.status_code, 400, blocked_delivery.get_json())
        self.logout()

        self.complete_initial_payment(contract_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待出库",
        )

        self.deliver_vehicle(vehicle_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "以租代售",
        )

    def test_overdue_urge_lock_and_return_stock_flow(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        self.approve_latest_flow("contract_delivery", contract_id)
        self.complete_initial_payment(contract_id)
        self.deliver_vehicle(vehicle_id)

        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )
        overdue_date = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
        self.db_execute(
            "UPDATE repayments SET due_date=?, status='待还款' WHERE id=?",
            (overdue_date, repayment_id),
        )

        overdue_response = self.client.get("/api/risk/overdue")
        self.assertEqual(overdue_response.status_code, 200)
        self.assertIn(
            repayment_id,
            [row["id"] for row in overdue_response.get_json()],
        )

        self.login("ops")
        urge_ops_response = self.client.post(f"/api/repayments/{repayment_id}/urge", json={})
        self.assertEqual(urge_ops_response.status_code, 200, urge_ops_response.get_json())
        self.logout()

        self.login("sales")
        urge_sales_response = self.client.post(f"/api/repayments/{repayment_id}/urge", json={})
        lock_response = self.client.post(
            "/api/lock-requests",
            json={
                "vehicle_id": vehicle_id,
                "contract_id": contract_id,
                "repayment_id": repayment_id,
                "overdue_days": 10,
                "reason": "测试锁车流程",
            },
        )
        self.assertEqual(urge_sales_response.status_code, 200, urge_sales_response.get_json())
        self.assertEqual(lock_response.status_code, 200, lock_response.get_json())
        self.logout()

        lock_request_id = self.db_value("SELECT MAX(id) FROM lock_requests")
        self.login("fin")
        approval_types = {
            row["ref_type"]
            for row in self.client.get("/api/approvals").get_json()
        }
        self.logout()
        self.assertIn("lock_request", approval_types)
        self.approve_latest_flow("lock_request", lock_request_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "已锁车",
        )

        inspection_id = self.create_return_inspection_only(vehicle_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "已锁车",
        )
        self.assertEqual(
            self.db_value("SELECT contract_status FROM contracts WHERE id=?", (contract_id,)),
            "执行中",
        )
        self.login("sales")
        submit_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/submit-approval",
            json={},
        )
        self.assertEqual(submit_response.status_code, 200, submit_response.get_json())
        self.logout()

        self.login("ops")
        approval_types = {
            row["ref_type"]
            for row in self.client.get("/api/approvals").get_json()
        }
        self.logout()
        self.assertIn("return_stock", approval_types)
        self.approve_latest_flow("return_stock", inspection_id)

        self.login("fleet")
        execute_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/execute-stock",
            json={},
        )
        self.assertEqual(execute_response.status_code, 200, execute_response.get_json())
        self.logout()

        inspection_status = self.db_value(
            "SELECT status FROM return_inspections WHERE id=?",
            (inspection_id,),
        )
        vehicle_state = self.db_value(
            "SELECT status, is_new FROM vehicles WHERE id=?",
            (vehicle_id,),
        )
        contract_status = self.db_value(
            "SELECT contract_status FROM contracts WHERE id=?",
            (contract_id,),
        )

        self.assertEqual(inspection_status, "已入库")
        self.assertEqual(vehicle_state["status"], "在库")
        self.assertEqual(vehicle_state["is_new"], "二手车")
        self.assertEqual(contract_status, "已结清")

    def test_sensitive_endpoints_require_roles(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )

        unauth_screenshot = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/x.png"},
        )
        self.assertEqual(unauth_screenshot.status_code, 401, unauth_screenshot.get_json())

        self.login("fin")
        finance_screenshot = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/x.png"},
        )
        self.assertEqual(finance_screenshot.status_code, 403, finance_screenshot.get_json())
        self.logout()

        self.login("sales")
        sales_receipt = self.client.post(
            f"/api/reconciliation/{repayment_id}/receipt",
            json={"bank_receipt_path": "/uploads/y.png"},
        )
        self.assertEqual(sales_receipt.status_code, 403, sales_receipt.get_json())
        self.logout()

        self.login("fin")
        finance_lock_request = self.client.post(
            "/api/lock-requests",
            json={"repayment_id": repayment_id, "reason": "越权测试"},
        )
        self.assertEqual(finance_lock_request.status_code, 403, finance_lock_request.get_json())
        self.logout()

        legal_user = self.login("legal")
        self.assertEqual(legal_user["role"], "法务")
        self.assertIn("contracts", legal_user["pages"])
        self.assertIn("approvals", legal_user["pages"])
        self.assertNotIn("reconciliation", legal_user["pages"])
        legal_recon = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/legal-x.png"},
        )
        self.assertEqual(legal_recon.status_code, 403, legal_recon.get_json())
        self.logout()

        self.login("sales")
        sales_me = self.client.get("/api/auth/me")
        self.assertEqual(sales_me.status_code, 200, sales_me.get_json())
        self.assertIn("reconciliation", sales_me.get_json()["user"]["pages"])
        self.logout()


if __name__ == "__main__":
    unittest.main(verbosity=2)
