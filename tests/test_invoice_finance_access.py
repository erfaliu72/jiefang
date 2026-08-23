import os
import shutil
import tempfile
import unittest

import app as app_module
import database


class InvoiceFinanceAccessTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-invoice-finance-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "invoice-finance.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def login(self, username):
        response = self.client.post(
            "/api/auth/login",
            json={"username": username, "password": "123456"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()["user"]

    def create_invoice_request(self):
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles (vin, plate_number, car_type, condition, status)
                VALUES ('INVOICEFINANCE0001', '陕A票测01', '发票测试车型', '新车', '在库')
                """
            )
            vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                "INSERT INTO contracts (vehicle_id, contract_type) VALUES (?, '租赁')",
                (vehicle_id,),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO invoice_requests
                    (contract_id, amount, invoice_entity_name, status, applied_by)
                VALUES (?, 1200, '发票测试公司', '待审批', '李运营')
                """,
                (contract_id,),
            )
            invoice_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            app_module.create_approval_flow(conn, "invoice", invoice_id)
            flow_id = conn.execute(
                "SELECT id FROM approval_flows WHERE ref_type='invoice' AND ref_id=?",
                (invoice_id,),
            ).fetchone()[0]
            conn.commit()
            return invoice_id, flow_id
        finally:
            conn.close()

    def test_finance_handles_boss_approved_invoice_in_approval_center(self):
        invoice_id, flow_id = self.create_invoice_request()

        self.login("boss")
        approval = self.client.post(f"/api/approvals/{flow_id}/approve", json={})
        self.assertEqual(approval.status_code, 200, approval.get_json())
        self.assertEqual(approval.get_json()["message"], "发票审批通过，等待财务开票")

        finance_user = self.login("fin")
        self.assertNotIn("invoice", finance_user["pages"])

        approvals = self.client.get("/api/approvals?ref_type=invoice")
        self.assertEqual(approvals.status_code, 200, approvals.get_json())
        invoice = next(item for item in approvals.get_json() if item["ref_id"] == invoice_id)
        self.assertEqual(invoice["delivery_status"], "待开票")
        self.assertEqual(invoice["follow_up_role"], "财务")

        issued = self.client.post(
            f"/api/invoice-requests/{invoice_id}/issue",
            json={"invoice_no": "FP-20260823-001", "invoice_file_path": "/uploads/test-invoice.pdf"},
        )
        self.assertEqual(issued.status_code, 200, issued.get_json())
        self.assertTrue(issued.get_json()["success"])

        approvals = self.client.get("/api/approvals?ref_type=invoice")
        self.assertEqual(approvals.status_code, 200, approvals.get_json())
        invoice = next(item for item in approvals.get_json() if item["ref_id"] == invoice_id)
        self.assertEqual(invoice["delivery_status"], "已开票")


if __name__ == "__main__":
    unittest.main()
