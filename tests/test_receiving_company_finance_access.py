import os
import shutil
import tempfile
import unittest

import app as app_module
import database


class ReceivingCompanyFinanceAccessTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-receiving-company-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "receiving-company.db")
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

    def test_finance_has_receiving_company_page_and_can_maintain_data(self):
        finance_user = self.login("fin")
        self.assertIn("receiving_companies", finance_user["pages"])

        created = self.client.post(
            "/api/receiving-companies",
            json={
                "company_name": "财务维护测试公司",
                "bank_name": "测试银行",
                "bank_account_no": "622202608250001",
                "tax_no": "91310000TEST00001",
                "status": "启用",
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        self.assertTrue(created.get_json()["success"])

        companies = self.client.get("/api/receiving-companies?include_disabled=1")
        self.assertEqual(companies.status_code, 200, companies.get_json())
        company = next(
            row
            for row in companies.get_json()
            if row["company_name"] == "财务维护测试公司"
        )
        self.assertEqual(company["updated_by"], "张财务")
        self.assertEqual(company["bank_account_no"], "622202608250001")

    def test_boss_keeps_receiving_company_page_access(self):
        boss_user = self.login("boss")
        self.assertIn("receiving_companies", boss_user["pages"])


if __name__ == "__main__":
    unittest.main()
