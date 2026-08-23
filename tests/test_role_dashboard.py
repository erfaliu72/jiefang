import os
import shutil
import tempfile
import unittest
from datetime import timedelta

import app as app_module
import database


class RoleDashboardTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-role-dashboard-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "role-dashboard.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.today = app_module.datetime.now().date()
        self.month_date = self.today.replace(day=10).strftime("%Y-%m-%d")
        self.created_at = f"{self.month_date} 10:00:00"
        self._seed_role_scoped_contracts()

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def login(self, username):
        response = self.client.post(
            "/api/auth/login",
            json={"username": username, "password": "123456"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def _seed_role_scoped_contracts(self):
        next_month_due = (self.today + timedelta(days=1)).strftime("%Y-%m-%d")
        conn = database.get_db()
        try:
            conn.execute(
                "INSERT INTO users (username, password, display_name, role) VALUES (?, ?, ?, ?)",
                ("sales_other", "123456", "另一销售", "销售"),
            )
            self._create_contract(
                conn,
                vin="DASHBOARDOWN00001",
                customer_name="仪表盘本人客户",
                salesperson="周销售",
                due_amount=1200,
                due_date=self.month_date,
            )
            self._create_contract(
                conn,
                vin="DASHBOARDOTHER0001",
                customer_name="仪表盘他人客户",
                salesperson="另一销售",
                due_amount=8800,
                due_date=self.month_date,
            )
            conn.commit()
        finally:
            conn.close()

    def _create_contract(self, conn, vin, customer_name, salesperson, due_amount, due_date):
        conn.execute(
            """
            INSERT INTO vehicles (vin, plate_number, car_type, condition, status)
            VALUES (?, ?, '仪表盘测试车型', '新车', '租赁中')
            """,
            (vin, f"陕A{vin[-5:]}"),
        )
        vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            """
            INSERT INTO customers (name, phone)
            VALUES (?, '13800000000')
            """,
            (customer_name,),
        )
        customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            """
            INSERT INTO sales_orders
                (customer_name, customer_phone, sales_mode, vehicle_id, vin, car_type,
                 sales_advisor, created_by, order_status, created_at)
            VALUES (?, '13800000000', '租赁', ?, ?, '仪表盘测试车型', ?, ?, '已激活', ?)
            """,
            (customer_name, vehicle_id, vin, salesperson, salesperson, self.created_at),
        )
        order_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            """
            INSERT INTO contracts
                (vehicle_id, customer_id, sales_order_id, contract_type, contract_status,
                 delivery_status, contract_file, created_at)
            VALUES (?, ?, ?, '租赁', '执行中', '已出库', '/uploads/dashboard-contract.pdf', ?)
            """,
            (vehicle_id, customer_id, order_id, self.created_at),
        )
        contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.execute(
            "UPDATE sales_orders SET contract_id=? WHERE id=?",
            (contract_id, order_id),
        )
        conn.execute(
            """
            INSERT INTO repayments (contract_id, period, due_date, amount, status)
            VALUES (?, 1, ?, ?, '待还款')
            """,
            (contract_id, due_date, due_amount),
        )

    def test_sales_workspace_is_scoped_to_own_orders_and_contracts(self):
        self.login("sales")
        response = self.client.get("/api/dashboard/workspace")
        self.assertEqual(response.status_code, 200, response.get_json())
        workspace = response.get_json()

        self.assertEqual(workspace["role"], "销售")
        cards = {card["key"]: card for card in workspace["cards"]}
        self.assertEqual(cards["monthly_orders"]["value"], 1)
        self.assertEqual(cards["active_contracts"]["value"], 1)
        self.assertEqual(cards["monthly_due"]["value"], 1200)
        self.assertNotIn("cash_profit", cards)

        task_text = " ".join(
            f"{item['title']} {item.get('subtitle', '')}"
            for group in workspace["todo_groups"]
            for item in group["items"]
        )
        self.assertNotIn("仪表盘他人客户", task_text)

        legacy_stats = self.client.get("/api/dashboard/stats")
        self.assertEqual(legacy_stats.status_code, 200, legacy_stats.get_json())
        self.assertNotIn("monthly_due", legacy_stats.get_json())
        self.assertNotIn("realized_cash_profit", legacy_stats.get_json())
        self.assertNotIn("customer_received", legacy_stats.get_json())

    def test_finance_and_fleet_receive_only_their_dashboard_dimensions(self):
        self.login("fin")
        finance_workspace = self.client.get("/api/dashboard/workspace").get_json()
        finance_cards = {card["key"]: card for card in finance_workspace["cards"]}
        self.assertEqual(finance_workspace["role"], "财务")
        self.assertEqual(finance_cards["cash_profit"]["value_type"], "money")
        self.assertIn("monthly_factory_paid", finance_cards)

        self.login("fleet")
        fleet_workspace = self.client.get("/api/dashboard/workspace").get_json()
        self.assertEqual(fleet_workspace["role"], "车管")
        self.assertTrue(fleet_workspace["cards"])
        self.assertTrue(all(card["value_type"] == "count" for card in fleet_workspace["cards"]))
        self.assertNotIn("cash_profit", {card["key"] for card in fleet_workspace["cards"]})


if __name__ == "__main__":
    unittest.main()
