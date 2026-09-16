import os
import shutil
import tempfile
import unittest
from datetime import date, timedelta

import app as app_module
import database


class GuidanceBoardTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-guidance-board-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "guidance-board.db")
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

    def test_board_uses_latest_price_and_current_effective_plans(self):
        self.login("boss")
        car_type = "指导价看板测试车型 / 厢货"

        for deposit, monthly in ((2000, 3000), (2600, 3600)):
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
            "/api/finance-plans",
            json={
                "car_type": car_type,
                "condition": "新车",
                "box_type": "厢货",
                "plan_name": "当前有效方案",
                "down_payment": 12000,
                "period_price": 3500,
                "periods": 24,
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

        conn = database.get_db()
        try:
            today = date.today()
            conn.execute(
                """
                INSERT INTO finance_plans
                    (car_type, condition, box_type, plan_name, down_payment, period_price,
                     periods, status, effective_date, updated_by, updated_at)
                VALUES (?, '新车', '厢货', '已停用方案', 10000, 3000, 24, '停用', ?, '测试', ?)
                """,
                (app_module.normalize_base_car_type(car_type), today.isoformat(), f"{today} 10:00:00"),
            )
            conn.execute(
                """
                INSERT INTO finance_plans
                    (car_type, condition, box_type, plan_name, down_payment, period_price,
                     periods, status, effective_date, updated_by, updated_at)
                VALUES (?, '新车', '厢货', '未来生效方案', 10000, 3000, 24, '启用', ?, '测试', ?)
                """,
                (
                    app_module.normalize_base_car_type(car_type),
                    (today + timedelta(days=1)).isoformat(),
                    f"{today} 10:00:00",
                ),
            )
            conn.commit()
        finally:
            conn.close()

        board = self.client.get("/api/guidance-board")
        self.assertEqual(board.status_code, 200, board.get_json())
        data = board.get_json()
        lease = next(row for row in data["lease_prices"] if row["car_type"] == app_module.normalize_base_car_type(car_type))
        self.assertEqual(lease["lease_deposit_guidance"], 2600)
        self.assertEqual(lease["box_standard_price"], 3600)
        self.assertEqual([row["plan_name"] for row in data["finance_plans"]], ["当前有效方案"])

    def test_board_is_visible_to_every_role(self):
        for username in ("boss", "ops", "fin", "fleet", "sales"):
            user = self.login(username)
            self.assertIn("guidance_board", user["pages"])
            response = self.client.get("/api/guidance-board")
            self.assertEqual(response.status_code, 200, (username, response.get_json()))


if __name__ == "__main__":
    unittest.main()
