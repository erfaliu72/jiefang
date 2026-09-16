import importlib.util
import os
import shutil
import tempfile
import unittest
from unittest.mock import patch

import app as app_module
import database


class _Cursor:
    def __init__(self, rows):
        self.rows = rows
        self.statements = []

    def execute(self, sql):
        self.statements.append(sql)
        return self

    def fetchall(self):
        return self.rows


def load_mysql_database_module():
    module_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "database.py",
    )
    spec = importlib.util.spec_from_file_location(
        "database_model_guidance_price_mysql_migration_test",
        module_path,
    )
    module = importlib.util.module_from_spec(spec)
    with patch.dict(os.environ, {"JJY_DB_HOST": "mysql-test-host"}, clear=False):
        spec.loader.exec_module(module)
    return module


class ModelGuidancePriceMysqlMigrationTestCase(unittest.TestCase):
    def test_removes_only_the_legacy_single_column_unique_index(self):
        database_module = load_mysql_database_module()
        cursor = _Cursor([
            {
                "Key_name": "PRIMARY",
                "Non_unique": 0,
                "Seq_in_index": 1,
                "Column_name": "id",
            },
            {
                "Key_name": "car_type",
                "Non_unique": 0,
                "Seq_in_index": 1,
                "Column_name": "car_type",
            },
            {
                "Key_name": "idx_model_guidance_prices_car_type_is_new",
                "Non_unique": 0,
                "Seq_in_index": 1,
                "Column_name": "car_type",
            },
            {
                "Key_name": "idx_model_guidance_prices_car_type_is_new",
                "Non_unique": 0,
                "Seq_in_index": 2,
                "Column_name": "is_new",
            },
            {
                "Key_name": "idx_model_guidance_prices_car_type",
                "Non_unique": 1,
                "Seq_in_index": 1,
                "Column_name": "car_type",
            },
        ])

        dropped = database_module.drop_legacy_model_guidance_price_unique_indexes(
            cursor
        )

        self.assertEqual(dropped, ["car_type"])
        self.assertEqual(
            cursor.statements,
            [
                "SHOW INDEX FROM model_guidance_prices",
                "DROP INDEX `car_type` ON `model_guidance_prices`",
            ],
        )


class ModelGuidancePriceScopeTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(
            prefix="jinjuyuan-model-guidance-price-scope-"
        )
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "test.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        response = self.client.post(
            "/api/auth/login",
            json={"username": "boss", "password": "123456"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def test_new_and_used_prices_for_the_same_model_do_not_overwrite_each_other(self):
        car_type = "解放虎6G宁德100度无尾板"
        for condition, monthly_price in (("新车", 3900), ("二手车", 4300)):
            response = self.client.post(
                "/api/model-guidance-prices",
                json={
                    "car_type": car_type,
                    "is_new": condition,
                    "box_high_rail_price": monthly_price,
                },
            )
            self.assertEqual(response.status_code, 200, response.get_json())

        conn = database.get_db()
        try:
            rows = conn.execute(
                """
                SELECT is_new, box_high_rail_price
                FROM model_guidance_prices
                WHERE car_type=?
                ORDER BY is_new
                """,
                (car_type,),
            ).fetchall()
            history_rows = conn.execute(
                """
                SELECT is_new
                FROM model_guidance_price_history
                WHERE car_type=?
                ORDER BY id
                """,
                (car_type,),
            ).fetchall()
        finally:
            conn.close()

        self.assertEqual(
            [(row["is_new"], row["box_high_rail_price"]) for row in rows],
            [("二手车", 4300), ("新车", 3900)],
        )
        self.assertEqual(
            [row["is_new"] for row in history_rows],
            ["新车", "二手车"],
        )


if __name__ == "__main__":
    unittest.main()
