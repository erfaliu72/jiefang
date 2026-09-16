import importlib.util
import os
import unittest
from unittest.mock import patch


class _RawCursor:
    def __init__(self):
        self.calls = []
        self.lastrowid = 0
        self.rowcount = 0

    def executemany(self, sql, params):
        self.calls.append((sql, params))
        self.rowcount = len(params)

    def execute(self, sql, params=None):
        self.calls.append((sql, params))


def load_mysql_database_module():
    module_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "database.py",
    )
    spec = importlib.util.spec_from_file_location("database_mysql_adapter_test", module_path)
    module = importlib.util.module_from_spec(spec)
    with patch.dict(os.environ, {"JJY_DB_HOST": "mysql-test-host"}, clear=False):
        spec.loader.exec_module(module)
    return module


class MysqlCursorAdapterTestCase(unittest.TestCase):
    def test_executemany_translates_sqlite_placeholders_and_reserved_condition(self):
        database_module = load_mysql_database_module()
        raw = _RawCursor()
        cursor = database_module._Cursor(raw, None)

        returned = cursor.executemany(
            "INSERT INTO finance_plans (condition, plan_name) VALUES (?, ?)",
            [("新车", "方案A"), ("二手车", "方案B")],
        )

        self.assertIs(returned, cursor)
        self.assertEqual(
            raw.calls,
            [(
                "INSERT INTO finance_plans (`condition`, plan_name) VALUES (%s, %s)",
                [("新车", "方案A"), ("二手车", "方案B")],
            )],
        )

    def test_executemany_ignores_empty_batches(self):
        database_module = load_mysql_database_module()
        raw = _RawCursor()
        cursor = database_module._Cursor(raw, None)

        returned = cursor.executemany("INSERT INTO sample (name) VALUES (?)", [])

        self.assertIs(returned, cursor)
        self.assertEqual(raw.calls, [])

    def test_translates_sale_payment_seed_concatenation(self):
        database_module = load_mysql_database_module()
        raw = _RawCursor()
        cursor = database_module._Cursor(raw, None)

        cursor.execute(
            "SELECT 'sale_payment_' || so.id || '_seed' FROM sales_orders so"
        )

        self.assertEqual(
            raw.calls,
            [("SELECT CONCAT('sale_payment_', so.id, '_seed') FROM sales_orders so", None)],
        )

    def test_translates_return_settlement_remark_concatenation(self):
        database_module = load_mysql_database_module()
        raw = _RawCursor()
        cursor = database_module._Cursor(raw, None)

        cursor.execute("UPDATE repayments SET remark=remark || '；' || ? WHERE id=?", ("note", 1))

        self.assertEqual(
            raw.calls,
            [(
                "UPDATE repayments SET remark=CONCAT(remark, '；', %s) WHERE id=%s",
                ("note", 1),
            )],
        )


if __name__ == "__main__":
    unittest.main()
