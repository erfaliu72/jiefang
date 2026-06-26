import os
import shutil
import tempfile
import unittest

import app as app_module
import database


class InitialPaymentRulesTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-payment-rules-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "jinjuyuan-test.db")
        database.init_db()

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def create_contract_shell(self, contract_type):
        vin_kind = "LTS" if contract_type == "以租代售" else "RENT"
        conn = database.get_db()
        try:
            conn.execute("""
                INSERT INTO vehicles (vin, plate_number, car_type, status)
                VALUES (?, ?, ?, '在库')
            """, (f"TPAY{vin_kind}000000000", f"陕付{contract_type}", "测试车型"))
            vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute("""
                INSERT INTO contracts (vehicle_id, contract_type, start_date, repayment_day)
                VALUES (?, ?, '2025-01-20', 20)
            """, (vehicle_id, contract_type))
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return contract_id
        finally:
            conn.close()

    def repayment_rows(self, contract_id):
        conn = database.get_db()
        try:
            return [
                dict(row)
                for row in conn.execute(
                    "SELECT period, due_date, amount, status, remark FROM repayments WHERE contract_id=? ORDER BY period",
                    (contract_id,),
                ).fetchall()
            ]
        finally:
            conn.close()

    def test_lease_to_sale_initial_payment_excludes_first_installment(self):
        contract = {
            "contract_type": "以租代售",
            "down_payment": 20000,
            "deposit": 0,
            "rent": 5000,
        }
        self.assertEqual(app_module.get_initial_payment_amount(contract), 20000)
        self.assertEqual(app_module.initial_payment_label(contract), "首付款审核")

        contract_id = self.create_contract_shell("以租代售")
        conn = database.get_db()
        try:
            app_module.generate_customer_repayment_plan(
                conn,
                contract_id,
                "以租代售",
                "2025-01-20",
                2,
                5000,
                repayment_day=20,
                down_payment=20000,
            )
            conn.commit()
        finally:
            conn.close()

        rows = self.repayment_rows(contract_id)
        self.assertEqual(rows[0]["period"], 0)
        self.assertEqual(rows[0]["amount"], 20000)
        self.assertEqual(rows[0]["remark"], "首付款")
        self.assertEqual(rows[1]["period"], 1)
        self.assertEqual(rows[1]["due_date"], "2025-02-20")
        self.assertEqual(rows[1]["amount"], 5000)

    def test_rental_initial_payment_includes_deposit_and_first_month(self):
        contract = {
            "contract_type": "租赁",
            "down_payment": 0,
            "deposit": 20000,
            "rent": 5000,
        }
        self.assertEqual(app_module.get_initial_payment_amount(contract), 25000)
        self.assertEqual(app_module.initial_payment_label(contract), "押金及首次支付审核")

        contract_id = self.create_contract_shell("租赁")
        conn = database.get_db()
        try:
            app_module.generate_customer_repayment_plan(
                conn,
                contract_id,
                "租赁",
                "2025-01-20",
                2,
                5000,
                repayment_day=20,
                deposit=20000,
            )
            conn.commit()
        finally:
            conn.close()

        rows = self.repayment_rows(contract_id)
        self.assertEqual(rows[0]["period"], 0)
        self.assertEqual(rows[0]["amount"], 20000)
        self.assertEqual(rows[0]["remark"], "押金")
        self.assertEqual(rows[1]["period"], 1)
        self.assertEqual(rows[1]["due_date"], "2025-01-20")
        self.assertEqual(rows[1]["amount"], 5000)

    def test_completed_initial_payment_rows_cannot_be_reconciled_again(self):
        contract_id = self.create_contract_shell("租赁")
        conn = database.get_db()
        try:
            conn.execute("""
                UPDATE contracts
                SET contract_file='/uploads/signed-contract.pdf', delivery_status='已出库'
                WHERE id=?
            """, (contract_id,))
            app_module.generate_customer_repayment_plan(
                conn,
                contract_id,
                "租赁",
                "2025-01-20",
                2,
                5000,
                repayment_day=20,
                deposit=20000,
            )
            conn.execute("""
                UPDATE repayments
                SET status='已还款', paid_amount=amount, verified_amount=amount
                WHERE contract_id=? AND period IN (0, 1)
            """, (contract_id,))
            period0_id = conn.execute(
                "SELECT id FROM repayments WHERE contract_id=? AND period=0",
                (contract_id,),
            ).fetchone()[0]
            period1_id = conn.execute(
                "SELECT id FROM repayments WHERE contract_id=? AND period=1",
                (contract_id,),
            ).fetchone()[0]
            conn.commit()

            _, period0_message, period0_status = app_module.reconciliation_gate_for_repayment(conn, period0_id)
            _, period1_message, period1_status = app_module.reconciliation_gate_for_repayment(conn, period1_id)
        finally:
            conn.close()

        self.assertEqual(period0_status, 400)
        self.assertIn("首次付款", period0_message)
        self.assertEqual(period1_status, 400)
        self.assertIn("已完成对账", period1_message)


if __name__ == "__main__":
    unittest.main()
