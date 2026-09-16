import os
import shutil
import tempfile
import unittest
from datetime import timedelta

from openpyxl import Workbook

import app as app_module
import database


class LegacyContractMigrationTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-legacy-migration-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "legacy-migration.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.today = app_module.datetime.now().date()
        self.cutover_date = (self.today + timedelta(days=1)).strftime("%Y-%m-%d")
        self.workbook_path = os.path.join(self.temp_dir, "legacy.xlsx")
        self.normal_vin = "LGCYTESTVIN000001"
        self.external_vin = "LGCYTESTVIN000002"
        self.renewal_vin = "LGCYTESTVIN000003"
        self.correction_vin = "LFNA4LJA5TAE03127"
        self._seed_vehicles()
        self._write_workbook()

    def tearDown(self):
        database.DATABASE = self.original_database
        shutil.rmtree(self.temp_dir)

    def _seed_vehicles(self):
        conn = database.get_db()
        try:
            for vin in (
                self.normal_vin,
                self.external_vin,
                self.renewal_vin,
                self.correction_vin,
            ):
                conn.execute(
                    """
                    INSERT INTO vehicles (vin, plate_number, car_type, condition, status)
                    VALUES (?, ?, '历史迁移测试车型', '新车', '在库')
                    """,
                    (vin, f"陕A{vin[-5:]}"),
                )
            correction_vehicle_id = conn.execute(
                "SELECT id FROM vehicles WHERE vin=?", (self.correction_vin,)
            ).fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, contract_type, contract_status, delivery_status, contract_file)
                VALUES (?, '租赁', '已结清', '已出库', '/uploads/settled-contract.pdf')
                """,
                (correction_vehicle_id,),
            )
            conn.execute(
                "INSERT INTO users (username, password, display_name, role) VALUES (?, ?, ?, ?)",
                ("legacy_other_sales", "123456", "另一销售", "销售"),
            )
            conn.commit()
        finally:
            conn.close()

    def _write_workbook(self):
        book = Workbook()
        master = book.active
        master.title = "销售台账（租赁）"
        headers = [
            "车辆状态", "合同编号", "车架号", "合同开始日", "合同到期日",
            "购买人/租赁人", "联系方式", "销售", "车辆上户公司", "租赁方式",
            "还款时间", "首付款/押金", "月租金额", "租期", "佣金（内勤）",
        ]
        master.append(["历史履约迁移测试"])
        master.append(headers)

        tomorrow = self.today + timedelta(days=1)
        old_start = self.today - timedelta(days=760)
        old_end = self.today - timedelta(days=400)
        renewal_start = old_end + timedelta(days=1)
        renewal_end = self.today + timedelta(days=360)
        master.append([
            "在租", "NORMAL-001", self.normal_vin,
            self.today - timedelta(days=30), self.today + timedelta(days=330),
            "正常客户", "13800000001", "周销售", "金聚源", "经营租赁",
            "每月15号", 5000, 3000, 12, 120,
        ])
        master.append([
            "在租（一汽还款）", "EXTERNAL-001", self.external_vin,
            self.today - timedelta(days=30), self.today + timedelta(days=330),
            "直还客户", "13800000002", "另一销售", "金聚源", "经营租赁",
            "每月15号", 0, 2800, 12, 0,
        ])
        master.append([
            "续租", "RENEW-OLD", self.renewal_vin, old_start, old_end,
            "续租客户", "13800000003", "周销售", "金聚源", "经营租赁",
            "每月15号", 0, 2600, 12, 0,
        ])
        master.append([
            "在租", "RENEW-NEW", self.renewal_vin, renewal_start, renewal_end,
            "续租客户", "13800000003", "周销售", "金聚源", "经营租赁",
            "每月15号", 0, 2700, 12, 0,
        ])
        master.append([
            "在租", "CORRECTION-001", self.correction_vin,
            self.today - timedelta(days=30), self.today + timedelta(days=330),
            "修正客户", "13800000004", "周销售", "金聚源", "经营租赁",
            "每月15号", 0, 3200, 12, 0,
        ])
        master.append([
            "在租", "BAD-VIN", "SHORTVIN",
            self.today - timedelta(days=30), self.today + timedelta(days=330),
            "异常客户", "13800000005", "周销售", "金聚源", "经营租赁",
            "每月15号", 0, 3000, 12, 0,
        ])

        due = book.create_sheet("收款及逾期")
        due.append(["应收账款逾期明细表"])
        due.append([])
        due.append([
            "公司", "客户名称", "车架号", "车牌号", "租赁方式", "租赁期", None,
            "月供/租金", None, "开票", None, "回款", None, "逾期金额", "逾期天数",
            "违约天数", "违约费率", "违约金", "是否收取", "销售",
        ])
        due.append([
            None, None, None, None, None, "租期起", "租期止", "应还款日期", "金额",
            "日期", "发票号码", "回款日期", "回款金额", None, None, None, None, None,
            None, None,
        ])

        def add_plan(vin, customer, lease_start, lease_end, amount, paid, late_fee=0):
            due.append([
                "金聚源", customer, vin, "陕A测试", "经营租赁", lease_start, lease_end,
                lease_start, amount, None, None, self.today, paid, 0, 0, 0, 0.0005,
                late_fee, None, "周销售",
            ])

        add_plan(
            self.normal_vin, "正常客户", self.today - timedelta(days=10),
            self.today + timedelta(days=20), 3000, 1000, 88,
        )
        add_plan(
            self.external_vin, "直还客户", self.today - timedelta(days=10),
            self.today + timedelta(days=20), 2800, 0,
        )
        add_plan(self.renewal_vin, "续租客户", old_start, old_start + timedelta(days=30), 2600, 2600)
        add_plan(
            self.renewal_vin, "续租客户", renewal_start,
            renewal_start + timedelta(days=30), 2700, 0,
        )
        add_plan(
            self.correction_vin, "修正客户", self.today - timedelta(days=10),
            self.today + timedelta(days=20), 3200, 0,
        )
        book.save(self.workbook_path)

    def _legacy_contract(self, contract_number):
        conn = database.get_db()
        try:
            return dict(conn.execute(
                """
                SELECT c.*
                FROM contracts c
                JOIN legacy_contract_migrations m ON m.contract_id=c.id
                WHERE m.original_contract_number=?
                """,
                (contract_number,),
            ).fetchone())
        finally:
            conn.close()

    def test_preview_import_permissions_and_cutover(self):
        preview = app_module.build_legacy_import_preview(
            self.workbook_path,
            self.cutover_date,
        )
        records = {row["contract_number"]: row for row in preview["records"]}
        self.assertEqual(preview["summary"]["total"], 6)
        self.assertEqual(records["BAD-VIN"]["import_status"], "blocked")
        self.assertEqual(records["RENEW-OLD"]["import_status"], "ready")
        self.assertEqual(records["RENEW-NEW"]["import_status"], "ready")
        self.assertEqual(records["NORMAL-001"]["plans"][0]["paid_amount"], 1000)
        self.assertEqual(records["NORMAL-001"]["plans"][0]["historic_late_fee"], 88)
        self.assertEqual(len(records["RENEW-OLD"]["plans"]), 1)
        self.assertEqual(len(records["RENEW-NEW"]["plans"]), 1)
        self.assertTrue(records["CORRECTION-001"]["legacy_correction_old_contract_ids"])

        conn = database.get_db()
        try:
            result = app_module._import_legacy_preview_records(
                conn,
                preview,
                [row["legacy_contract_key"] for row in preview["records"] if row["import_status"] == "ready"],
                "王老板",
                "legacy.xlsx",
                "/uploads/legacy.xlsx",
            )
            conn.commit()
        finally:
            conn.close()
        self.assertEqual(len(result["imported_ids"]), 5)

        normal_contract = self._legacy_contract("NORMAL-001")
        external_contract = self._legacy_contract("EXTERNAL-001")
        old_renewal_contract = self._legacy_contract("RENEW-OLD")
        self.assertEqual(normal_contract["contract_origin"], "legacy")
        self.assertEqual(normal_contract["contract_status"], "执行中")
        self.assertEqual(old_renewal_contract["contract_status"], "已结清")

        conn = database.get_db()
        try:
            normal_plan = conn.execute(
                "SELECT paid_amount, verified_amount, status FROM repayments WHERE contract_id=?",
                (normal_contract["id"],),
            ).fetchone()
            self.assertEqual(normal_plan["paid_amount"], 1000)
            self.assertEqual(normal_plan["verified_amount"], 1000)
            self.assertEqual(normal_plan["status"], "部分核销")
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM repayments WHERE contract_id=?",
                (external_contract["id"],),
            ).fetchone()[0], 0)
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM external_finance_repayments WHERE contract_id=?",
                (external_contract["id"],),
            ).fetchone()[0], 1)
            self.assertEqual(conn.execute(
                "SELECT amount_due FROM contract_fee_items WHERE contract_id=? AND fee_type='legacy_late_fee'",
                (normal_contract["id"],),
            ).fetchone()[0], 88)
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM legacy_contract_corrections WHERE replacement_contract_id=?",
                (self._legacy_contract("CORRECTION-001")["id"],),
            ).fetchone()[0], 1)
            self.assertEqual(conn.execute(
                """
                SELECT COUNT(*)
                FROM contract_role_assignments cra
                JOIN users u ON u.id=cra.user_id
                WHERE cra.contract_id=? AND u.display_name='周销售'
                """,
                (normal_contract["id"],),
            ).fetchone()[0], 1)
        finally:
            conn.close()

        app_module._run_daily_collect(force=True)
        conn = database.get_db()
        try:
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM late_fee_ledger WHERE contract_id=?",
                (normal_contract["id"],),
            ).fetchone()[0], 0)
            conn.execute(
                "UPDATE contracts SET late_fee_accrual_start_date=? WHERE id=?",
                (self.today.strftime("%Y-%m-%d"), normal_contract["id"]),
            )
            conn.commit()
        finally:
            conn.close()
        app_module._run_daily_collect(force=True)
        conn = database.get_db()
        try:
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM late_fee_ledger WHERE contract_id=?",
                (normal_contract["id"],),
            ).fetchone()[0], 1)
        finally:
            conn.close()

        # 导入后为“待确认”：运营确认前全端不可见
        self.assertEqual(normal_contract["legacy_visibility_status"], "待确认")

        self.client.post("/api/auth/login", json={"username": "boss", "password": "123456"})
        boss_contracts = self.client.get("/api/contracts").get_json()
        self.assertNotIn(normal_contract["id"], {row.get("id") for row in boss_contracts})

        self.client.post("/api/auth/login", json={"username": "sales", "password": "123456"})
        sales_contracts = self.client.get("/api/contracts").get_json()
        self.assertNotIn(normal_contract["id"], {row.get("id") for row in sales_contracts})
        denied = self.client.get("/api/legacy-import/pending-contracts")
        self.assertEqual(denied.status_code, 403)

        self.client.post("/api/auth/login", json={"username": "fin", "password": "123456"})
        fin_hidden = self.client.get(
            f"/api/contracts/{normal_contract['id']}/external-finance-repayments"
        )
        self.assertEqual(fin_hidden.status_code, 404)

        # 运营在待确认列表处理；确认后全端可见
        self.client.post("/api/auth/login", json={"username": "ops", "password": "123456"})
        pending = self.client.get("/api/legacy-import/pending-contracts").get_json()
        self.assertEqual(pending["count"], 5)
        self.assertIn(normal_contract["id"], {row["id"] for row in pending["items"]})
        confirm = self.client.post(
            "/api/legacy-import/pending-contracts/confirm",
            json={"confirm_all": True},
        )
        self.assertEqual(confirm.status_code, 200)
        self.assertEqual(confirm.get_json()["confirmed_count"], 5)
        pending_after = self.client.get("/api/legacy-import/pending-contracts").get_json()
        self.assertEqual(pending_after["count"], 0)

        self.client.post("/api/auth/login", json={"username": "sales", "password": "123456"})
        sales_contracts = self.client.get("/api/contracts").get_json()
        sales_numbers = {row.get("id") for row in sales_contracts}
        self.assertIn(normal_contract["id"], sales_numbers)
        self.assertNotIn(external_contract["id"], sales_numbers)
        hidden_external = self.client.get(
            f"/api/contracts/{external_contract['id']}/external-finance-repayments"
        )
        self.assertEqual(hidden_external.status_code, 404)

        conn = database.get_db()
        try:
            duplicate_result = app_module._import_legacy_preview_records(
                conn,
                preview,
                [row["legacy_contract_key"] for row in preview["records"] if row["import_status"] == "ready"],
                "王老板",
                "legacy.xlsx",
                "/uploads/legacy.xlsx",
            )
            conn.commit()
            self.assertEqual(len(duplicate_result["imported_ids"]), 0)
            self.assertEqual(duplicate_result["skipped"], 5)
            self.assertEqual(conn.execute(
                "SELECT COUNT(*) FROM legacy_contract_migrations"
            ).fetchone()[0], 5)
        finally:
            conn.close()


if __name__ == "__main__":
    unittest.main()
