import os
import shutil
import tempfile
import unittest
from datetime import datetime, timedelta

from openpyxl import Workbook

import app as app_module
import database


class FullFlowTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jinjuyuan-test-")
        self.original_database = database.DATABASE
        self.original_upload_dir = app_module.UPLOAD_DIR
        self.original_contract_output_dir = app_module.CONTRACT_OUTPUT_DIR

        database.DATABASE = os.path.join(self.temp_dir, "jinjuyuan-test.db")
        app_module.UPLOAD_DIR = os.path.join(self.temp_dir, "uploads")
        app_module.CONTRACT_OUTPUT_DIR = os.path.join(self.temp_dir, "generated-contracts")
        os.makedirs(app_module.UPLOAD_DIR, exist_ok=True)
        os.makedirs(app_module.CONTRACT_OUTPUT_DIR, exist_ok=True)

        database.init_db()
        database.seed_data()

        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.vin_index = 0

    def tearDown(self):
        database.DATABASE = self.original_database
        app_module.UPLOAD_DIR = self.original_upload_dir
        app_module.CONTRACT_OUTPUT_DIR = self.original_contract_output_dir
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

    def next_vehicle_payload(self, overrides=None):
        self.vin_index += 1
        suffix = f"{self.vin_index:016d}"
        payload = {
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
        if overrides:
            payload.update(overrides)
        return payload

    def create_vehicle(self, overrides=None):
        self.login("fleet")
        response = self.client.post("/api/vehicles", json=self.next_vehicle_payload(overrides))
        self.assertEqual(response.status_code, 200, response.get_json())
        vehicle_id = response.get_json()["id"]
        self.logout()
        return vehicle_id

    def get_vehicle(self, vehicle_id):
        return self.db_value(
            "SELECT id, vin, plate_number, status, insurance_expiry_date, annual_review_date FROM vehicles WHERE id=?",
            (vehicle_id,),
        )

    def create_contract(self, vehicle_id, contract_type="租赁", business_mode="转租", overrides=None):
        self.login("ops")
        payload = {
            "vehicle_id": vehicle_id,
            "customer_name": f"测试客户{vehicle_id}",
            "customer_phone": f"1380000{vehicle_id:04d}",
            "company": "陕西金聚源汽车服务有限公司",
            "yard": "西安一号车场",
            "contract_type": contract_type,
            "business_mode": business_mode,
            "rental_method": "经营租赁",
            "repayment_day": 15,
            "start_date": "2026-01-01",
            "loan_periods": 12,
            "customer_loan_amount": 24000,
            "rent": 3000,
            "factory_guarantee_deposit": 1500,
            "factory_repayment_months": 12,
            "factory_periods": 12,
            "monthly_payment": 2500,
            "deposit": 2000,
            "down_payment": 1000,
            "total_price": 120000,
            "loan_amount": 90000,
            "lease_bank_name": "中国建设银行西安分行",
            "lease_bank_card_no": "6227000012345678901",
            "contract_file": "/uploads/signed-contract.pdf",
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
        if overrides:
            payload.update(overrides)
        self.logout()
        self.login("sales")
        forbidden_response = self.client.post("/api/contracts", json=payload)
        self.assertEqual(forbidden_response.status_code, 403, forbidden_response.get_json())
        self.logout()
        self.login("ops")
        response = self.client.post("/api/contracts", json=payload)
        self.assertEqual(response.status_code, 200, response.get_json())
        contract_id = response.get_json()["id"]
        self.logout()
        return contract_id

    def create_factory_plan_xlsx(self, rows):
        path = os.path.join(app_module.UPLOAD_DIR, "factory-plan.xlsx")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["序号", "应还款日期", "客户应还金额合计"])
        for row in rows:
            sheet.append(row)
        workbook.save(path)
        return "/uploads/factory-plan.xlsx"

    def create_factory_plan_pdf(self, rows):
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.pdfgen import canvas

        path = os.path.join(app_module.UPLOAD_DIR, "factory-plan.pdf")
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        c = canvas.Canvas(path)
        c.setFont("STSong-Light", 10)
        y = 800
        lines = [
            "租赁合同号 FCSSSXJJ00528 承租人姓名 张佳新",
            "计息本金(元) 50,000.00 租赁期限(月) 14",
            "主车车牌号 陕U869GD 预计每期还租日 20",
            "还租计划",
            "序号 应还款日期 利率 客户应还本金 客户应还利息 贴息 罚金 客户应还金额合计 还款标记",
        ]
        for line in lines:
            c.drawString(40, y, line)
            y -= 18
        for row in rows:
            period, due_date, principal, interest, subsidy, penalty, amount, status = row
            c.drawString(
                40,
                y,
                f"{period} {due_date} 3.6 {principal:,.2f} {interest:,.2f} {subsidy:,.2f} {penalty:,.2f} {amount:,.2f} {status}",
            )
            y -= 18
        c.drawString(40, y, "合计 50,000.00 275.00 961.49 0.00 50,275.00")
        c.save()
        return "/uploads/factory-plan.pdf"

    def prepare_sales_order_plan_for_activation(self, order_id, plan_overrides=None, factory_rows=None, confirm_difference=False):
        self.login("ops")
        plan_response = self.client.put(
            f"/api/sales-orders/{order_id}/planning-contract",
            json=plan_overrides or {},
        )
        self.assertEqual(plan_response.status_code, 200, plan_response.get_json())
        contract_id = plan_response.get_json()["contract_id"]
        self.logout()

        xlsx_url = self.create_factory_plan_xlsx(factory_rows or [
            [1, "2026-02-15", 2500],
            [2, "2026-03-15", 2500],
        ])
        self.login("fin")
        import_response = self.client.post(
            f"/api/contracts/{contract_id}/factory-repayments/import",
            json={"file_url": xlsx_url},
        )
        self.assertEqual(import_response.status_code, 200, import_response.get_json())
        if confirm_difference and import_response.get_json()["comparison"]["status"] == "差异待处理":
            confirm_response = self.client.post(
                f"/api/contracts/{contract_id}/plan-compare/confirm-difference",
                json={"comment": "自动化测试确认计划差异"},
            )
            self.assertEqual(confirm_response.status_code, 200, confirm_response.get_json())
        self.logout()
        return contract_id

    def test_sales_order_activation_factory_import_warnings_and_export_flow(self):
        warning_date = (datetime.now() + timedelta(days=20)).strftime("%Y-%m-%d")
        vehicle_id = self.create_vehicle({
            "insurance_expiry_date": warning_date,
            "annual_review_date": warning_date,
        })
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-12",
                "customer_name": "报单客户",
                "customer_phone": "13900000000",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"].lower(),
                "car_type": "解放轻卡4米2-领途150马力",
                "vehicle_color": "白色",
                "plate_number": vehicle["plate_number"],
                "lease_term": "36期",
                "cargo_length": "4米2",
                "sale_total_price": 128000,
                "payment_category": "定金",
                "car_purchase_amount": 30000,
                "vehicle_rent_amount": 3500,
                "receiving_company": "陕西金聚源汽车服务有限公司",
                "wechat_interest": 600,
                "wechat_registration_fee": 300,
                "wechat_purchase_tax": 5000,
                "sales_advisor": "周销售",
                "full_package": True,
                "wechat_private_fee": 188,
                "gifted_items": "赠送首保",
                "deposit_amount": 3000,
                "remark": "端到端测试报单",
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        locked_state = self.db_value(
            "SELECT status FROM vehicles WHERE id=?",
            (vehicle_id,),
        )
        self.assertEqual(locked_state, "报单锁定中")

        self.login("ops")
        premature_contract = self.client.post(
            "/api/contracts",
            json={
                "sales_order_id": order_id,
                "vehicle_id": vehicle_id,
                "customer_name": "报单客户",
                "customer_phone": "13900000000",
                "contract_type": "租赁",
                "business_mode": "转租",
                "rental_method": "经营租赁",
                "repayment_day": 10,
                "start_date": "2026-06-01",
                "loan_periods": 6,
                "rent": 3500,
                "monthly_payment": 2800,
                "factory_periods": 6,
                "factory_repayment_months": 6,
                "deposit": 3000,
                "down_payment": 2000,
                "customer_loan_amount": 18000,
            },
        )
        self.assertEqual(premature_contract.status_code, 400, premature_contract.get_json())
        self.assertIn("财务确认报单", premature_contract.get_json()["message"])
        self.logout()

        self.prepare_sales_order_plan_for_activation(
            order_id,
            plan_overrides={
                "contract_type": "租赁",
                "business_mode": "转租",
                "repayment_day": 10,
                "start_date": "2026-06-01",
                "loan_periods": 6,
                "rent": 3500,
                "monthly_payment": 2800,
                "factory_periods": 6,
                "factory_repayment_months": 6,
                "deposit": 3000,
                "down_payment": 2000,
                "customer_loan_amount": 18000,
            },
            factory_rows=[
                [1, "2026-07-10", 2800],
                [2, "2026-08-10", 2800],
                [3, "2026-09-10", 2800],
                [4, "2026-10-10", 2800],
                [5, "2026-11-10", 2800],
                [6, "2026-12-10", 2800],
            ],
        )

        self.login("fin")
        activate_response = self.client.post(f"/api/sales-orders/{order_id}/activate", json={})
        self.assertEqual(activate_response.status_code, 200, activate_response.get_json())
        self.logout()

        contract_id = self.create_contract(vehicle_id, overrides={"sales_order_id": order_id})
        order_state = self.db_value(
            """
            SELECT order_status, finance_confirmed_by, contract_id, car_type, vehicle_color, plate_number,
                   lease_term, cargo_length, sale_total_price, payment_category, car_purchase_amount,
                   vehicle_rent_amount, receiving_company, wechat_interest, wechat_registration_fee,
                   wechat_purchase_tax, sales_advisor
            FROM sales_orders WHERE id=?
            """,
            (order_id,),
        )
        self.assertEqual(order_state["order_status"], "已激活")
        self.assertEqual(order_state["finance_confirmed_by"], "张财务")
        self.assertEqual(order_state["contract_id"], contract_id)
        self.assertEqual(order_state["car_type"], "解放轻卡4米2-领途150马力")
        self.assertEqual(order_state["vehicle_color"], "白色")
        self.assertEqual(order_state["plate_number"], vehicle["plate_number"])
        self.assertEqual(order_state["lease_term"], "36期")
        self.assertEqual(order_state["cargo_length"], "4米2")
        self.assertEqual(order_state["sale_total_price"], 128000)
        self.assertEqual(order_state["payment_category"], "定金")
        self.assertEqual(order_state["car_purchase_amount"], 30000)
        self.assertEqual(order_state["vehicle_rent_amount"], 3500)
        self.assertEqual(order_state["receiving_company"], "陕西金聚源汽车服务有限公司")
        self.assertEqual(order_state["wechat_interest"], 600)
        self.assertEqual(order_state["wechat_registration_fee"], 300)
        self.assertEqual(order_state["wechat_purchase_tax"], 5000)
        self.assertEqual(order_state["sales_advisor"], "周销售")

        self.login("fleet")
        warnings_response = self.client.get("/api/risk/insurance-expiry")
        self.assertEqual(warnings_response.status_code, 200, warnings_response.get_json())
        warning_rows = [row for row in warnings_response.get_json() if row["id"] == vehicle_id]
        self.assertEqual(len(warning_rows), 1)
        self.assertTrue(warning_rows[0]["need_warning"])
        self.assertLessEqual(warning_rows[0]["insurance_days_left"], 30)
        self.logout()

        xlsx_url = self.create_factory_plan_xlsx([
            [1, "2026-07-10", 4100],
            [2, "2026-08-10", None],
        ])
        self.login("ops")
        import_response = self.client.post(
            f"/api/contracts/{contract_id}/factory-repayments/import",
            json={"file_url": xlsx_url},
        )
        self.assertEqual(import_response.status_code, 200, import_response.get_json())
        self.logout()

        imported_plan = self.client.get(f"/api/contracts/{contract_id}/factory-repayments")
        self.assertEqual(imported_plan.status_code, 200, imported_plan.get_json())
        plan_rows = imported_plan.get_json()
        self.assertEqual(len(plan_rows), 2)
        self.assertEqual(plan_rows[0]["due_date"], "2026-07-10")
        self.assertEqual(plan_rows[0]["amount"], 4100)
        self.assertEqual(plan_rows[1]["amount"], 2500)

        pdf_url = self.create_factory_plan_pdf([
            [1, "2026-07-20", 4166.67, 0, 149.55, 0, 4166.67, "已归还"],
            [2, "2026-08-20", 4166.67, 0, 123.88, 0, 4166.67, "未归还"],
        ])
        self.login("ops")
        pdf_import = self.client.post(
            f"/api/contracts/{contract_id}/factory-repayments/import",
            json={"file_url": pdf_url},
        )
        self.assertEqual(pdf_import.status_code, 200, pdf_import.get_json())
        self.assertEqual(pdf_import.get_json()["import_meta"]["source_format"], "factory_plan_pdf")
        self.assertEqual(pdf_import.get_json()["import_meta"]["row_count"], 2)
        self.logout()

        pdf_plan = self.client.get(f"/api/contracts/{contract_id}/factory-repayments")
        self.assertEqual(pdf_plan.status_code, 200, pdf_plan.get_json())
        pdf_rows = pdf_plan.get_json()
        self.assertEqual(len(pdf_rows), 2)
        self.assertEqual(pdf_rows[0]["due_date"], "2026-07-20")
        self.assertEqual(pdf_rows[0]["amount"], 4166.67)
        self.assertIn("PDF状态:已归还", pdf_rows[0]["remark"])
        self.assertIn("PDF状态:未归还", pdf_rows[1]["remark"])

        self.login("fin")
        csv_path = os.path.join(app_module.UPLOAD_DIR, "factory-plan.csv")
        with open(csv_path, "w", encoding="utf-8") as csv_file:
            csv_file.write("序号,应还款日期,客户应还金额合计\n1,2026-09-10,4100\n")
        csv_import = self.client.post(
            f"/api/contracts/{contract_id}/factory-repayments/import",
            json={"file_url": "/uploads/factory-plan.csv"},
        )
        self.assertEqual(csv_import.status_code, 400, csv_import.get_json())
        self.assertIn("xlsx 或 PDF", csv_import.get_json()["message"])
        self.logout()

        self.login("sales")
        export_response = self.client.post(f"/api/contracts/{contract_id}/export", json={})
        self.assertEqual(export_response.status_code, 403, export_response.get_json())
        self.logout()

        self.login("ops")
        export_response = self.client.post(f"/api/contracts/{contract_id}/export", json={})
        self.assertIn(export_response.status_code, (200, 400), export_response.get_json())
        export_json = export_response.get_json()
        if export_response.status_code == 200:
            self.assertTrue(export_json["docx_url"].endswith(".docx"))
            self.assertTrue(export_json["pdf_url"].endswith(".pdf"))
            self.assertTrue(os.path.exists(os.path.join(
                app_module.CONTRACT_OUTPUT_DIR,
                os.path.basename(export_json["docx_url"]),
            )))
            self.assertTrue(os.path.exists(os.path.join(
                app_module.CONTRACT_OUTPUT_DIR,
                os.path.basename(export_json["pdf_url"]),
            )))
        else:
            self.assertIn("生成合同失败", export_json["message"])
        self.logout()

        self.login("boss")
        boss_export_response = self.client.post(f"/api/contracts/{contract_id}/export", json={})
        self.assertIn(boss_export_response.status_code, (200, 400), boss_export_response.get_json())
        self.logout()

    def test_sales_order_list_carries_existing_contract_customer_installments(self):
        vehicle_id = self.create_vehicle()
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-06-05",
                "customer_name": "补合同客户",
                "customer_phone": "13955556666",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"],
                "lease_term": "36期",
                "sale_total_price": 180000,
                "vehicle_rent_amount": 0,
                "payment_category": "押金",
                "deposit_amount": 20000,
                "receiving_company": "陕西金聚源汽车服务有限公司",
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        self.prepare_sales_order_plan_for_activation(
            order_id,
            plan_overrides={
                "contract_type": "租赁",
                "business_mode": "经营租赁",
                "start_date": "2026-06-05",
                "loan_periods": 36,
                "rent": 5000,
                "deposit": 20000,
                "down_payment": 0,
            },
            factory_rows=[
                [idx, f"2026-{month:02d}-05", 4000]
                for idx, month in enumerate(list(range(7, 13)) + list(range(1, 13)) * 3, start=1)
            ][:36],
        )

        self.login("fin")
        activate_response = self.client.post(f"/api/sales-orders/{order_id}/activate", json={})
        self.assertEqual(activate_response.status_code, 200, activate_response.get_json())
        self.logout()

        contract_id = self.create_contract(
            vehicle_id,
            overrides={
                "sales_order_id": order_id,
                "start_date": "2026-06-05",
                "loan_periods": 36,
                "rent": 5000,
                "deposit": 20000,
                "down_payment": 0,
            },
        )

        self.login("ops")
        list_response = self.client.get("/api/sales-orders")
        self.assertEqual(list_response.status_code, 200, list_response.get_json())
        rows = list_response.get_json()
        row = next(item for item in rows if item["id"] == order_id)
        self.assertEqual(row["contract_id"], contract_id)
        self.assertEqual(row["vehicle_rent_amount"], 0)
        self.assertEqual(row["contract_rent"], 5000)
        self.assertEqual(row["contract_loan_periods"], 36)
        self.assertEqual(row["contract_deposit"], 20000)
        self.assertEqual(row["contract_start_date"], "2026-06-05")
        self.assertEqual(row["customer_plan_periods"], 36)
        self.assertEqual(row["customer_plan_total"], 180000)
        self.assertEqual(row["customer_plan_avg"], 5000)
        self.assertEqual(row["customer_plan_first_due_date"], "2026-06-05")
        self.assertEqual(row["customer_plan_last_due_date"], "2029-05-05")
        self.logout()

    def test_below_guidance_sales_order_requires_boss_price_approval_and_snapshots_price_history(self):
        vehicle_id = self.create_vehicle({"guidance_price": 100000})
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        low_order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-13",
                "customer_name": "低价客户",
                "customer_phone": "13900000001",
                "sales_mode": "卖车",
                "vin": vehicle["vin"],
                "sale_total_price": 90000,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(low_order_response.status_code, 200, low_order_response.get_json())
        low_order_id = low_order_response.get_json()["id"]
        self.logout()

        low_order = self.db_value(
            "SELECT order_status, snapshot_guidance_price, price_check_status FROM sales_orders WHERE id=?",
            (low_order_id,),
        )
        self.assertEqual(low_order["order_status"], "待价格特批")
        self.assertEqual(low_order["snapshot_guidance_price"], 98000)
        self.assertEqual(low_order["price_check_status"], "待老板审批")
        self.assertEqual(
            [step["required_role"] for step in self.get_approval_steps("price_exception", low_order_id)],
            ["老板"],
        )

        self.login("fin")
        blocked_activate = self.client.post(f"/api/sales-orders/{low_order_id}/activate", json={})
        self.assertEqual(blocked_activate.status_code, 400, blocked_activate.get_json())
        self.logout()

        self.approve_latest_flow("price_exception", low_order_id)
        approved_order = self.db_value(
            "SELECT order_status, price_check_status, boss_price_approved_by FROM sales_orders WHERE id=?",
            (low_order_id,),
        )
        self.assertEqual(approved_order["order_status"], "待财务确认")
        self.assertEqual(approved_order["price_check_status"], "已通过")
        self.assertEqual(approved_order["boss_price_approved_by"], "王老板")

        self.login("boss")
        update_response = self.client.post(
            f"/api/vehicles/{vehicle_id}/guidance_price",
            json={"price": 120000},
        )
        self.assertEqual(update_response.status_code, 200, update_response.get_json())
        self.logout()

        history = self.db_value(
            "SELECT old_price, new_price, changed_by FROM vehicle_guidance_price_history WHERE vehicle_id=? ORDER BY id DESC LIMIT 1",
            (vehicle_id,),
        )
        self.assertEqual(history["old_price"], 100000)
        self.assertEqual(history["new_price"], 120000)
        self.assertEqual(history["changed_by"], "王老板")
        self.assertEqual(
            self.db_value("SELECT snapshot_guidance_price FROM sales_orders WHERE id=?", (low_order_id,)),
            98000,
        )
        self.login("fin")
        ops_approvals = self.client.get("/api/approvals?ref_type=price_exception").get_json()
        self.assertEqual(len(ops_approvals), 1)
        self.assertEqual(ops_approvals[0]["ref_type"], "price_exception")
        self.logout()

        vehicle_id_2 = self.create_vehicle({"guidance_price": 120000})
        vehicle_2 = self.get_vehicle(vehicle_id_2)
        self.login("sales")
        normal_order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-13",
                "customer_name": "正常客户",
                "customer_phone": "13900000002",
                "sales_mode": "卖车",
                "vin": vehicle_2["vin"],
                "sale_total_price": 125000,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(normal_order_response.status_code, 200, normal_order_response.get_json())
        normal_order_id = normal_order_response.get_json()["id"]
        self.logout()
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (normal_order_id,)),
            "待财务确认",
        )
        self.assertEqual(self.get_approval_steps("price_exception", normal_order_id), [])

    def test_missing_guidance_sales_order_requires_boss_price_approval_and_vehicle_lock(self):
        vehicle_id = self.create_vehicle({"guidance_price": 0, "car_type": "未知车型-缺少指导价"})
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-13",
                "customer_name": "缺指导客户",
                "customer_phone": "13900000003",
                "sales_mode": "卖车",
                "vin": vehicle["vin"],
                "sale_total_price": 120000,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        order = self.db_value(
            "SELECT order_status, price_check_status, snapshot_guidance_price FROM sales_orders WHERE id=?",
            (order_id,),
        )
        self.assertEqual(order["order_status"], "待价格特批")
        self.assertEqual(order["price_check_status"], "待老板审批")
        self.assertEqual(order["snapshot_guidance_price"], 0)
        self.assertEqual(
            [step["required_role"] for step in self.get_approval_steps("price_exception", order_id)],
            ["老板"],
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "报单锁定中",
        )

        # Add a second sales user for duplicate lock prevention.
        self.db_execute(
            "INSERT OR IGNORE INTO users (username, password, display_name, role) VALUES (?, ?, ?, ?)",
            ("sales2", "123456", "李销售", "销售"),
        )
        self.login("sales2")
        duplicate_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-13",
                "customer_name": "重复客户",
                "customer_phone": "13900000004",
                "sales_mode": "卖车",
                "vin": vehicle["vin"],
                "sale_total_price": 125000,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(duplicate_response.status_code, 400)
        self.assertIn("报单锁定", duplicate_response.get_json().get("message", ""))
        self.logout()

    def test_boss_can_manage_model_guidance_price_for_car_type(self):
        car_type = "测试车型-统一指导价"
        vehicle_id = self.create_vehicle({"car_type": car_type, "guidance_price": 90000})
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        list_as_sales = self.client.get("/api/model-guidance-prices")
        self.assertEqual(list_as_sales.status_code, 200, list_as_sales.get_json())
        forbidden_response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "guidance_price": 130000,
                "lease_installment_price": 3500,
                "sale_total_price": 130000,
                "lease_deposit_ratio": 0.10,
                "lease_repayment_ratio": 0.025,
                "sale_down_payment_ratio": 0.15,
                "sale_repayment_ratio": 0.025,
            },
        )
        self.assertEqual(forbidden_response.status_code, 403, forbidden_response.get_json())
        self.logout()

        self.login("boss")
        update_response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "guidance_price": 130000,
                "lease_installment_price": 3500,
                "sale_total_price": 130000,
                "lease_deposit_ratio": 0.10,
                "lease_repayment_ratio": 0.025,
                "sale_down_payment_ratio": 0.15,
                "sale_repayment_ratio": 0.025,
                "remark": "测试车型统一调价",
            },
        )
        self.assertEqual(update_response.status_code, 200, update_response.get_json())
        list_response = self.client.get("/api/model-guidance-prices")
        self.assertEqual(list_response.status_code, 200, list_response.get_json())
        self.assertIn(car_type, [row["car_type"] for row in list_response.get_json()])
        self.logout()

        self.assertEqual(
            self.db_value("SELECT guidance_price FROM vehicles WHERE id=?", (vehicle_id,)),
            130000,
        )
        model_history = self.db_value(
            "SELECT old_price, new_price, affected_vehicle_count FROM model_guidance_price_history WHERE car_type=? ORDER BY id DESC LIMIT 1",
            (car_type,),
        )
        self.assertEqual(model_history["old_price"], 0)
        self.assertEqual(model_history["new_price"], 130000)
        self.assertEqual(model_history["affected_vehicle_count"], 1)

    def test_sales_order_creates_finance_approval_and_can_be_seen_by_finance(self):
        vehicle_id = self.create_vehicle({"guidance_price": 100000})
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-14",
                "customer_name": "财务待办客户",
                "customer_phone": "13922223333",
                "sales_mode": "整车销售",
                "vin": vehicle["vin"],
                "sale_total_price": 125000,
                "payment_category": "定金",
                "deposit_amount": 5000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        steps = self.get_approval_steps("sale_payment", order_id)
        self.assertEqual([step["required_role"] for step in steps], ["财务"])
        self.assertEqual(self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)), "待财务确认")

        self.login("fin")
        reminders = self.client.get("/api/approvals?ref_type=sale_payment").get_json()
        self.assertIn(order_id, [item["ref_id"] for item in reminders])
        self.logout()

        self.login("fin")
        activate_response = self.client.post(f"/api/sales-orders/{order_id}/activate", json={})
        self.assertEqual(activate_response.status_code, 200, activate_response.get_json())
        self.logout()

        self.assertEqual(self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)), "已激活")

    def test_finance_approval_center_allows_installment_order_without_factory_plan(self):
        vehicle_id = self.create_vehicle({"guidance_price": 100000})
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-14",
                "customer_name": "审批中心分期客户",
                "customer_phone": "13922224444",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"],
                "lease_term": "2期",
                "sale_total_price": 128000,
                "vehicle_rent_amount": 3500,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        flow_id = self.get_approval_steps("sale_payment", order_id)[0]["id"]
        self.login("fin")
        approve_response = self.client.post(
            f"/api/approvals/{flow_id}/approve",
            json={"comment": "客户计划已生成，厂家计划无需上传"},
        )
        self.assertEqual(approve_response.status_code, 200, approve_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT status FROM approval_flows WHERE id=?", (flow_id,)),
            "已通过",
        )
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "已激活",
        )
        order_plan = self.db_value(
            "SELECT customer_plan_match_status, factory_plan_match_status FROM sales_orders WHERE id=?",
            (order_id,),
        )
        self.assertEqual(order_plan["customer_plan_match_status"], "已生成")
        self.assertEqual(order_plan["factory_plan_match_status"], "无需上传")

    def test_low_guidance_order_visible_to_sales_boss_and_ops_after_approval(self):
        vehicle_id = self.create_vehicle({"guidance_price": 100000})
        vehicle = self.get_vehicle(vehicle_id)
        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-14",
                "customer_name": "低价流程客户",
                "customer_phone": "13922220000",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"],
                "sale_total_price": 90000,
                "vehicle_rent_amount": 2500,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        sales_approvals = self.client.get("/api/approvals?ref_type=price_exception").get_json()
        self.assertEqual(len(sales_approvals), 1)
        self.logout()

        self.approve_latest_flow("price_exception", order_id)
        contract_id = self.db_value("SELECT contract_id FROM sales_orders WHERE id=?", (order_id,))
        approval_status = self.client.get(f"/api/contracts/{contract_id}/approval-status").get_json()
        self.assertEqual(approval_status["payment"]["status"], "审批中")
        self.assertEqual(approval_status["payment"]["steps"][0]["ref_id"], order_id)

        self.login("sales")
        sales_approvals_after = self.client.get("/api/approvals?ref_type=price_exception").get_json()
        self.assertEqual(len(sales_approvals_after), 1)
        self.assertEqual(sales_approvals_after[0]["overall_status"], "已完成")
        self.logout()

        self.login("fin")
        ops_approvals = self.client.get("/api/approvals?ref_type=price_exception").get_json()
        self.assertEqual(len(ops_approvals), 1)
        self.assertEqual(ops_approvals[0]["overall_status"], "已完成")
        self.assertEqual(ops_approvals[0]["follow_up_role"], "财务")
        self.logout()

    def test_boss_receives_alert_when_new_vehicle_has_no_guidance_price(self):
        car_type = "测试车型-缺指导价提醒"
        vehicle_id = self.create_vehicle({
            "car_type": car_type,
            "guidance_price": 0,
            "invoice_price": 88000,
        })
        self.assertEqual(
            self.db_value("SELECT guidance_price FROM vehicles WHERE id=?", (vehicle_id,)),
            0,
        )

        self.login("fleet")
        forbidden_alerts = self.client.get("/api/guidance-price-alerts")
        self.assertEqual(forbidden_alerts.status_code, 403, forbidden_alerts.get_json())
        self.logout()

        self.login("boss")
        alerts_response = self.client.get("/api/guidance-price-alerts")
        self.assertEqual(alerts_response.status_code, 200, alerts_response.get_json())
        alerts = alerts_response.get_json()
        self.assertGreaterEqual(alerts["count"], 1)
        self.assertIn(vehicle_id, [row["id"] for row in alerts["items"]])

        update_response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "guidance_price": 99000,
                "lease_installment_price": 3200,
                "sale_total_price": 99000,
                "lease_deposit_ratio": 0.10,
                "lease_repayment_ratio": 0.025,
                "sale_down_payment_ratio": 0.15,
                "sale_repayment_ratio": 0.025,
                "remark": "补充缺失指导价",
            },
        )
        self.assertEqual(update_response.status_code, 200, update_response.get_json())
        cleared_alerts = self.client.get("/api/guidance-price-alerts").get_json()
        self.assertNotIn(vehicle_id, [row["id"] for row in cleared_alerts["items"]])
        self.logout()

        vehicle_id_2 = self.create_vehicle({
            "car_type": car_type,
            "guidance_price": 0,
            "invoice_price": 90000,
        })
        self.assertEqual(
            self.db_value("SELECT guidance_price FROM vehicles WHERE id=?", (vehicle_id_2,)),
            99000,
        )

    def test_waterfall_reconciliation_allocates_fee_items_before_rent(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        self.approve_latest_flow("contract_delivery", contract_id)
        self.complete_initial_payment(contract_id)
        self.deliver_vehicle(vehicle_id)

        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2",
            (contract_id,),
        )

        self.login("ops")
        insurance_fee = self.client.post(
            f"/api/contracts/{contract_id}/fee-items",
            json={
                "fee_type": "insurance_fee",
                "amount_due": 800,
                "due_date": "2026-06-01",
                "description": "保险上浮费",
            },
        )
        penalty_fee = self.client.post(
            f"/api/contracts/{contract_id}/fee-items",
            json={
                "fee_type": "penalty_fee",
                "amount_due": 500,
                "due_date": "2026-06-02",
                "description": "违章罚金",
            },
        )
        late_fee = self.client.post(
            f"/api/contracts/{contract_id}/fee-items",
            json={
                "fee_type": "late_fee",
                "amount_due": 300,
                "due_date": "2026-06-03",
                "description": "滞纳金",
            },
        )
        self.assertEqual(insurance_fee.status_code, 200, insurance_fee.get_json())
        self.assertEqual(penalty_fee.status_code, 200, penalty_fee.get_json())
        self.assertEqual(late_fee.status_code, 200, late_fee.get_json())
        self.logout()

        self.login("ops")
        screenshot_response = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/customer-waterfall.png"},
        )
        self.assertEqual(screenshot_response.status_code, 200, screenshot_response.get_json())
        self.logout()

        self.login("fin")
        receipt_response = self.client.post(
            f"/api/reconciliation/{repayment_id}/receipt",
            json={"bank_receipt_path": "/uploads/bank-waterfall.png"},
        )
        verify_response = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "WF123456", "received_amount": 2000},
        )
        self.assertEqual(receipt_response.status_code, 200, receipt_response.get_json())
        self.assertEqual(verify_response.status_code, 200, verify_response.get_json())
        allocation = verify_response.get_json()["allocation"]
        self.assertEqual(allocation["status"], "部分核销")
        self.assertEqual(allocation["rent_allocated"], 400)
        self.assertEqual(allocation["shortfall_amount"], 2600)
        self.assertTrue(allocation["shortfall_receivable_id"])
        self.assertIn("insurance_fee", allocation["summary"])
        self.assertIn("penalty_fee", allocation["summary"])
        self.assertIn("late_fee", allocation["summary"])
        self.logout()

        self.login("ops")
        fee_states = self.client.get(f"/api/contracts/{contract_id}/fee-items").get_json()
        self.logout()
        self.assertTrue(all(row["status"] == "已支付" for row in fee_states))
        repayment_state = self.db_value(
            "SELECT paid_amount, verified_amount, status, waterfall_summary FROM repayments WHERE id=?",
            (repayment_id,),
        )
        self.assertEqual(repayment_state["paid_amount"], 400)
        self.assertEqual(repayment_state["verified_amount"], 2000)
        self.assertEqual(repayment_state["status"], "部分核销")
        self.assertIn("rent", repayment_state["waterfall_summary"])
        self.assertEqual(
            self.db_value("SELECT amount FROM repayments WHERE contract_id=? AND period=3", (contract_id,)),
            3000.0,
        )
        receivable = self.db_value(
            "SELECT receivable_type, amount, status FROM receivables WHERE repayment_id=?",
            (repayment_id,),
        )
        self.assertEqual(receivable["receivable_type"], "period_shortfall")
        self.assertEqual(receivable["amount"], 2600)
        self.assertEqual(receivable["status"], "待归还")

        self.login("fin")
        allocations = self.client.get(f"/api/reconciliation/{repayment_id}/allocations").get_json()
        self.logout()
        self.assertEqual(
            [row["allocation_type"] for row in allocations],
            ["insurance_fee", "penalty_fee", "late_fee", "rent"],
        )

    def test_initial_payment_shortage_requires_boss_approval_and_creates_receivable(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        payment_id = self.initiate_initial_payment(contract_id)

        self.login("fin")
        receipt_response = self.client.post(
            f"/api/initial-payments/{payment_id}/receipt",
            json={
                "bank_receipt_path": "/uploads/initial-bank-short.png",
                "bank_serial": "INITSHORT001",
                "received_amount": 4000,
                "shortage_reason": "客户临时资金不足",
                "promised_repay_date": "2026-02-01",
            },
        )
        self.assertEqual(receipt_response.status_code, 200, receipt_response.get_json())
        finance_flow_id = self.get_approval_steps("initial_payment", payment_id)[0]["id"]
        approve_response = self.client.post(
            f"/api/approvals/{finance_flow_id}/approve",
            json={"comment": "实收不足，转老板审批"},
        )
        self.assertEqual(approve_response.status_code, 200, approve_response.get_json())
        self.logout()

        payment_state = self.db_value(
            "SELECT status, shortage_amount, shortage_status FROM contract_initial_payments WHERE id=?",
            (payment_id,),
        )
        self.assertEqual(payment_state["status"], "待老板审批")
        self.assertEqual(payment_state["shortage_amount"], 1000)
        self.assertEqual(payment_state["shortage_status"], "待老板审批")
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "首付不足待审批",
        )

        self.approve_latest_flow("initial_payment_shortage", payment_id)
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待出库",
        )
        receivable = self.db_value(
            "SELECT receivable_type, amount, promised_repay_date, status FROM receivables WHERE initial_payment_id=?",
            (payment_id,),
        )
        self.assertEqual(receivable["receivable_type"], "initial_payment_shortfall")
        self.assertEqual(receivable["amount"], 1000)
        self.assertEqual(receivable["promised_repay_date"], "2026-02-01")
        self.assertEqual(receivable["status"], "待归还")

        self.deliver_vehicle(vehicle_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "租赁中",
        )

    def test_overpayment_requires_confirmed_periods_and_prepays_selected_months(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        self.approve_latest_flow("contract_delivery", contract_id)
        self.complete_initial_payment(contract_id)
        self.deliver_vehicle(vehicle_id)

        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2",
            (contract_id,),
        )

        self.login("ops")
        self.assertEqual(
            self.client.post(
                f"/api/reconciliation/{repayment_id}/screenshot",
                json={"screenshot_path": "/uploads/customer-overpay.png"},
            ).status_code,
            200,
        )
        self.logout()

        self.login("fin")
        self.assertEqual(
            self.client.post(
                f"/api/reconciliation/{repayment_id}/receipt",
                json={"bank_receipt_path": "/uploads/bank-overpay.png"},
            ).status_code,
            200,
        )
        missing_period_response = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "OP123456", "received_amount": 6000},
        )
        self.assertEqual(missing_period_response.status_code, 400, missing_period_response.get_json())

        verify_response = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "OP123456", "received_amount": 6000, "extra_alloc_periods": [3]},
        )
        self.assertEqual(verify_response.status_code, 200, verify_response.get_json())
        allocation = verify_response.get_json()["allocation"]
        self.assertEqual(allocation["status"], "已还款")
        # 第3期被一整期完全覆盖 -> 直接“已还款”（不再是“预抵”），与来源第2期共用流水号/截图
        self.assertEqual(allocation["extra_allocated_rows"], [{"period": 3, "amount": 3000.0, "status": "已还款"}])
        self.logout()

        period3 = self.db_value(
            "SELECT paid_amount, verified_amount, status, waterfall_summary, bank_serial FROM repayments WHERE contract_id=? AND period=3",
            (contract_id,),
        )
        self.assertEqual(period3["paid_amount"], 3000.0)
        self.assertEqual(period3["verified_amount"], 3000.0)
        self.assertEqual(period3["status"], "已还款")
        self.assertEqual(period3["bank_serial"], "OP123456")
        self.assertIn("多还抵扣来源第2期", period3["waterfall_summary"])

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
        role_to_user = {"运营": "ops", "财务": "fin", "老板": "boss"}
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
        self.login("ops")
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
            json={
                "bank_receipt_path": "/uploads/initial-bank-receipt.png",
                "bank_serial": f"INIT{payment_id:06d}",
                "received_amount": self.db_value(
                    "SELECT amount FROM contract_initial_payments WHERE id=?",
                    (payment_id,),
                ),
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.logout()

    def complete_initial_payment(self, contract_id, amount=None):
        payment_id = self.initiate_initial_payment(contract_id, amount=amount)
        self.upload_initial_payment_receipt(payment_id)
        self.approve_latest_flow("initial_payment", payment_id)
        return payment_id

    def upload_delivery_files(self, contract_id, username="fleet"):
        self.login(username)
        response = self.client.post(
            f"/api/contracts/{contract_id}/delivery-files",
            json={
                "delivery_photo_path": "/uploads/delivery-photo.png",
                "delivery_document_path": "/uploads/delivery-document.png",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        self.logout()

    def deliver_vehicle(self, vehicle_id):
        contract_id = self.db_value(
            "SELECT id FROM contracts WHERE vehicle_id=? ORDER BY id DESC LIMIT 1",
            (vehicle_id,),
        )
        self.upload_delivery_files(contract_id)
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
                "mileage": "99999",
                "tool_triangle": True,
                "rent_late_fee": 999,
                "actual_refund": 999,
                "refund_bank_card_no": "9999999999999999999",
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        inspection_id = response.get_json()["id"]
        created_state = self.db_value(
            "SELECT status, mileage, tool_triangle, rent_late_fee, actual_refund, refund_bank_card_no FROM return_inspections WHERE id=?",
            (inspection_id,),
        )
        self.assertEqual(created_state["status"], "待车管验车")
        self.assertEqual(created_state["mileage"], "")
        self.assertEqual(created_state["tool_triangle"], 0)
        self.assertEqual(created_state["rent_late_fee"], 0)
        self.assertEqual(created_state["actual_refund"], 0)
        self.assertNotEqual(created_state["refund_bank_card_no"], "9999999999999999999")
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

    def prepare_delivered_rental_vehicle(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        self.approve_latest_flow("contract_delivery", contract_id)
        self.complete_initial_payment(contract_id)
        self.deliver_vehicle(vehicle_id)
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "租赁中",
        )
        return vehicle_id, contract_id

    def test_rental_delivery_reject_resubmit_and_reconciliation_flow(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)

        self.assertEqual(self.get_approval_steps("contract_delivery", contract_id), [])
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
            "待首付款",
        )
        contract_clause = self.db_value("SELECT company, yard FROM contracts WHERE id=?", (contract_id,))
        self.assertEqual(contract_clause["company"], "陕西金聚源汽车服务有限公司")
        self.assertEqual(contract_clause["yard"], "西安一号车场")

        self.login("fleet")
        premature_delivery = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(premature_delivery.status_code, 400, premature_delivery.get_json())
        self.logout()

        payment_id = self.initiate_initial_payment(contract_id)
        payment_roles = [step["required_role"] for step in self.get_approval_steps("initial_payment", payment_id)]
        self.assertEqual(payment_roles, ["财务"])

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

        self.login("fleet")
        delivery_without_files = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(delivery_without_files.status_code, 400, delivery_without_files.get_json())
        self.logout()

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
        self.assertEqual(initial_repayment_state["verified_by"], "张财务")

        recon_repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2",
            (contract_id,),
        )
        factory_repayment_id = self.db_value(
            "SELECT id FROM factory_repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )

        self.login("ops")
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
            "SELECT id FROM repayments WHERE contract_id=? AND period=3",
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

    def test_guidance_price_and_business_flows_do_not_require_boss_after_price_check(self):
        scenarios = [
            ("卖车", "销售", "已售"),
            ("经营租赁", "租赁", "租赁中"),
            ("以租代售", "以租代售", "以租代售"),
        ]
        for sales_mode, contract_type, expected_vehicle_status in scenarios:
            with self.subTest(sales_mode=sales_mode):
                vehicle_id = self.create_vehicle({"guidance_price": 100000})
                vehicle = self.get_vehicle(vehicle_id)

                self.login("sales")
                order_response = self.client.post(
                    "/api/sales-orders",
                    json={
                        "payment_date": "2026-05-14",
                        "customer_name": f"{sales_mode}高价客户",
                        "customer_phone": "13911110000",
                        "sales_mode": sales_mode,
                        "vin": vehicle["vin"],
                        "sale_total_price": 108000,
                        "payment_category": "定金",
                        "deposit_amount": 3000,
                    },
                )
                self.assertEqual(order_response.status_code, 200, order_response.get_json())
                order_id = order_response.get_json()["id"]
                self.logout()

                self.assertEqual(self.get_approval_steps("price_exception", order_id), [])
                self.assertEqual(
                    self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
                    "待财务确认",
                )

                if contract_type != "销售":
                    self.prepare_sales_order_plan_for_activation(
                        order_id,
                        plan_overrides={
                            "contract_type": contract_type,
                            "business_mode": sales_mode,
                            "start_date": "2026-01-01",
                            "loan_periods": 12,
                            "rent": 3000,
                            "monthly_payment": 2500,
                            "factory_periods": 12,
                            "deposit": 2000 if contract_type == "租赁" else 0,
                            "down_payment": 30000 if contract_type == "以租代售" else 0,
                            "customer_loan_amount": 24000,
                        },
                        factory_rows=[
                            [idx, f"2026-{month:02d}-15", 2500]
                            for idx, month in enumerate(range(2, 14), start=1)
                        ],
                    )

                self.login("fin")
                activate_response = self.client.post(f"/api/sales-orders/{order_id}/activate", json={})
                self.assertEqual(activate_response.status_code, 200, activate_response.get_json())
                self.logout()

                contract_id = self.create_contract(
                    vehicle_id,
                    contract_type=contract_type,
                    business_mode=sales_mode,
                    overrides={"sales_order_id": order_id},
                )
                self.assertEqual(self.get_approval_steps("contract_delivery", contract_id), [])
                self.assertEqual(
                    self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (contract_id,)),
                    "待首付款",
                )
                payment_id = self.complete_initial_payment(contract_id)
                self.assertEqual(
                    [step["required_role"] for step in self.get_approval_steps("initial_payment", payment_id)],
                    ["财务"],
                )
                self.deliver_vehicle(vehicle_id)
                self.assertEqual(
                    self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
                    expected_vehicle_status,
                )

        low_vehicle_id = self.create_vehicle({"guidance_price": 100000})
        low_vehicle = self.get_vehicle(low_vehicle_id)
        self.login("sales")
        low_order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-14",
                "customer_name": "低价只走老板特批客户",
                "customer_phone": "13911119999",
                "sales_mode": "经营租赁",
                "vin": low_vehicle["vin"],
                "sale_total_price": 88000,
                "vehicle_rent_amount": 2500,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(low_order_response.status_code, 200, low_order_response.get_json())
        low_order_id = low_order_response.get_json()["id"]
        self.logout()

        self.assertEqual(
            [step["required_role"] for step in self.get_approval_steps("price_exception", low_order_id)],
            ["老板"],
        )
        self.approve_latest_flow("price_exception", low_order_id)
        self.prepare_sales_order_plan_for_activation(
            low_order_id,
            plan_overrides={
                "contract_type": "租赁",
                "business_mode": "经营租赁",
                "start_date": "2026-01-01",
                "loan_periods": 12,
                "rent": 3000,
                "monthly_payment": 2500,
                "factory_periods": 12,
                "deposit": 2000,
                "customer_loan_amount": 24000,
            },
            factory_rows=[
                [idx, f"2026-{month:02d}-15", 2500]
                for idx, month in enumerate(range(2, 14), start=1)
            ],
        )
        self.login("fin")
        activate_low_response = self.client.post(f"/api/sales-orders/{low_order_id}/activate", json={})
        self.assertEqual(activate_low_response.status_code, 200, activate_low_response.get_json())
        self.logout()

        low_contract_id = self.create_contract(low_vehicle_id, overrides={"sales_order_id": low_order_id})
        self.assertEqual(self.get_approval_steps("contract_delivery", low_contract_id), [])
        self.assertEqual(
            self.db_value("SELECT delivery_status FROM contracts WHERE id=?", (low_contract_id,)),
            "待首付款",
        )

    def test_overdue_urge_lock_and_return_stock_flow(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        self.approve_latest_flow("contract_delivery", contract_id)
        self.complete_initial_payment(contract_id)
        self.deliver_vehicle(vehicle_id)

        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1",
            (contract_id,),
        )
        overdue_date = (datetime.now() - timedelta(days=10)).strftime("%Y-%m-%d")
        self.db_execute(
            "UPDATE repayments SET due_date=?, status='待还款', paid_amount=0, verified_amount=0 WHERE id=?",
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
        self.login("ops")
        ops_lock_response = self.client.post(f"/api/lock-requests/{lock_request_id}/ops-review", json={"decision": "approve"})
        self.assertEqual(ops_lock_response.status_code, 200, ops_lock_response.get_json())
        self.logout()
        self.login("boss")
        boss_lock_response = self.client.post(f"/api/lock-requests/{lock_request_id}/boss-approve", json={"decision": "approve"})
        self.assertEqual(boss_lock_response.status_code, 200, boss_lock_response.get_json())
        self.logout()
        self.login("sales")
        confirm_lock_response = self.client.post(f"/api/lock-requests/{lock_request_id}/confirm", json={})
        self.assertEqual(confirm_lock_response.status_code, 200, confirm_lock_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT lock_status FROM vehicles WHERE id=?", (vehicle_id,)),
            "车辆已锁",
        )

        self.login("sales")
        locked_return_response = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "到期退车"},
        )
        self.assertEqual(locked_return_response.status_code, 400, locked_return_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT lock_status FROM vehicles WHERE id=?", (vehicle_id,)),
            "车辆已锁",
        )
        self.assertEqual(
            self.db_value("SELECT contract_status FROM contracts WHERE id=?", (contract_id,)),
            "执行中",
        )

        return_vehicle_id, return_contract_id = self.prepare_delivered_rental_vehicle()
        inspection_id = self.create_return_inspection_only(return_vehicle_id)
        self.login("fleet")
        fleet_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/fleet",
            json={
                "tool_triangle": True,
                "tool_vest": True,
                "tool_extinguisher": True,
                "tool_wedge": True,
                "tool_jack": True,
                "doc_license": True,
                "doc_keys": True,
                "mileage": "12000",
                "body_tire_clean": "已清理",
                "accident_info": "无",
                "insurance_surcharge": "0",
                "violation_info": "无",
                "etc_info": "正常",
                "maintenance_info": "正常",
            },
        )
        self.assertEqual(fleet_response.status_code, 200, fleet_response.get_json())
        self.logout()

        self.login("ops")
        operator_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/operator",
            json={
                "rent_late_fee": 0,
                "return_late_fee": 0,
                "deposit_rent_receivable": 0,
                "deposit_paid": 0,
                "total_deduction": 0,
                "actual_refund": 5000,
                "remark": "运营复核完成",
            },
        )
        self.assertEqual(operator_response.status_code, 200, operator_response.get_json())
        self.logout()

        self.login("fin")
        finance_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/finance",
            json={
                "refund_company_name": "陕西金聚源汽车服务有限公司",
                "refund_bank_name": "中国建设银行西安分行",
                "refund_bank_card_no": "6227000012345678901",
            },
        )
        self.assertEqual(finance_response.status_code, 200, finance_response.get_json())
        mismatch_response = self.client.post(
            f"/api/return-inspections/{inspection_id}/finance",
            json={
                "refund_company_name": "陕西金聚源汽车服务有限公司",
                "refund_bank_name": "中国建设银行西安分行",
                "refund_bank_card_no": "9999999999999999999",
            },
        )
        self.assertEqual(mismatch_response.status_code, 400, mismatch_response.get_json())
        self.logout()

        self.login("boss")
        boss_response = self.client.post(f"/api/return-inspections/{inspection_id}/boss-approve", json={})
        self.assertEqual(boss_response.status_code, 200, boss_response.get_json())
        self.logout()

        self.login("fin")
        pay_response = self.client.post(f"/api/return-inspections/{inspection_id}/pay", json={})
        self.assertEqual(pay_response.status_code, 200, pay_response.get_json())
        self.logout()

        inspection_status = self.db_value(
            "SELECT status FROM return_inspections WHERE id=?",
            (inspection_id,),
        )
        vehicle_state = self.db_value(
            "SELECT status, is_new FROM vehicles WHERE id=?",
            (return_vehicle_id,),
        )
        contract_status = self.db_value(
            "SELECT contract_status FROM contracts WHERE id=?",
            (return_contract_id,),
        )

        self.assertEqual(inspection_status, "已完成")
        self.assertEqual(vehicle_state["status"], "在库")
        self.assertEqual(vehicle_state["is_new"], "二手车")
        self.assertEqual(contract_status, "已结清")

    def test_receiving_company_invoice_rent_waiver_and_transfer_minimum_flows(self):
        self.login("fin")
        company_response = self.client.post(
            "/api/receiving-companies",
            json={
                "company_name": "陕西金聚源测试收款有限公司",
                "bank_name": "测试银行",
                "bank_account_no": "100200300",
                "tax_no": "TAX-001",
                "remark": "测试主数据",
            },
        )
        self.assertEqual(company_response.status_code, 200, company_response.get_json())
        self.logout()

        self.login("sales")
        company_list = self.client.get("/api/receiving-companies")
        self.assertEqual(company_list.status_code, 200, company_list.get_json())
        self.assertIn(
            "陕西金聚源测试收款有限公司",
            [row["company_name"] for row in company_list.get_json()],
        )
        self.logout()

        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)

        self.login("ops")
        invoice_response = self.client.post(
            f"/api/contracts/{contract_id}/invoices",
            json={
                "amount": 3000,
                "period": 1,
                "receiving_company": "陕西金聚源测试收款有限公司",
                "invoice_entity_name": "测试开票客户",
                "invoice_entity_tax_no": "91330000TEST",
            },
        )
        self.assertEqual(invoice_response.status_code, 200, invoice_response.get_json())
        invoice_id = invoice_response.get_json()["id"]
        self.logout()

        self.login("fin")
        issue_response = self.client.post(
            f"/api/invoice-requests/{invoice_id}/issue",
            json={"invoice_no": "INV-001", "invoice_file_path": "/uploads/invoice-001.pdf"},
        )
        self.assertEqual(issue_response.status_code, 200, issue_response.get_json())
        red_response = self.client.post(
            f"/api/invoice-requests/{invoice_id}/red",
            json={"red_invoice_no": "RED-001", "reason": "抬头错误"},
        )
        self.assertEqual(red_response.status_code, 200, red_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT status FROM invoice_requests WHERE id=?", (invoice_id,)),
            "已红冲",
        )

        self.login("sales")
        waiver_response = self.client.post(
            f"/api/contracts/{contract_id}/waivers",
            json={
                "waiver_kind": "rent",
                "target_period_list": [2],
                "waive_amount": 500,
                "reason": "客户协商租金减免",
            },
        )
        self.assertEqual(waiver_response.status_code, 200, waiver_response.get_json())
        waiver_id = waiver_response.get_json()["waiver_id"]
        self.logout()

        self.login("boss")
        approve_response = self.client.post(f"/api/waivers/{waiver_id}/approve", json={})
        self.assertEqual(approve_response.status_code, 200, approve_response.get_json())
        self.logout()

        self.login("fin")
        execute_response = self.client.post(f"/api/waivers/{waiver_id}/execute", json={})
        self.assertEqual(execute_response.status_code, 200, execute_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT amount FROM repayments WHERE contract_id=? AND period=2", (contract_id,)),
            2500.0,
        )

        lease_to_sale_vehicle = self.create_vehicle()
        lease_to_sale_contract = self.create_contract(
            lease_to_sale_vehicle,
            contract_type="以租代售",
            business_mode="以租代售",
        )
        self.db_execute(
            "UPDATE repayments SET status='已还款', paid_amount=amount, verified_amount=amount WHERE contract_id=? AND period>=1",
            (lease_to_sale_contract,),
        )
        self.login("ops")
        transfer_response = self.client.post(
            f"/api/contracts/{lease_to_sale_contract}/ownership-transfer",
            json={"settle_type": "natural_settle", "idempotency_key": "test-transfer-1"},
        )
        self.assertEqual(transfer_response.status_code, 200, transfer_response.get_json())
        transfer_id = transfer_response.get_json()["id"]
        complete_response = self.client.post(
            f"/api/ownership-transfers/{transfer_id}/complete",
            json={
                "transfer_date": "2026-06-02",
                "new_owner_name": "测试过户客户",
                "new_owner_id_card": "610100199001010011",
                "transfer_doc_path": "/uploads/transfer.pdf",
            },
        )
        self.assertEqual(complete_response.status_code, 200, complete_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (lease_to_sale_vehicle,)),
            "已售/已过户",
        )
        self.assertEqual(
            self.db_value("SELECT contract_status FROM contracts WHERE id=?", (lease_to_sale_contract,)),
            "已结清",
        )

    def test_plan_difference_confirmation_is_optional_for_linked_order_activation(self):
        vehicle_id = self.create_vehicle()
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-12",
                "customer_name": "计划差异客户",
                "customer_phone": "13922220000",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"],
                "sale_total_price": 128000,
                "vehicle_rent_amount": 3500,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        self.login("ops")
        plan_response = self.client.put(
            f"/api/sales-orders/{order_id}/planning-contract",
            json={
                "expected_profit_floor": 999999,
                "expected_profit_ceiling": 1000000,
                "start_date": "2026-01-01",
                "loan_periods": 2,
                "rent": 3500,
                "monthly_payment": 2800,
                "factory_periods": 2,
            },
        )
        self.assertEqual(plan_response.status_code, 200, plan_response.get_json())
        contract_id = plan_response.get_json()["contract_id"]
        self.logout()

        self.login("fin")
        activate_response = self.client.post(f"/api/sales-orders/{order_id}/activate", json={})
        self.assertEqual(activate_response.status_code, 200, activate_response.get_json())
        self.logout()

        xlsx_url = self.create_factory_plan_xlsx([
            [1, "2026-02-01", 2800],
            [2, "2026-03-01", 2800],
        ])
        self.login("fin")
        import_response = self.client.post(
            f"/api/contracts/{contract_id}/factory-repayments/import",
            json={"file_url": xlsx_url},
        )
        self.assertEqual(import_response.status_code, 200, import_response.get_json())
        self.assertEqual(import_response.get_json()["comparison"]["status"], "差异待处理")
        confirm_response = self.client.post(
            f"/api/contracts/{contract_id}/plan-compare/confirm-difference",
            json={"comment": "财务确认该利差符合线下审批"},
        )
        self.assertEqual(confirm_response.status_code, 200, confirm_response.get_json())
        self.logout()
        self.assertEqual(
            self.db_value("SELECT customer_plan_match_status FROM contracts WHERE id=?", (contract_id,)),
            "差异已确认",
        )

    def test_finance_cannot_save_customer_installment_plan(self):
        vehicle_id = self.create_vehicle()
        vehicle = self.get_vehicle(vehicle_id)

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-05-12",
                "customer_name": "财务不可保存客户分期",
                "customer_phone": "13922225555",
                "sales_mode": "经营租赁",
                "vin": vehicle["vin"],
                "sale_total_price": 128000,
                "vehicle_rent_amount": 3500,
                "payment_category": "定金",
                "deposit_amount": 3000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())
        order_id = order_response.get_json()["id"]
        self.logout()

        self.login("fin")
        plan_response = self.client.put(
            f"/api/sales-orders/{order_id}/planning-contract",
            json={"loan_periods": 2, "rent": 3500},
        )
        self.assertEqual(plan_response.status_code, 403, plan_response.get_json())
        self.assertIn("客户分期由运营维护", plan_response.get_json()["message"])
        self.logout()

    def test_profit_boss_dashboard_and_sla_visibility(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1",
            (contract_id,),
        )
        factory_repayment_id = self.db_value(
            "SELECT id FROM factory_repayments WHERE contract_id=? ORDER BY period ASC LIMIT 1",
            (contract_id,),
        )

        self.login("fin")
        self.assertEqual(self.client.post(f"/api/repayments/{repayment_id}/confirm", json={}).status_code, 200)
        self.assertEqual(self.client.post(f"/api/factory-repayments/{factory_repayment_id}/confirm", json={}).status_code, 200)
        rebate_response = self.client.post(
            "/api/vehicle-rebates",
            json={"vehicle_id": vehicle_id, "contract_id": contract_id, "rebate_amount": 1200},
        )
        self.assertEqual(rebate_response.status_code, 200, rebate_response.get_json())
        profit_response = self.client.get("/api/profit/by-vehicle")
        self.assertEqual(profit_response.status_code, 200, profit_response.get_json())
        profit_rows = [row for row in profit_response.get_json() if row["contract_id"] == contract_id]
        self.assertEqual(len(profit_rows), 1)
        self.assertIn("planned_single_vehicle_profit", profit_rows[0])
        self.assertIn("realized_cash_profit", profit_rows[0])
        summary_response = self.client.get("/api/profit/summary")
        self.assertEqual(summary_response.status_code, 200, summary_response.get_json())
        self.logout()

        self.login("sales")
        sales_profit = self.client.get("/api/profit/by-vehicle")
        self.assertEqual(sales_profit.status_code, 403, sales_profit.get_json())
        self.logout()

        old_time = (datetime.now() - timedelta(hours=30)).strftime("%Y-%m-%d %H:%M:%S")
        self.db_execute(
            "UPDATE approval_flows SET created_at=? WHERE ref_type='contract_delivery' AND ref_id=?",
            (old_time, contract_id),
        )
        self.login("boss")
        dashboard_response = self.client.get("/api/boss-dashboard/overview")
        self.assertEqual(dashboard_response.status_code, 200, dashboard_response.get_json())
        self.assertIn("realized_cash_profit", dashboard_response.get_json()["overview"])
        sla_response = self.client.get("/api/sla/reminders")
        self.assertEqual(sla_response.status_code, 200, sla_response.get_json())
        self.assertIn("items", sla_response.get_json())
        self.logout()

    def test_dashboard_stats_are_realtime_and_consistent_across_dashboards(self):
        today = datetime.now().date()
        month_start = today.replace(day=1)
        next_month = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1)
        month_due = today + timedelta(days=1)
        if month_due >= next_month:
            month_due = today
        overdue_due = today - timedelta(days=1)

        conn = database.get_db()
        try:
            for table in [
                "reconciliation_allocations", "contract_initial_payments", "factory_repayments",
                "repayments", "contracts", "sales_orders", "vehicles", "customers",
            ]:
                conn.execute(f"DELETE FROM {table}")
            conn.execute("""
                INSERT INTO vehicles
                    (id, vin, plate_number, company, car_type, invoice_date, invoice_price,
                     estimated_residual_value, insurance_expiry_date, status)
                VALUES
                    (101, 'DASHBOARDVIN00001', '陕仪0001', '陕西金聚源汽车服务有限公司', '解放轻卡',
                     ?, 120000, 80000, ?, '租赁中'),
                    (102, 'DASHBOARDVIN00002', '陕仪0002', '陕西金聚源汽车服务有限公司', '解放轻卡',
                     ?, 100000, 70000, NULL, '在库')
            """, (
                today.strftime("%Y-%m-%d"),
                (today + timedelta(days=10)).strftime("%Y-%m-%d"),
                today.strftime("%Y-%m-%d"),
            ))
            conn.execute("INSERT INTO customers (id, name, phone) VALUES (201, '仪表盘客户', '13800000000')")
            conn.execute("""
                INSERT INTO contracts
                    (id, vehicle_id, customer_id, contract_type, contract_status, delivery_status,
                     contract_file, rent, loan_periods)
                VALUES
                    (301, 101, 201, '租赁', '执行中', '已出库', '/uploads/signed.pdf', 3000, 12),
                    (302, 102, 201, '租赁', '已结清', '待首付款', '/uploads/signed.pdf', 3000, 12)
            """)
            conn.execute("""
                INSERT INTO repayments
                    (contract_id, period, due_date, amount, paid_amount, status)
                VALUES
                    (301, 1, ?, 3000, 500, '待还款'),
                    (301, 2, ?, 1000, 0, '待还款'),
                    (301, 0, ?, 999, 0, '未激活'),
                    (302, 1, ?, 888, 0, '待还款')
            """, (
                month_due.strftime("%Y-%m-%d"),
                overdue_due.strftime("%Y-%m-%d"),
                month_due.strftime("%Y-%m-%d"),
                month_due.strftime("%Y-%m-%d"),
            ))
            conn.execute("""
                INSERT INTO factory_repayments (contract_id, period, due_date, amount, status)
                VALUES
                    (301, 1, ?, 700, '待还款'),
                    (301, 2, ?, 300, '已还款')
            """, (month_due.strftime("%Y-%m-%d"), month_due.strftime("%Y-%m-%d")))
            conn.execute("""
                INSERT INTO sales_orders
                    (id, vehicle_id, vin, customer_name, customer_phone, sales_mode, order_status, contract_id)
                VALUES
                    (401, 101, 'DASHBOARDVIN00001', '仪表盘客户', '13800000000', '经营租赁', '待财务确认', NULL),
                    (402, 101, 'DASHBOARDVIN00001', '仪表盘客户', '13800000000', '经营租赁', '已激活', NULL),
                    (403, 101, 'DASHBOARDVIN00001', '仪表盘客户', '13800000000', '经营租赁', '已激活', 301)
            """)
            conn.commit()
        finally:
            conn.close()

        self.login("sales")
        stats_response = self.client.get("/api/dashboard/stats")
        self.assertEqual(stats_response.status_code, 200, stats_response.get_json())
        stats = stats_response.get_json()
        self.logout()

        expected_monthly_due = 2500
        if overdue_due >= month_start:
            expected_monthly_due += 1000
        self.assertEqual(stats["total_vehicles"], 2)
        self.assertEqual(stats["active_contract_count"], 1)
        self.assertEqual(stats["open_order_count"], 2)
        self.assertEqual(stats["active_order_contract_count"], 3)
        self.assertEqual(stats["monthly_due"], expected_monthly_due)
        self.assertEqual(stats["monthly_factory_due"], 700)
        self.assertEqual(stats["overdue_count"], 1)
        self.assertEqual(stats["expiring_insurance_count"], 1)
        self.assertEqual(stats["total_customer_received"], 500)
        self.assertEqual(stats["gross_profit"], 200)
        self.assertEqual(len(stats["asset_chart"]["months"]), 12)

        self.login("boss")
        boss_response = self.client.get("/api/boss-dashboard/overview")
        self.assertEqual(boss_response.status_code, 200, boss_response.get_json())
        overview = boss_response.get_json()["overview"]
        self.assertEqual(overview["active_contract_count"], stats["active_contract_count"])
        self.assertEqual(overview["open_order_count"], stats["open_order_count"])
        self.assertEqual(overview["overdue_repayment_count"], stats["overdue_count"])
        self.assertEqual(overview["customer_received"], stats["total_customer_received"])
        self.assertEqual(overview["factory_paid"], stats["total_factory_paid"])
        self.assertEqual(overview["realized_cash_profit"], stats["realized_cash_profit"])
        self.logout()

    def test_sensitive_endpoints_require_roles(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=1",
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
        sales_screenshot = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/x.png"},
        )
        self.assertEqual(sales_screenshot.status_code, 403, sales_screenshot.get_json())
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

        legal_login = self.client.post(
            "/api/auth/login",
            json={"username": "legal", "password": "123456"},
        )
        self.assertEqual(legal_login.status_code, 401, legal_login.get_json())

        self.login("sales")
        sales_delivery_files = self.client.post(
            f"/api/contracts/{contract_id}/delivery-files",
            json={"delivery_photo_path": "/uploads/not-allowed.png"},
        )
        self.assertEqual(sales_delivery_files.status_code, 403, sales_delivery_files.get_json())
        self.logout()

        self.login("boss")
        boss_delivery_files = self.client.post(
            f"/api/contracts/{contract_id}/delivery-files",
            json={
                "delivery_photo_path": "/uploads/boss-delivery-photo.png",
                "delivery_document_path": "/uploads/boss-delivery-doc.png",
            },
        )
        self.assertEqual(boss_delivery_files.status_code, 200, boss_delivery_files.get_json())
        self.logout()

        self.login("sales")
        sales_me = self.client.get("/api/auth/me")
        self.assertEqual(sales_me.status_code, 200, sales_me.get_json())
        self.assertNotIn("bills", sales_me.get_json()["user"]["pages"])
        self.assertNotIn("reconciliation", sales_me.get_json()["user"]["pages"])
        self.logout()

        self.login("ops")
        ops_me = self.client.get("/api/auth/me")
        self.assertEqual(ops_me.status_code, 200, ops_me.get_json())
        self.assertIn("bills", ops_me.get_json()["user"]["pages"])
        self.assertIn("reconciliation", ops_me.get_json()["user"]["pages"])
        self.logout()

    def test_sales_approval_center_only_shows_return_stock(self):
        vehicle_id = self.create_vehicle()
        contract_id = self.create_contract(vehicle_id)

        self.login("ops")
        ops_types = {row["ref_type"] for row in self.client.get("/api/approvals").get_json()}
        self.assertEqual(ops_types, set())
        self.logout()

        self.login("fin")
        fin_types = {row["ref_type"] for row in self.client.get("/api/approvals").get_json()}
        self.assertEqual(fin_types, set())
        self.logout()

        self.login("boss")
        boss_types = {row["ref_type"] for row in self.client.get("/api/approvals").get_json()}
        self.assertNotIn("contract_delivery", boss_types)
        self.logout()

        self.login("sales")
        approvals_response = self.client.get("/api/approvals")
        self.assertEqual(approvals_response.status_code, 200, approvals_response.get_json())
        self.assertEqual(approvals_response.get_json(), [])
        self.logout()

        self.login("sales")
        in_stock_return_response = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "到期退车"},
        )
        self.assertEqual(in_stock_return_response.status_code, 400, in_stock_return_response.get_json())
        self.logout()

        return_vehicle_id, _ = self.prepare_delivered_rental_vehicle()
        self.login("sales")
        inspection_response = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": return_vehicle_id, "return_reason": "到期退车"},
        )
        self.assertEqual(inspection_response.status_code, 200, inspection_response.get_json())
        inspection_id = inspection_response.get_json()["id"]
        self.logout()

        self.login("sales")
        approvals_response = self.client.get("/api/approvals")
        self.assertEqual(approvals_response.status_code, 200, approvals_response.get_json())
        approval_types = {row["ref_type"] for row in approvals_response.get_json()}
        self.assertEqual(approval_types, set())
        filtered_response = self.client.get("/api/approvals?ref_type=contract_delivery")
        self.assertEqual(filtered_response.status_code, 200, filtered_response.get_json())
        self.assertEqual({row["ref_type"] for row in filtered_response.get_json()}, set())
        self.logout()


if __name__ == "__main__":
    unittest.main(verbosity=2)
