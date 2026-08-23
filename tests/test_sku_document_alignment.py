import os
import shutil
import tempfile
import unittest

import app as app_module
import database


class SkuDocumentAlignmentTestCase(unittest.TestCase):
    """V3.2 SKU 文档：字典、车型计算、VIN 报单规格必须保持一致。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-sku-alignment-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "sku-alignment.db")
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

    def db_row(self, sql, params=()):
        conn = database.get_db()
        try:
            row = conn.execute(sql, params).fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def test_dictionary_natural_key_is_deduplicated_and_seed_is_idempotent(self):
        conn = database.get_db()
        try:
            conn.execute("DROP INDEX idx_data_dictionaries_natural_key")
            conn.execute(
                """
                INSERT INTO data_dictionaries (category, value, energy_type, status)
                VALUES (' box_type ', ' 冷藏 ', '纯电', '停用')
                """
            )
            conn.execute(
                """
                INSERT INTO data_dictionaries (category, value, energy_type, status)
                VALUES ('box_type', '冷藏', '纯电', '启用')
                """
            )
            conn.commit()
        finally:
            conn.close()

        database.init_db()
        conn = database.get_db()
        try:
            app_module.seed_data_dictionaries(conn)
            app_module.seed_data_dictionaries(conn)
            duplicate_count = conn.execute(
                """
                SELECT COUNT(*) AS count
                FROM data_dictionaries
                WHERE category='box_type' AND value='冷藏' AND energy_type='纯电'
                """
            ).fetchone()["count"]
            status = conn.execute(
                """
                SELECT status FROM data_dictionaries
                WHERE category='box_type' AND value='冷藏' AND energy_type='纯电'
                """
            ).fetchone()["status"]
            totals = conn.execute(
                """
                SELECT COUNT(*) AS total,
                       COUNT(DISTINCT category || char(31) || value || char(31) || energy_type) AS unique_total
                FROM data_dictionaries
                """
            ).fetchone()
        finally:
            conn.close()

        self.assertEqual(duplicate_count, 1)
        self.assertEqual(status, "启用")
        self.assertEqual(totals["total"], totals["unique_total"])

    def test_sales_vehicle_list_classifies_legacy_electric_vehicle_and_seeds_its_capacity(self):
        """历史导入车缺燃料形式时，销售端四维筛选仍应可识别为纯电。"""
        legacy_vin = "LEGACYFILTER000001"
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles
                    (vin, plate_number, car_type, condition, status, box_type,
                     battery_brand, battery_capacity)
                VALUES (?, '陕筛134度', '二手车零米110kw快递版134度冷藏',
                        '二手车', '在库', '冷藏', '宁德', '134度')
                """,
                (legacy_vin,),
            )
            conn.commit()
            app_module.seed_data_dictionaries(conn)
            dict_row = conn.execute(
                """
                SELECT value, energy_type
                FROM data_dictionaries
                WHERE category='battery_capacity' AND value='134度'
                """
            ).fetchone()
        finally:
            conn.close()

        self.assertIsNotNone(dict_row)
        self.assertEqual(dict_row["energy_type"], "纯电")

        self.login("sales")
        response = self.client.get("/api/vehicles/list")
        self.assertEqual(response.status_code, 200, response.get_json())
        row = next(item for item in response.get_json() if item["vin"] == legacy_vin)
        self.assertEqual(row["condition"], "二手车")
        self.assertEqual(row["box_type"], "冷藏")
        self.assertEqual(row["battery_capacity"], "134度")
        self.assertEqual(row["energy_type"], "纯电")

    def test_guidance_workbench_lists_complete_submodels_and_scoped_finance_plans(self):
        base_type = "解放J6F锡柴170"
        self.login("boss")
        guidance = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": f"新车{base_type}厢货",
                "is_new": "新车",
                "lease_deposit_guidance": 2200,
                "box_standard_price": 3000,
                "box_refrigerated_price": 3500,
                "tail_plate_price": 300,
            },
        )
        self.assertEqual(guidance.status_code, 200, guidance.get_json())

        conn = database.get_db()
        try:
            conn.executemany(
                """
                INSERT INTO vehicles (vin, car_type, condition, box_type, status)
                VALUES (?, ?, ?, ?, '在库')
                """,
                [
                    ("WORKBENCH000000001", f"新车{base_type}厢货", "新车", "厢货"),
                    ("WORKBENCH000000002", f"新车{base_type}冷藏", "新车", "冷藏"),
                    ("WORKBENCH000000003", f"二手车{base_type}高栏", "二手车", "高栏"),
                ],
            )
            conn.commit()
        finally:
            conn.close()

        plan = self.client.post(
            "/api/finance-plans",
            json={
                "car_type": f"新车{base_type}冷藏",
                "plan_name": "冷藏36期",
                "down_payment": 30000,
                "period_price": 3600,
                "periods": 36,
                "status": "启用",
            },
        )
        self.assertEqual(plan.status_code, 200, plan.get_json())

        response = self.client.get(
            "/api/model-guidance-workbench",
            query_string={"car_type": f"二手车{base_type}高栏"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        data = response.get_json()
        self.assertTrue(data["success"])
        canonical_base = base_type + "无尾板"
        self.assertEqual(data["car_type"], canonical_base)
        self.assertEqual(len(data["lease_rows"]), 10)
        self.assertEqual(data["summary"]["total_count"], 10)
        self.assertEqual(data["summary"]["configured_count"], 2)
        self.assertEqual(data["summary"]["missing_count"], 8)
        self.assertEqual(data["summary"]["rent_to_buy_total_count"], 5)
        self.assertEqual(data["summary"]["rent_to_buy_configured_count"], 1)
        self.assertEqual(data["summary"]["rent_to_buy_plan_count"], 1)

        rows = {
            (row["condition"], row["box_type"]): row
            for row in data["lease_rows"]
        }
        self.assertEqual(rows[("新车", "厢货")]["inventory_count"], 1)
        self.assertEqual(rows[("新车", "冷藏")]["inventory_count"], 1)
        self.assertEqual(rows[("二手车", "高栏")]["inventory_count"], 1)
        self.assertEqual(rows[("新车", "厢货")]["status"], "已设置")
        self.assertEqual(rows[("新车", "冷藏")]["status"], "已设置")
        self.assertEqual(rows[("新车", "宽体")]["status"], "待设置")
        self.assertEqual(rows[("二手车", "高栏")]["status"], "待设置")
        self.assertEqual(rows[("新车", "冷藏")]["monthly_price"], 3500)
        self.assertEqual(rows[("二手车", "高栏")]["lease_deposit_guidance"], 0)
        self.assertEqual(
            [(item["car_type"], item["plan_name"]) for item in data["finance_plans"]],
            [(canonical_base, "冷藏36期")],
        )
        plan_rows = {item["box_type"]: item for item in data["rent_to_buy_rows"]}
        self.assertEqual(set(plan_rows), {"厢货", "宽体", "高栏", "冷藏", "平板"})
        self.assertEqual(plan_rows["冷藏"]["plan_count"], 1)
        self.assertEqual(plan_rows["冷藏"]["active_plan_count"], 1)
        self.assertEqual(plan_rows["冷藏"]["status"], "已设置")
        self.assertEqual(plan_rows["厢货"]["plan_count"], 0)
        self.assertEqual(plan_rows["厢货"]["status"], "待设置")

        cold_plans = self.client.get(
            "/api/finance-plans",
            query_string={"car_type": canonical_base, "condition": "新车", "box_type": "冷藏"},
        ).get_json()
        cargo_plans = self.client.get(
            "/api/finance-plans",
            query_string={"car_type": canonical_base, "condition": "新车", "box_type": "厢货"},
        ).get_json()
        self.assertEqual([item["id"] for item in cold_plans], [plan.get_json()["id"]])
        self.assertEqual(cargo_plans, [])
        conn = database.get_db()
        try:
            self.assertIsNotNone(
                app_module.resolve_finance_plan(
                    conn, plan.get_json()["id"], car_type=canonical_base,
                    condition="新车", box_type="冷藏",
                )
            )
            self.assertIsNone(
                app_module.resolve_finance_plan(
                    conn, plan.get_json()["id"], car_type=canonical_base,
                    condition="新车", box_type="厢货",
                )
            )
        finally:
            conn.close()

        stored_plan = self.db_row(
            "SELECT car_type, condition, box_type FROM finance_plans WHERE id=?",
            (plan.get_json()["id"],),
        )
        self.assertEqual(stored_plan["car_type"], canonical_base)
        self.assertEqual(stored_plan["condition"], "新车")
        self.assertEqual(stored_plan["box_type"], "冷藏")

    def test_vehicle_model_is_computed_and_tailgate_creates_tailgate_sku(self):
        self.login("boss")
        guidance = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": "解放J6F锡柴150有尾板",
                "is_new": "新车",
                "lease_deposit_guidance": 2000,
                "box_refrigerated_price": 3500,
                "tail_plate_price": 300,
            },
        )
        self.assertEqual(guidance.status_code, 200, guidance.get_json())

        self.login("fleet")
        created = self.client.post(
            "/api/vehicles",
            json={
                "vin": "SKU20260811000001",
                "condition": "新车",
                "brand": "解放",
                "product_series": "J6F",
                "fuel_form": "柴油",
                "horsepower": "锡柴150",
                "gear_position": "八档",
                "box_type": "冷藏",
                "tailgate": "带尾板",
                "car_type": "不应被信任的车型",
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())

        vehicle = self.db_row(
            "SELECT car_type, tailgate FROM vehicles WHERE vin=?",
            ("SKU20260811000001",),
        )
        sku = self.db_row(
            """
            SELECT id, car_type, condition, box_type, tailgate
            FROM skus
            WHERE car_type=? AND condition=? AND box_type=? AND tailgate=?
            """,
            ("解放J6F锡柴150八档有尾板", "新车", "冷藏", "有"),
        )
        self.assertEqual(vehicle["car_type"], "解放J6F锡柴150八档有尾板")
        self.assertEqual(vehicle["tailgate"], "有")
        self.assertIsNotNone(sku)
        inventory = self.client.get("/api/skus/inventory")
        self.assertEqual(inventory.status_code, 200, inventory.get_json())
        matched_inventory = next(
            item for item in inventory.get_json()
            if item["sku_id"] == sku["id"]
        )
        self.assertEqual(matched_inventory["in_stock_count"], 1)

    def test_sales_order_uses_actual_vin_box_tailgate_and_car_type(self):
        self.login("boss")
        guidance = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": "解放J6F锡柴150八档有尾板",
                "is_new": "新车",
                "lease_deposit_guidance": 2000,
                "box_refrigerated_price": 3500,
                "box_standard_price": 3000,
                "tail_plate_price": 300,
            },
        )
        self.assertEqual(guidance.status_code, 200, guidance.get_json())

        self.login("fleet")
        created = self.client.post(
            "/api/vehicles",
            json={
                "vin": "SKU20260811000001",
                "condition": "新车",
                "brand": "解放",
                "product_series": "J6F",
                "fuel_form": "柴油",
                "horsepower": "锡柴150",
                "gear_position": "八档",
                "box_type": "冷藏",
                "tailgate": "带尾板",
            },
        )
        self.assertEqual(created.status_code, 200, created.get_json())

        self.login("sales")
        order_response = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-08-11",
                "customer_name": "SKU 文档客户",
                "customer_phone": "13800000000",
                "sales_mode": "租赁",
                "vin": "SKU20260811000001",
                "car_type": "伪造车型",
                "vehicle_box_type": "厢货",
                "tail_plate": "无",
                "lease_term": "12期",
                "deposit_amount": 2000,
                "vehicle_rent_amount": 3000,
                "customer_screenshot_path": "/uploads/sku-first-payment.jpg",
                "first_payment_received_amount": 5000,
            },
        )
        self.assertEqual(order_response.status_code, 200, order_response.get_json())

        order = self.db_row(
            """
            SELECT car_type, vehicle_box_type, tail_plate, order_status, price_exception_reason
            FROM sales_orders WHERE id=?
            """,
            (order_response.get_json()["id"],),
        )
        self.assertEqual(order["car_type"], "解放J6F锡柴150八档有尾板")
        self.assertEqual(order["vehicle_box_type"], "冷藏")
        self.assertEqual(order["tail_plate"], "有")
        self.assertEqual(order["order_status"], "待老板审批")
        self.assertIn("指导月供 ¥3800.0", order["price_exception_reason"])

    def test_base_model_fields_follow_energy_type_and_exclude_child_dimensions(self):
        electric = {
            "fuel_form": "纯电",
            "brand": "解放",
            "product_series": "虎6G",
            "battery_brand": "宁德",
            "battery_capacity": "100度",
            "condition": "新车",
            "box_type": "冷藏",
            "tailgate": "有",
        }
        fuel = {
            "fuel_form": "柴油",
            "brand": "解放",
            "product_series": "J6F",
            "horsepower": "锡柴150",
            "gear_position": "八档",
            "condition": "二手车",
            "box_type": "高栏",
            "tailgate": "无",
        }

        self.assertEqual(
            app_module.compute_car_type(electric),
            "解放虎6G宁德100度有尾板",
        )
        self.assertEqual(
            app_module.compute_car_type(fuel),
            "解放J6F锡柴150八档无尾板",
        )
        self.assertEqual(
            app_module.normalize_base_car_type("新车解放J6F锡柴150八档冷藏带尾板"),
            "解放J6F锡柴150八档有尾板",
        )
        self.assertIn(
            "纯电车型不应填写“马力”",
            app_module.base_model_validation_errors({**electric, "horsepower": "150马力"})[0],
        )

    def test_guidance_cleanup_only_targets_unreferenced_numeric_legacy_names(self):
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO model_guidance_prices (car_type, is_new, guidance_price)
                VALUES ('100宁德虎六G', '新车', 100000)
                """
            )
            conn.execute(
                """
                INSERT INTO model_guidance_prices (car_type, is_new, guidance_price)
                VALUES ('解放J6F锡柴150八档无尾板', '新车', 100000)
                """
            )
            conn.commit()
            rows = conn.execute(
                """
                SELECT car_type
                FROM model_guidance_prices p
                WHERE p.car_type GLOB '[0-9]*'
                  AND NOT EXISTS (
                      SELECT 1 FROM vehicles v
                      WHERE (v.is_deleted IS NULL OR v.is_deleted=0)
                        AND v.car_type=p.car_type
                  )
                  AND NOT EXISTS (
                      SELECT 1 FROM sales_orders o WHERE o.car_type=p.car_type
                  )
                  AND NOT EXISTS (
                      SELECT 1 FROM finance_plans f WHERE f.car_type=p.car_type
                  )
                """
            ).fetchall()
        finally:
            conn.close()

        self.assertEqual([row["car_type"] for row in rows], ["100宁德虎六G"])

    def test_dictionary_enforces_energy_scope_and_hides_retired_categories(self):
        self.login("boss")
        created = self.client.post(
            "/api/data-dictionaries",
            json={"category": "battery_capacity", "value": "100度", "energy_type": "纯电"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())

        bad_energy = self.client.post(
            "/api/data-dictionaries",
            json={"category": "battery_capacity", "value": "120度", "energy_type": "燃油车"},
        )
        self.assertEqual(bad_energy.status_code, 400, bad_energy.get_json())

        bad_common = self.client.post(
            "/api/data-dictionaries",
            json={"category": "brand", "value": "解放", "energy_type": "纯电"},
        )
        self.assertEqual(bad_common.status_code, 400, bad_common.get_json())

        retired = self.client.post(
            "/api/data-dictionaries",
            json={"category": "box_dimension", "value": "4.2米", "energy_type": ""},
        )
        self.assertEqual(retired.status_code, 400, retired.get_json())

        conn = database.get_db()
        try:
            conn.execute(
                """INSERT INTO data_dictionaries (category, value, energy_type, status)
                   VALUES ('box_dimension', '4.2米', '', '启用')"""
            )
            conn.execute(
                """INSERT INTO data_dictionaries (category, value, energy_type, status)
                   VALUES ('vehicle_color', '白色', '', '启用')"""
            )
            conn.execute(
                """INSERT INTO data_dictionaries (category, value, energy_type, status)
                   VALUES ('cab_type', '中体单排', '', '启用')"""
            )
            conn.commit()
        finally:
            conn.close()
        rows = self.client.get("/api/data-dictionaries").get_json()
        hidden_categories = {"box_dimension", "vehicle_color", "cab_type"}
        self.assertFalse(any(row["category"] in hidden_categories for row in rows))

    def test_manual_vehicle_rejects_cross_energy_base_fields(self):
        self.login("fleet")
        pure_electric_with_fuel_fields = self.client.post(
            "/api/vehicles",
            json={
                "vin": "ENERGY20260811001",
                "fuel_form": "纯电",
                "brand": "解放",
                "product_series": "虎6G",
                "battery_brand": "宁德",
                "battery_capacity": "100度",
                "horsepower": "150马力",
            },
        )
        self.assertEqual(pure_electric_with_fuel_fields.status_code, 400)
        self.assertIn("不应填写", pure_electric_with_fuel_fields.get_json()["message"])

        fuel_with_electric_fields = self.client.post(
            "/api/vehicles",
            json={
                "vin": "ENERGY20260811002",
                "fuel_form": "柴油",
                "brand": "解放",
                "product_series": "J6F",
                "horsepower": "150马力",
                "gear_position": "八档",
                "battery_brand": "宁德",
            },
        )
        self.assertEqual(fuel_with_electric_fields.status_code, 400)
        self.assertIn("不应填写", fuel_with_electric_fields.get_json()["message"])

    def test_boss_cannot_change_imported_fuel_form_through_vehicle_update(self):
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles (
                    vin, car_type, fuel_form, brand, product_series, horsepower,
                    gear_position, tailgate, condition, box_type, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '在库')
                """,
                (
                    "LOCKEDFUEL20260801",
                    "解放J6F锡柴150八档无尾板",
                    "柴油",
                    "解放",
                    "J6F",
                    "锡柴150",
                    "八档",
                    "无",
                    "新车",
                    "厢货",
                ),
            )
            vehicle_id = conn.execute(
                "SELECT id FROM vehicles WHERE vin=?",
                ("LOCKEDFUEL20260801",),
            ).fetchone()["id"]
            conn.commit()
        finally:
            conn.close()

        self.login("boss")
        response = self.client.put(
            f"/api/vehicles/{vehicle_id}",
            json={"fuel_form": "纯电"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        stored = self.db_row(
            "SELECT fuel_form, car_type FROM vehicles WHERE id=?",
            (vehicle_id,),
        )
        self.assertEqual(stored["fuel_form"], "柴油")
        self.assertEqual(stored["car_type"], "解放J6F锡柴150八档无尾板")


if __name__ == "__main__":
    unittest.main()
