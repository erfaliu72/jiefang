# -*- coding: utf-8 -*-
"""2026-09-16 测试计划方案 API-01~16 接口扩展用例。"""
import io
import json
import os
import shutil
import tempfile
import unittest

from openpyxl import Workbook

import app as app_module
import database


class Plan20260916ApiTestCase(unittest.TestCase):
    """所有用例使用独立临时库，setUp 重建 schema 并 seed 五角色。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="jiefang-plan-20260916-")
        self.original_database = database.DATABASE
        database.DATABASE = os.path.join(self.temp_dir, "plan.db")
        database.init_db()
        database.seed_data()
        app_module.app.config["TESTING"] = True
        self.client = app_module.app.test_client()
        self.vehicle_seq = 0

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

    def db_rows(self, sql, params=()):
        conn = database.get_db()
        try:
            return [dict(row) for row in conn.execute(sql, params).fetchall()]
        finally:
            conn.close()

    def db_value(self, sql, params=()):
        row = self.db_row(sql, params)
        return next(iter(row.values())) if row else None

    def db_exec(self, sql, params=()):
        conn = database.get_db()
        try:
            conn.execute(sql, params)
            conn.commit()
        finally:
            conn.close()

    def create_second_sales(self):
        self.db_exec(
            "INSERT INTO users (username, password, display_name, role) VALUES (?, ?, ?, ?)",
            ("sales2", "123456", "李销售", "销售"),
        )

    def create_vehicle(self, car_type, status="在库", condition="新车", purchase_price=90000):
        self.vehicle_seq += 1
        vin = f"W{self.vehicle_seq:016d}"
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO vehicles
                    (vin, plate_number, car_type, condition, status, box_type,
                     vehicle_box_type, purchase_price, validation_status, lock_status)
                VALUES (?, ?, ?, ?, ?, '厢货', '厢货', ?, 'valid', '未锁')
                """,
                (vin, f"陕测{self.vehicle_seq:04d}", car_type, condition, status, purchase_price),
            )
            vehicle_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return vehicle_id, vin
        finally:
            conn.close()

    def save_lease_guidance(self, car_type, deposit=2000, monthly=3000):
        self.login("boss")
        response = self.client.post(
            "/api/model-guidance-prices",
            json={
                "car_type": car_type,
                "is_new": "新车",
                "lease_deposit_guidance": deposit,
                "box_standard_price": monthly,
                "tail_plate_price": 300,
            },
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def create_order(self, vin, car_type, username="sales", **overrides):
        payload = {
            "payment_date": "2026-09-16",
            "customer_name": "测试计划客户",
            "customer_phone": "13800000000",
            "sales_mode": "租赁",
            "vin": vin,
            "car_type": car_type,
            "vehicle_box_type": "厢货",
            "lease_term": "3期",
            "deposit_amount": 2000,
            "vehicle_rent_amount": 3000,
            "customer_screenshot_path": "/uploads/customer-first-payment.jpg",
            "first_payment_received_amount": 5000,
        }
        payload.update(overrides)
        self.login(username)
        response = self.client.post("/api/sales-orders", json=payload)
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()["id"]

    def submit_order(self, order_id, username="sales"):
        self.login(username)
        response = self.client.put(
            f"/api/sales-orders/{order_id}",
            json={"action": "submit"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()

    def confirm_intent_deposit(self, order_id, serial="INTENT20260916001", username="fin", receipt_path=""):
        self.login(username)
        payload = {"serial": serial}
        if receipt_path:
            payload["receipt_path"] = receipt_path
        response = self.client.post(
            f"/api/sales-orders/{order_id}/confirm-intent-deposit",
            json=payload,
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        return response.get_json()

    def create_active_contract(self, end_date="2026-12-31", created_by="周销售",
                               deposit=2000, collected=2000, rent=3000):
        vehicle_id, _ = self.create_vehicle("续租测试车型", "租赁中")
        conn = database.get_db()
        try:
            conn.execute(
                "INSERT INTO customers (name, phone) VALUES ('续租测试客户', '13900000000')"
            )
            customer_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO contracts
                    (vehicle_id, customer_id, contract_type, start_date, end_date, rent,
                     loan_periods, deposit, collected_deposit, billing_cycle,
                     contract_file, contract_status, delivery_status,
                     lease_bank_name, lease_bank_card_no, created_by)
                VALUES (?, ?, '租赁', '2026-01-01', ?, ?, 12, ?, ?, '按月',
                        '/uploads/accepted-contract.pdf', '执行中', '已出库',
                        '测试银行', '6222000012345678', ?)
                """,
                (vehicle_id, customer_id, end_date, rent, deposit, collected, created_by),
            )
            contract_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
            return vehicle_id, contract_id
        finally:
            conn.close()

    def deliver_order_pipeline(self, order_id, vehicle_id):
        self.login("fin")
        activation = self.client.post(
            f"/api/sales-orders/{order_id}/activate",
            json={"bank_serial": f"ACT{order_id:08d}", "bank_receipt_path": "/uploads/company-receipt.pdf"},
        )
        self.assertEqual(activation.status_code, 200, activation.get_json())

        self.login("ops")
        upload = self.client.post(
            "/api/contracts",
            json={
                "sales_order_id": order_id,
                "vehicle_id": vehicle_id,
                "contract_type": "租赁",
                "customer_name": "测试计划客户",
                "customer_phone": "13800000000",
                "contract_file": "/uploads/signed-contract.pdf",
            },
        )
        self.assertEqual(upload.status_code, 200, upload.get_json())
        contract_id = upload.get_json()["id"]

        self.login("fleet")
        saved = self.client.post(
            f"/api/contracts/{contract_id}/delivery-files",
            json={
                "delivery_photo_path": "/uploads/delivery-photo.jpg",
                "delivery_document_path": "/uploads/delivery-document.jpg",
            },
        )
        self.assertEqual(saved.status_code, 200, saved.get_json())
        delivered = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(delivered.status_code, 200, delivered.get_json())
        return contract_id

    def approve_flow(self, ref_type, ref_id, role_user="boss"):
        flow = self.db_row(
            """
            SELECT id FROM approval_flows
            WHERE ref_type=? AND ref_id=? AND status='待审批'
            ORDER BY id DESC LIMIT 1
            """,
            (ref_type, ref_id),
        )
        self.assertIsNotNone(flow, f"缺少 {ref_type}#{ref_id} 待审批流")
        self.login(role_user)
        response = self.client.post(
            f"/api/approvals/{flow['id']}/approve",
            json={"comment": "测试计划审批通过"},
        )
        self.assertEqual(response.status_code, 200, response.get_json())

    def create_and_activate_renewal(self, contract_id, **payload):
        self.login("sales")
        created = self.client.post(f"/api/contracts/{contract_id}/renewals", json=payload)
        self.assertEqual(created.status_code, 200, created.get_json())
        renewal_id = created.get_json()["id"]
        self.approve_flow("renewal", renewal_id)
        self.login("ops")
        activated = self.client.post(
            f"/api/renewals/{renewal_id}/activate",
            json={
                "contract_file": "/uploads/renewal-contract.pdf",
                "contract_number": f"XZ-2026-{renewal_id:04d}",
            },
        )
        self.assertEqual(activated.status_code, 200, activated.get_json())
        return renewal_id

    # ---------------- API-01 报单锁定不设期限 ----------------
    def test_api01_draft_lock_has_no_expiry(self):
        car_type = "API01锁车车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)

        # 模拟 45 天前创建的历史草稿
        self.db_exec(
            "UPDATE sales_orders SET created_at='2026-08-01 09:00:00' WHERE id=?",
            (order_id,),
        )
        self.login("sales")
        listing = self.client.get("/api/sales-orders")
        self.assertEqual(listing.status_code, 200, listing.get_json())
        row = next(r for r in listing.get_json() if r["id"] == order_id)
        self.assertEqual(row["order_status"], "待提交", "历史草稿不应被自动过期")
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "报单锁定中",
            "锁定不设期限，车辆应保持锁定",
        )
        self.submit_order(order_id)
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "待财务确认",
        )

    # ---------------- API-02 多车（SKU）报单 ----------------
    def test_api02_multi_vehicle_order_and_conflicts(self):
        car_type = "API02多车车型"
        _, vin1 = self.create_vehicle(car_type)
        _, vin2 = self.create_vehicle(car_type)
        order_id = self.create_order(vin1, car_type, vehicle_ids=[vin1, vin2])
        order = self.db_row("SELECT vehicle_ids, order_status FROM sales_orders WHERE id=?", (order_id,))
        self.assertEqual(order["order_status"], "待提交")
        self.assertEqual(len(json.loads(order["vehicle_ids"])), 2)
        locked = self.db_value(
            "SELECT COUNT(*) FROM vehicles WHERE vin IN (?, ?) AND status='报单锁定中'",
            (vin1, vin2),
        )
        self.assertEqual(locked, 2)

        # 冲突一：车辆存在未结清合同
        v4, vin4 = self.create_vehicle(car_type)
        self.db_exec(
            "INSERT INTO contracts (vehicle_id, contract_type, contract_status) VALUES (?, '租赁', '执行中')",
            (v4,),
        )
        self.login("sales")
        conflict = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-09-16",
                "customer_name": "冲突客户",
                "customer_phone": "13700000000",
                "sales_mode": "租赁",
                "vin": vin4,
                "vehicle_ids": [vin4],
                "car_type": car_type,
                "customer_screenshot_path": "/uploads/pay.jpg",
            },
        )
        self.assertEqual(conflict.status_code, 400, conflict.get_json())
        self.assertIn("未结清合同", conflict.get_json()["message"])

        # 冲突二：他人（李销售）锁定中的车辆不能二次报单
        self.create_second_sales()
        _, vin5 = self.create_vehicle(car_type)
        self.create_order(vin5, car_type, username="sales2",
                          customer_name="李销售客户", customer_phone="13600000000")
        self.login("sales")
        dup = self.client.post(
            "/api/sales-orders",
            json={
                "payment_date": "2026-09-16",
                "customer_name": "抢单客户",
                "customer_phone": "13500000000",
                "sales_mode": "租赁",
                "vin": vin5,
                "vehicle_ids": [vin5],
                "car_type": car_type,
                "customer_screenshot_path": "/uploads/pay.jpg",
            },
        )
        self.assertEqual(dup.status_code, 400, dup.get_json())
        self.assertIn("二次报单", dup.get_json()["message"])

    # ---------------- API-03 黑名单客户报单 ----------------
    def test_api03_blacklist_order_routes_to_boss(self):
        car_type = "API03黑名单车型"
        self.save_lease_guidance(car_type)
        self.login("boss")
        created = self.client.post(
            "/api/customer-blacklist",
            json={"customer_name": "黑名单客户", "customer_phone": "13400000000",
                  "reason": "恶意拖欠三期租金"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())

        _, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin, car_type, customer_name="黑名单客户", customer_phone="13400000000",
        )
        order = self.db_row(
            "SELECT blacklist_hit, order_exception_reason FROM sales_orders WHERE id=?",
            (order_id,),
        )
        self.assertEqual(order["blacklist_hit"], 1)
        self.assertIn("恶意拖欠三期租金", order["order_exception_reason"])
        self.assertIn("标记时间", order["order_exception_reason"])

        self.submit_order(order_id)
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "待老板审批",
            "黑名单客户报单仍可提交，但必须自动转老板审批",
        )
        flow = self.db_row(
            "SELECT required_role FROM approval_flows "
            "WHERE ref_type='order_exception' AND ref_id=? AND status='待审批'",
            (order_id,),
        )
        self.assertIsNotNone(flow)
        self.assertEqual(flow["required_role"], "老板")

    # ---------------- API-04 取消订单边界 ----------------
    def test_api04_cancel_order_boundaries(self):
        car_type = "API04取消车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        self.submit_order(order_id)

        # 他人名下报单 → 403
        self.create_second_sales()
        self.login("sales2")
        forbidden = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "他人尝试取消"},
        )
        self.assertEqual(forbidden.status_code, 403, forbidden.get_json())

        # 缺取消原因 → 400
        self.login("sales")
        no_reason = self.client.post("/api/order-refunds", json={"sales_order_id": order_id})
        self.assertEqual(no_reason.status_code, 400, no_reason.get_json())

        # 正常发起 → 取消审批中，应退=全额定金
        created = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "客户放弃租赁"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        refund_id = created.get_json()["id"]
        self.assertEqual(
            self.db_value("SELECT refund_amount FROM order_refunds WHERE id=?", (refund_id,)),
            2000,
        )
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "取消审批中",
        )

        # 重复发起 → 400
        duplicate = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "重复取消"},
        )
        self.assertEqual(duplicate.status_code, 400, duplicate.get_json())

        # 老板审批 → 待退款
        self.login("boss")
        approved = self.client.post(f"/api/order-refunds/{refund_id}/approve", json={})
        self.assertEqual(approved.status_code, 200, approved.get_json())

        # 财务退款：金额不符 → 400；流水号过短 → 400；正确 → 已完成并解锁车辆
        self.login("fin")
        wrong_amount = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND20260916001", "refund_paid_amount": 1500},
        )
        self.assertEqual(wrong_amount.status_code, 400, wrong_amount.get_json())
        short_serial = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "ab", "refund_paid_amount": 2000},
        )
        self.assertEqual(short_serial.status_code, 400, short_serial.get_json())
        paid = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND20260916001", "refund_paid_amount": 2000},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "已作废",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "在库",
            "退款完成后车辆应解除锁定回到在库",
        )

        # 已出库订单不能取消
        vehicle2_id, vin2 = self.create_vehicle(car_type)
        order2_id = self.create_order(vin2, car_type, customer_phone="13300000000")
        self.submit_order(order2_id)
        self.deliver_order_pipeline(order2_id, vehicle2_id)
        self.login("sales")
        delivered_cancel = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order2_id, "reason": "已出库尝试取消"},
        )
        self.assertEqual(delivered_cancel.status_code, 400, delivered_cancel.get_json())
        self.assertIn("退车流程", delivered_cancel.get_json()["message"])

    def test_api04b_cancel_intent_only_order_refunds_full_intent_deposit(self):
        """仅支付意向定金（未补押金/首付款）的意向单取消：全额退意向定金。"""
        car_type = "API04B意向定金车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin, car_type,
            intent_deposit_amount=1500,
            deposit_amount=8000,
            first_payment_received_amount=0,
            first_payment_shortage_reason="客户仅付意向定金，首付未补",
            first_payment_promised_date="2026-09-30",
        )
        self.confirm_intent_deposit(order_id)
        self.submit_order(order_id)

        self.login("sales")
        created = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "客户仅付定金后放弃"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        refund_id = created.get_json()["id"]
        self.assertEqual(
            self.db_value("SELECT refund_amount FROM order_refunds WHERE id=?", (refund_id,)),
            1500,
            "仅意向阶段的取消应退全额意向定金，而不是押金/首付款",
        )

        self.login("boss")
        approved = self.client.post(f"/api/order-refunds/{refund_id}/approve", json={})
        self.assertEqual(approved.status_code, 200, approved.get_json())

        self.login("fin")
        wrong_amount = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND20260916002", "refund_paid_amount": 8000},
        )
        self.assertEqual(wrong_amount.status_code, 400, wrong_amount.get_json())
        paid = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND20260916002", "refund_paid_amount": 1500},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        self.assertEqual(
            self.db_value("SELECT order_status FROM sales_orders WHERE id=?", (order_id,)),
            "已作废",
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "在库",
        )

        # 已补款（实收>0）的订单维持押金/首付款口径，不受意向定金影响
        vehicle2_id, vin2 = self.create_vehicle(car_type)
        order2_id = self.create_order(
            vin2, car_type,
            customer_phone="13400000000",
            intent_deposit_amount=1500,
            deposit_amount=2000,
            first_payment_received_amount=5000,
        )
        self.confirm_intent_deposit(order2_id, serial="INTENT20260916002")
        self.submit_order(order2_id)
        self.login("sales")
        created2 = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order2_id, "reason": "已补款后取消"},
        )
        self.assertEqual(created2.status_code, 200, created2.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT refund_amount FROM order_refunds WHERE id=?",
                (created2.get_json()["id"],),
            ),
            2000,
            "已补款订单应退金额仍按押金/首付款口径",
        )

    def test_api04c_intent_deposit_finance_confirmation_gating(self):
        """意向定金财务确认闸口：未确认不能提交/不能取消；销售无权确认；确认后放行；改金额旧确认作废。"""
        car_type = "API04C定金确认车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(
            vin, car_type,
            intent_deposit_amount=1200,
            deposit_amount=8000,
            first_payment_received_amount=0,
            first_payment_shortage_reason="客户仅付意向定金，首付未补",
            first_payment_promised_date="2026-09-30",
        )
        self.login("fin")
        finance_orders = self.client.get("/api/sales-orders")
        self.assertEqual(finance_orders.status_code, 200, finance_orders.get_json())
        finance_order = next(row for row in finance_orders.get_json() if row["id"] == order_id)
        self.assertEqual(
            finance_order["customer_screenshot_path"],
            "/uploads/customer-first-payment.jpg",
            "财务端列表必须返回销售上传的客户定金支付凭证",
        )

        # 未确认：提交报单被拦
        self.login("sales")
        blocked_submit = self.client.put(f"/api/sales-orders/{order_id}", json={"action": "submit"})
        self.assertEqual(blocked_submit.status_code, 400, blocked_submit.get_json())
        self.assertIn("尚未财务确认入账", blocked_submit.get_json()["message"])

        # 未确认：发起取消退款被拦
        blocked_cancel = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "未确认就取消"},
        )
        self.assertEqual(blocked_cancel.status_code, 400, blocked_cancel.get_json())
        self.assertIn("尚未财务确认入账", blocked_cancel.get_json()["message"])

        # 即使已收到首付款，只要存在未核对的意向定金，也不能进入退款流程
        vehicle2_id, vin2 = self.create_vehicle(car_type)
        paid_order_id = self.create_order(
            vin2, car_type,
            customer_phone="13500000000",
            intent_deposit_amount=1200,
            deposit_amount=2000,
            first_payment_received_amount=5000,
        )
        self.login("sales")
        blocked_paid_cancel = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": paid_order_id, "reason": "已补首付但定金未核对"},
        )
        self.assertEqual(blocked_paid_cancel.status_code, 400, blocked_paid_cancel.get_json())
        self.assertIn("尚未财务确认入账", blocked_paid_cancel.get_json()["message"])

        # 销售无权确认定金（403）
        denied = self.client.post(f"/api/sales-orders/{order_id}/confirm-intent-deposit", json={})
        self.assertEqual(denied.status_code, 403, denied.get_json())

        # 财务必须有销售支付凭证和公司到账流水号才能完成核对
        self.login("fin")
        missing_serial = self.client.post(
            f"/api/sales-orders/{order_id}/confirm-intent-deposit",
            json={},
        )
        self.assertEqual(missing_serial.status_code, 400, missing_serial.get_json())
        self.assertIn("流水号", missing_serial.get_json()["message"])

        self.db_exec(
            "UPDATE sales_orders SET customer_screenshot_path=NULL WHERE id=?",
            (order_id,),
        )
        missing_customer_proof = self.client.post(
            f"/api/sales-orders/{order_id}/confirm-intent-deposit",
            json={"serial": "BANK20260916XYZ"},
        )
        self.assertEqual(missing_customer_proof.status_code, 400, missing_customer_proof.get_json())
        self.assertIn("客户定金支付凭证", missing_customer_proof.get_json()["message"])
        self.db_exec(
            "UPDATE sales_orders SET customer_screenshot_path=? WHERE id=?",
            ("/uploads/customer-first-payment.jpg", order_id),
        )

        # 财务确认入账（带流水号 + 公司回单）
        self.confirm_intent_deposit(
            order_id,
            serial="BANK20260916XYZ",
            receipt_path="/uploads/company-deposit-receipt.jpg",
        )
        self.assertEqual(
            self.db_value("SELECT intent_deposit_serial FROM sales_orders WHERE id=?", (order_id,)),
            "BANK20260916XYZ",
        )
        self.assertEqual(
            self.db_value("SELECT intent_deposit_receipt_path FROM sales_orders WHERE id=?", (order_id,)),
            "/uploads/company-deposit-receipt.jpg",
        )
        self.assertEqual(
            self.db_value("SELECT intent_deposit_confirmed_by FROM sales_orders WHERE id=?", (order_id,)),
            self.db_value("SELECT display_name FROM users WHERE username='fin'"),
        )

        # 重复确认被拦
        self.login("fin")
        dup = self.client.post(f"/api/sales-orders/{order_id}/confirm-intent-deposit", json={})
        self.assertEqual(dup.status_code, 400, dup.get_json())

        # 同一请求内修改客户支付凭证并提交时，必须先作废旧确认，不能借用旧确认放行
        self.login("sales")
        blocked_changed_submit = self.client.put(
            f"/api/sales-orders/{order_id}",
            json={
                "customer_screenshot_path": "/uploads/customer-first-payment-v2.jpg",
                "action": "submit",
            },
        )
        self.assertEqual(blocked_changed_submit.status_code, 400, blocked_changed_submit.get_json())
        self.assertIn("尚未财务确认入账", blocked_changed_submit.get_json()["message"])

        # 确认后改动客户支付凭证 → 旧确认作废，提交再次被拦
        edited_proof = self.client.put(
            f"/api/sales-orders/{order_id}",
            json={"customer_screenshot_path": "/uploads/customer-first-payment-v2.jpg"},
        )
        self.assertEqual(edited_proof.status_code, 200, edited_proof.get_json())
        self.assertIsNone(
            self.db_value("SELECT intent_deposit_confirmed_at FROM sales_orders WHERE id=?", (order_id,))
        )
        blocked_after_proof = self.client.put(f"/api/sales-orders/{order_id}", json={"action": "submit"})
        self.assertEqual(blocked_after_proof.status_code, 400, blocked_after_proof.get_json())

        # 再次确认后改动意向定金金额 → 旧确认再次作废
        self.confirm_intent_deposit(order_id, serial="BANK20260916YYY")
        self.login("sales")
        edited = self.client.put(f"/api/sales-orders/{order_id}", json={"intent_deposit_amount": 2000})
        self.assertEqual(edited.status_code, 200, edited.get_json())
        self.assertIsNone(
            self.db_value("SELECT intent_deposit_confirmed_at FROM sales_orders WHERE id=?", (order_id,))
        )
        blocked_again = self.client.put(f"/api/sales-orders/{order_id}", json={"action": "submit"})
        self.assertEqual(blocked_again.status_code, 400, blocked_again.get_json())

        # 重新确认后提交放行
        self.confirm_intent_deposit(order_id, serial="BANK20260916ZZZ")
        self.submit_order(order_id)

        # 确认后取消退款放行（仅意向阶段，应退=意向定金新金额 2000）
        self.login("sales")
        created = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order_id, "reason": "确认后正常取消"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT refund_amount FROM order_refunds WHERE id=?",
                (created.get_json()["id"],),
            ),
            2000,
        )

    def test_api04d_hidden_legacy_contract_cannot_initiate_return(self):
        """历史合同未匹配销售时不可见且不能退车；补齐销售归属后恢复正常。"""
        vehicle_id, contract_id = self.create_active_contract(created_by="历史导入")
        self.db_exec(
            "UPDATE contracts SET contract_origin='legacy', legacy_visibility_status='已启用' WHERE id=?",
            (contract_id,),
        )

        self.login("sales")
        contracts = self.client.get("/api/contracts").get_json()
        self.assertNotIn(contract_id, {row["id"] for row in contracts})

        blocked = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "历史合同退车"},
        )
        self.assertEqual(blocked.status_code, 403, blocked.get_json())
        self.assertIn("销售归属", blocked.get_json()["message"])
        self.assertEqual(
            self.db_value("SELECT COUNT(*) FROM return_inspections WHERE vehicle_id=?", (vehicle_id,)),
            0,
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "租赁中",
        )

        sales_user_id = self.db_value("SELECT id FROM users WHERE username='sales'")
        self.db_exec(
            """
            INSERT INTO contract_role_assignments
                (contract_id, assignment_role, user_id, source_name, status, assigned_by)
            VALUES (?, 'signing_sales', ?, '周销售', '启用', '测试')
            """,
            (contract_id, sales_user_id),
        )

        visible_contracts = self.client.get("/api/contracts").get_json()
        self.assertIn(contract_id, {row["id"] for row in visible_contracts})
        created = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "历史合同退车"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        self.assertEqual(
            self.db_value("SELECT contract_id FROM return_inspections WHERE id=?", (created.get_json()["id"],)),
            contract_id,
        )
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "退车中",
        )

    # ---------------- API-05 续租全规则 ----------------
    def test_api05_renewal_guard_rules(self):
        _, contract_id = self.create_active_contract(end_date="2026-09-15")
        self.login("sales")
        expired = self.client.post(
            f"/api/contracts/{contract_id}/renewals",
            json={"billing_cycle": "按月", "monthly_rent": 3000, "loan_periods": 1},
        )
        self.assertEqual(expired.status_code, 400, expired.get_json())
        self.assertIn("到期日前", expired.get_json()["message"])

        self.db_exec("UPDATE contracts SET end_date='2027-01-31' WHERE id=?", (contract_id,))
        tampered = self.client.post(
            f"/api/contracts/{contract_id}/renewals",
            json={"billing_cycle": "按月", "monthly_rent": 3000, "loan_periods": 1,
                  "new_start_date": "2027-01-31"},
        )
        self.assertEqual(tampered.status_code, 400, tampered.get_json())
        self.assertIn("次日", tampered.get_json()["message"])

        created = self.client.post(
            f"/api/contracts/{contract_id}/renewals",
            json={"billing_cycle": "按月", "monthly_rent": 3000, "loan_periods": 1},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        renewal_id = created.get_json()["id"]
        self.assertEqual(
            self.db_value("SELECT new_start_date FROM renewal_applications WHERE id=?", (renewal_id,)),
            "2027-02-01",
        )
        self.approve_flow("renewal", renewal_id)

        self.login("ops")
        no_file = self.client.post(
            f"/api/renewals/{renewal_id}/activate",
            json={"contract_number": "XZ-2026-0001"},
        )
        self.assertEqual(no_file.status_code, 400, no_file.get_json())
        self.assertIn("续租合同", no_file.get_json()["message"])
        no_number = self.client.post(
            f"/api/renewals/{renewal_id}/activate",
            json={"contract_file": "/uploads/renewal.pdf"},
        )
        self.assertEqual(no_number.status_code, 400, no_number.get_json())
        self.assertIn("合同编号", no_number.get_json()["message"])

    # ---------------- API-06 再次续租 ----------------
    def test_api06_second_renewal_appends_third_segment(self):
        _, contract_id = self.create_active_contract(end_date="2026-12-31")
        first_id = self.create_and_activate_renewal(
            contract_id, billing_cycle="按月", monthly_rent=3000, loan_periods=1,
        )
        self.assertEqual(
            self.db_value("SELECT end_date FROM contracts WHERE id=?", (contract_id,)),
            "2027-01-31",
        )
        second_id = self.create_and_activate_renewal(
            contract_id, billing_cycle="按月", monthly_rent=3200, loan_periods=2,
        )
        self.assertEqual(
            self.db_value("SELECT end_date FROM contracts WHERE id=?", (contract_id,)),
            "2027-03-31",
            "再次续租应在同一合同主档上继续顺延",
        )
        monthly_bills = self.db_value(
            "SELECT COUNT(*) FROM repayments WHERE contract_id=? AND remark='续租按月账单'",
            (contract_id,),
        )
        self.assertEqual(monthly_bills, 3)
        statuses = self.db_rows(
            "SELECT status FROM renewal_applications WHERE id IN (?, ?)",
            (first_id, second_id),
        )
        self.assertTrue(all(row["status"] == "已生效" for row in statuses))

    # ---------------- API-07 短租按天续租（跨月） ----------------
    def test_api07_daily_renewal_cross_month_amounts(self):
        _, contract_id = self.create_active_contract(end_date="2027-01-29")
        renewal_id = self.create_and_activate_renewal(
            contract_id, billing_cycle="按天", monthly_rent=3000, rental_days=4,
        )
        self.assertEqual(
            self.db_value("SELECT end_date FROM contracts WHERE id=?", (contract_id,)),
            "2027-02-02",
        )
        bills = self.db_rows(
            "SELECT due_date, amount FROM repayments "
            "WHERE contract_id=? AND remark='续租按日账单' ORDER BY due_date",
            (contract_id,),
        )
        self.assertEqual(len(bills), 4)
        expected = {
            "2027-01-30": 96.77,
            "2027-01-31": 96.77,
            "2027-02-01": 107.14,
            "2027-02-02": 107.14,
        }
        for bill in bills:
            self.assertEqual(bill["amount"], expected[bill["due_date"]], bill)
        self.assertEqual(
            self.db_value("SELECT status FROM renewal_applications WHERE id=?", (renewal_id,)),
            "已生效",
        )

    # ---------------- API-08 退车/维修照片必填 ----------------
    def test_api08_return_and_repair_photos_required(self):
        vehicle_id, _ = self.create_active_contract()
        self.login("sales")
        created = self.client.post(
            "/api/return-inspections",
            json={"vehicle_id": vehicle_id, "return_reason": "到期退车"},
        )
        self.assertEqual(created.status_code, 200, created.get_json())
        return_id = created.get_json()["id"]

        base_fields = {
            "mileage": "30000", "body_tire_clean": "已清理", "accident_info": "无出险",
            "insurance_surcharge": "无", "violation_info": "无违章", "etc_info": "已注销",
            "maintenance_info": "正常",
        }
        self.login("fleet")
        missing_tool = self.client.post(
            f"/api/return-inspections/{return_id}/fleet",
            json={
                **base_fields,
                "appearance_photos": "/uploads/return-appearance.jpg",
                "mileage_photos": "/uploads/return-mileage.jpg",
            },
        )
        self.assertEqual(missing_tool.status_code, 400, missing_tool.get_json())
        self.assertIn("工具照片", missing_tool.get_json()["message"])

        saved = self.client.post(
            f"/api/return-inspections/{return_id}/fleet",
            json={
                **base_fields,
                "appearance_photos": "/uploads/return-appearance.jpg",
                "mileage_photos": "/uploads/return-mileage.jpg",
                "tools_photos": "/uploads/return-tools.jpg",
            },
        )
        self.assertEqual(saved.status_code, 200, saved.get_json())

        # 维修完成三照必填
        repair_vehicle_id, _ = self.create_vehicle("维修测试车型", "待维修")
        self.login("fleet")
        no_photo = self.client.post(
            f"/api/vehicles/{repair_vehicle_id}/repair/complete",
            json={"repair_completion_note": "已修复"},
        )
        self.assertEqual(no_photo.status_code, 400, no_photo.get_json())
        self.assertIn("维修后外观照片", no_photo.get_json()["message"])
        completed = self.client.post(
            f"/api/vehicles/{repair_vehicle_id}/repair/complete",
            json={
                "repair_completion_note": "已修复",
                "repair_appearance_photos": "/uploads/repair-appearance.jpg",
                "repair_mileage_photos": "/uploads/repair-mileage.jpg",
                "repair_tools_photos": "/uploads/repair-tools.jpg",
                "repair_cost": 800,
            },
        )
        self.assertEqual(completed.status_code, 200, completed.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (repair_vehicle_id,)),
            "待整备",
        )

    # ---------------- API-09 整备状态机 ----------------
    def test_api09_refurbishment_state_machine(self):
        vehicle_id, _ = self.create_vehicle("整备测试车型", "待维修")
        self.login("fleet")
        completed = self.client.post(
            f"/api/vehicles/{vehicle_id}/repair/complete",
            json={
                "repair_completion_note": "已修复",
                "repair_appearance_photos": "/uploads/repair-appearance.jpg",
                "repair_mileage_photos": "/uploads/repair-mileage.jpg",
                "repair_tools_photos": "/uploads/repair-tools.jpg",
            },
        )
        self.assertEqual(completed.status_code, 200, completed.get_json())
        refurb_id = self.db_value(
            "SELECT id FROM refurbishment_records WHERE vehicle_id=? ORDER BY id DESC LIMIT 1",
            (vehicle_id,),
        )
        self.assertIsNotNone(refurb_id)
        self.assertEqual(
            self.db_value("SELECT status FROM refurbishment_records WHERE id=?", (refurb_id,)),
            "待整备",
        )

        premature = self.client.post(
            f"/api/refurbishment-records/{refurb_id}/complete",
            json={"completion_note": "未开始直接完成"},
        )
        self.assertEqual(premature.status_code, 400, premature.get_json())

        started = self.client.post(f"/api/refurbishment-records/{refurb_id}/start", json={})
        self.assertEqual(started.status_code, 200, started.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle_id,)),
            "整备中",
        )
        restart = self.client.post(f"/api/refurbishment-records/{refurb_id}/start", json={})
        self.assertEqual(restart.status_code, 400, restart.get_json())

        no_note = self.client.post(f"/api/refurbishment-records/{refurb_id}/complete", json={})
        self.assertEqual(no_note.status_code, 400, no_note.get_json())
        done = self.client.post(
            f"/api/refurbishment-records/{refurb_id}/complete",
            json={"completion_note": "整备完毕", "available_for": "可租"},
        )
        self.assertEqual(done.status_code, 200, done.get_json())
        vehicle = self.db_row("SELECT status, condition FROM vehicles WHERE id=?", (vehicle_id,))
        self.assertEqual(vehicle["status"], "在库")
        self.assertEqual(vehicle["condition"], "二手车", "出库回库整备完成后应为二手车")
        again = self.client.post(
            f"/api/refurbishment-records/{refurb_id}/complete",
            json={"completion_note": "重复完成"},
        )
        self.assertEqual(again.status_code, 400, again.get_json())

    # ---------------- API-10 提前退车折算边界 ----------------
    def test_api10_prorated_rent_month_boundaries(self):
        feb = app_module.calculate_prorated_rent(3000, "2026-02-10", "2026-02-01")
        self.assertEqual(feb["days_in_month"], 28)
        self.assertEqual(feb["days_used"], 10)
        self.assertEqual(feb["amount"], round(3000 / 28 * 10, 2))

        jan_full = app_module.calculate_prorated_rent(3000, "2026-01-31", "2026-01-01")
        self.assertEqual(jan_full["days_in_month"], 31)
        self.assertEqual(jan_full["days_used"], 31)
        self.assertEqual(jan_full["amount"], 3000)

        mid = app_module.calculate_prorated_rent(3000, "2026-03-15", "2026-02-20")
        self.assertEqual(mid["days_in_month"], 31)
        self.assertEqual(mid["days_used"], 15)

    # ---------------- 对账单合同级详情 ----------------
    def test_reconciliation_contract_detail_fields_and_scope(self):
        vehicle_id, contract_id = self.create_active_contract(
            end_date="2026-10-31", created_by="周销售"
        )
        self.db_exec(
            """
            UPDATE contracts
            SET company='陕西解放公司', rental_method='经营租赁',
                start_date='2026-08-01', end_date='2026-10-31'
            WHERE id=?
            """,
            (contract_id,),
        )
        self.db_exec(
            "UPDATE vehicles SET company='车辆所属公司', plate_number='陕A12345' WHERE id=?",
            (vehicle_id,),
        )
        conn = database.get_db()
        try:
            conn.executemany(
                """
                INSERT INTO repayments
                    (contract_id, period, due_date, amount, paid_amount, status,
                     paid_at, screenshot_path, bank_serial, reported_amount,
                     verified_by, verified_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (
                        contract_id, 1, "2026-08-01", 3000, 1000, "部分核销",
                        "2026-08-03", "/uploads/customer-payment-1.jpg",
                        "BANK-DETAIL-001", 1000, "财务甲", "2026-08-03 10:00:00",
                    ),
                    (
                        contract_id, 2, "2026-09-01", 3000, 3000, "已还款",
                        "2026-09-01", "/uploads/customer-payment-2.jpg",
                        "BANK-DETAIL-002", 3000, "财务甲", "2026-09-01 10:00:00",
                    ),
                    (
                        contract_id, 3, "2026-10-01", 3000, 0, "待还款",
                        "", "", "", 0, "", "",
                    ),
                ],
            )
            repayment_id = conn.execute(
                "SELECT id FROM repayments WHERE contract_id=? AND period=1",
                (contract_id,),
            ).fetchone()[0]
            conn.executemany(
                """
                INSERT INTO reconciliation_allocations
                    (repayment_id, contract_id, allocation_type, allocated_amount, created_by)
                VALUES (?, ?, ?, ?, ?)
                """,
                [
                    (repayment_id, contract_id, "rent_waiver", 100, "财务甲"),
                    (repayment_id, contract_id, "bad_debt", 200, "财务甲"),
                    (repayment_id, contract_id, "late_fee_waiver", 30, "财务甲"),
                ],
            )
            conn.executemany(
                """
                INSERT INTO late_fee_ledger
                    (repayment_id, contract_id, period, accrued_date,
                     outstanding, daily_amount, cumulative_amount, waived, waived_amount)
                VALUES (?, ?, 1, ?, 2000, 1.5, ?, 0, 0)
                """,
                [
                    (repayment_id, contract_id, "2026-08-02", 1.5),
                    (repayment_id, contract_id, "2026-08-03", 3.0),
                ],
            )
            conn.execute(
                """
                INSERT INTO invoice_requests
                    (contract_id, period, amount, invoice_no, invoice_date,
                     status, applied_by, processed_by)
                VALUES (?, 2, 3000, 'INV-DETAIL-002', '2026-09-05',
                        '已开票', '运营甲', '财务甲')
                """,
                (contract_id,),
            )
            conn.commit()
        finally:
            conn.close()

        self.login("sales")
        response = self.client.get(f"/api/reconciliation/{repayment_id}/detail")
        self.assertEqual(response.status_code, 200, response.get_json())
        payload = response.get_json()
        self.assertTrue(payload["success"])
        self.assertEqual(payload["summary"]["company"], "车辆所属公司")
        self.assertEqual(payload["summary"]["customer_name"], "续租测试客户")
        self.assertEqual(payload["summary"]["plate_number"], "陕A12345")
        self.assertEqual(payload["summary"]["rental_method"], "经营租赁")
        self.assertEqual(payload["summary"]["sales_advisor"], "周销售")
        self.assertEqual([item["period"] for item in payload["periods"]], [1, 2, 3])

        first, second, third = payload["periods"]
        self.assertEqual(first["period_start"], "2026-08-01")
        self.assertEqual(first["period_end"], "2026-08-31")
        self.assertEqual(first["paid_at"], "2026-08-03")
        self.assertEqual(first["waiver_amount"], 100)
        self.assertEqual(first["bad_debt_amount"], 200)
        self.assertEqual(first["overdue_amount"], 2000)
        self.assertGreater(first["overdue_days"], 0)
        self.assertEqual(first["breach_days"], 2)
        self.assertEqual(first["breach_rate"], "0.05%")
        self.assertEqual(first["late_fee_amount"], 3)
        self.assertEqual(first["late_fee_waived_amount"], 30)
        self.assertEqual(first["screenshot_path"], "/uploads/customer-payment-1.jpg")
        self.assertEqual(second["invoice_date"], "2026-09-05")
        self.assertEqual(second["invoice_no"], "INV-DETAIL-002")
        self.assertEqual(second["paid_amount"], 3000)
        self.assertEqual(second["overdue_amount"], 0)
        self.assertEqual(third["period_start"], "2026-10-01")
        self.assertEqual(third["period_end"], "2026-10-31")

        self.create_second_sales()
        self.login("sales2")
        denied = self.client.get(f"/api/reconciliation/{repayment_id}/detail")
        self.assertEqual(denied.status_code, 403, denied.get_json())

    def test_reconciliation_detail_bad_debt_split_from_finance_execution(self):
        vehicle_id, contract_id = self.create_active_contract(
            end_date="2026-12-31", created_by="周销售"
        )
        self.db_exec(
            "UPDATE vehicles SET company='车辆所属公司' WHERE id=?",
            (vehicle_id,),
        )
        conn = database.get_db()
        try:
            conn.execute(
                """
                INSERT INTO repayments
                    (contract_id, period, due_date, amount, paid_amount, status)
                VALUES (?, 1, '2026-10-01', 3000, 0, '待还款')
                """,
                (contract_id,),
            )
            repayment_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.execute(
                """
                INSERT INTO waivers
                    (contract_id, waiver_kind, target_period_list, waive_amount,
                     reason, attachment_path, status, sales_applied_by)
                VALUES (?, 'rent', 'all', 300, '历史欠款核销',
                        '/uploads/customer-payment.jpg', '待审批', '周销售')
                """,
                (contract_id,),
            )
            waiver_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            conn.commit()
        finally:
            conn.close()

        self.login("boss")
        approved = self.client.post(
            f"/api/waivers/{waiver_id}/approve",
            json={"decision": "approve", "comment": "同意"},
        )
        self.assertEqual(approved.status_code, 200, approved.get_json())

        self.login("ops")
        operations = self.client.post(
            f"/api/waivers/{waiver_id}/operations",
            json={
                "approved_waive_amount": 100,
                "bad_debt_amount": 200,
                "carryover_start_period": 1,
            },
        )
        self.assertEqual(operations.status_code, 200, operations.get_json())

        self.login("fin")
        executed = self.client.post(f"/api/waivers/{waiver_id}/execute", json={})
        self.assertEqual(executed.status_code, 200, executed.get_json())
        self.assertEqual(
            self.db_value(
                """
                SELECT ROUND(SUM(allocated_amount), 2)
                FROM reconciliation_allocations
                WHERE repayment_id=? AND allocation_type='rent_waiver'
                """,
                (repayment_id,),
            ),
            100,
        )
        self.assertEqual(
            self.db_value(
                """
                SELECT ROUND(SUM(allocated_amount), 2)
                FROM reconciliation_allocations
                WHERE repayment_id=? AND allocation_type='bad_debt'
                """,
                (repayment_id,),
            ),
            200,
        )

        detail = self.client.get(f"/api/reconciliation/{repayment_id}/detail")
        self.assertEqual(detail.status_code, 200, detail.get_json())
        period = detail.get_json()["periods"][0]
        self.assertEqual(period["waiver_amount"], 100)
        self.assertEqual(period["bad_debt_amount"], 200)
        self.assertEqual(period["amount"], 2700)

    # ---------------- API-11 发票作废/红冲 ----------------
    def test_api11_invoice_void_red_and_reject(self):
        _, contract_id = self.create_active_contract()

        def create_invoice():
            self.login("ops")
            response = self.client.post(
                f"/api/contracts/{contract_id}/invoices",
                json={"amount": 3000, "invoice_entity_name": "测试计划客户公司"},
            )
            self.assertEqual(response.status_code, 200, response.get_json())
            return response.get_json()["id"]

        def approve_and_issue(invoice_id):
            self.login("boss")
            approved = self.client.post(f"/api/invoice-requests/{invoice_id}/approve", json={})
            self.assertEqual(approved.status_code, 200, approved.get_json())
            self.login("fin")
            issued = self.client.post(
                f"/api/invoice-requests/{invoice_id}/issue",
                json={"invoice_no": f"INV{invoice_id:08d}", "invoice_date": "2026-09-16"},
            )
            self.assertEqual(issued.status_code, 200, issued.get_json())

        inv1 = create_invoice()
        approve_and_issue(inv1)
        self.login("fin")
        no_reason = self.client.post(f"/api/invoice-requests/{inv1}/void", json={})
        self.assertEqual(no_reason.status_code, 400, no_reason.get_json())
        voided = self.client.post(
            f"/api/invoice-requests/{inv1}/void",
            json={"reason": "开票信息错误"},
        )
        self.assertEqual(voided.status_code, 200, voided.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM invoice_requests WHERE id=?", (inv1,)),
            "已作废",
        )
        again = self.client.post(
            f"/api/invoice-requests/{inv1}/void",
            json={"reason": "重复作废"},
        )
        self.assertEqual(again.status_code, 400, again.get_json())

        inv2 = create_invoice()
        approve_and_issue(inv2)
        self.login("fin")
        no_no = self.client.post(
            f"/api/invoice-requests/{inv2}/red",
            json={"reason": "客户退票"},
        )
        self.assertEqual(no_no.status_code, 400, no_no.get_json())
        red = self.client.post(
            f"/api/invoice-requests/{inv2}/red",
            json={"red_invoice_no": "RED20260916001", "reason": "客户退票"},
        )
        self.assertEqual(red.status_code, 200, red.get_json())
        row = self.db_row("SELECT status, red_invoice_no FROM invoice_requests WHERE id=?", (inv2,))
        self.assertEqual(row["status"], "已红冲")
        self.assertEqual(row["red_invoice_no"], "RED20260916001")

        inv3 = create_invoice()
        self.login("boss")
        reject_no_reason = self.client.post(f"/api/invoice-requests/{inv3}/reject", json={})
        self.assertEqual(reject_no_reason.status_code, 400, reject_no_reason.get_json())
        rejected = self.client.post(
            f"/api/invoice-requests/{inv3}/reject",
            json={"reason": "金额不符"},
        )
        self.assertEqual(rejected.status_code, 200, rejected.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM invoice_requests WHERE id=?", (inv3,)),
            "已驳回",
        )
        inv4 = create_invoice()
        self.login("fin")
        pending_void = self.client.post(
            f"/api/invoice-requests/{inv4}/void",
            json={"reason": "待审批尝试作废"},
        )
        self.assertEqual(pending_void.status_code, 400, pending_void.get_json())

    def test_invoice_records_list_role_scope_and_status_filter(self):
        self.create_second_sales()
        _, own_contract_id = self.create_active_contract(created_by="周销售")
        _, other_contract_id = self.create_active_contract(created_by="李销售")
        self.db_exec(
            """
            INSERT INTO invoice_requests
                (contract_id, period, amount, receiving_company, invoice_entity_name,
                 invoice_no, invoice_date, status, applied_by, processed_by)
            VALUES
                (?, 1, 3000, '金聚源', '自有客户公司', 'INV-OWN-001',
                 '2026-09-16', '已开票', '李运营', '张财务'),
                (?, 1, 5000, '金聚源', '其他客户公司', 'INV-OTHER-001',
                 '2026-09-17', '已开票', '李运营', '张财务')
            """,
            (own_contract_id, other_contract_id),
        )
        self.db_exec(
            """
            INSERT INTO invoice_requests
                (contract_id, period, amount, invoice_entity_name, status, applied_by)
            VALUES (?, 1, 9000, '待开票公司', '待开票', '李运营')
            """,
            (own_contract_id,),
        )

        anonymous = self.client.get("/api/invoice-records")
        self.assertEqual(anonymous.status_code, 401, anonymous.get_json())

        self.login("fin")
        finance = self.client.get("/api/invoice-records")
        self.assertEqual(finance.status_code, 200, finance.get_json())
        self.assertEqual(len(finance.get_json()), 2)

        self.login("ops")
        operations = self.client.get("/api/invoice-records")
        self.assertEqual(operations.status_code, 200, operations.get_json())
        self.assertEqual(
            {row["invoice_no"] for row in operations.get_json()},
            {"INV-OWN-001", "INV-OTHER-001"},
        )

        self.login("boss")
        boss = self.client.get("/api/invoice-records")
        self.assertEqual(boss.status_code, 200, boss.get_json())
        self.assertEqual(len(boss.get_json()), 2)

        self.login("sales")
        own_records = self.client.get("/api/invoice-records")
        self.assertEqual(own_records.status_code, 200, own_records.get_json())
        self.assertEqual([row["invoice_no"] for row in own_records.get_json()], ["INV-OWN-001"])
        self.assertEqual(own_records.get_json()[0]["processed_by"], "张财务")

        self.login("sales2")
        other_records = self.client.get("/api/invoice-records")
        self.assertEqual(other_records.status_code, 200, other_records.get_json())
        self.assertEqual([row["invoice_no"] for row in other_records.get_json()], ["INV-OTHER-001"])

    # ---------------- API-12 Excel 导入入库 ----------------
    def _build_vehicle_import_workbook(self, vin):
        wb = Workbook()
        ws = wb.active
        ws.append(["序号", "成色", "品牌", "品系", "车型", "VIN", "网员价", "厢型"])
        ws.append([1, "新车", "解放", "虎V", "API12导入车型", vin, 90000, "厢货"])
        ws.append([2, "新车", "解放", "虎V", "API12导入车型", "SHORT", 90000, "厢货"])
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    def test_api12_excel_import_inventory(self):
        vin = "IMP2026091600001A"
        self.assertEqual(len(vin), 17)

        self.login("sales")
        denied = self.client.post(
            "/api/vehicles/import",
            data={"file": (self._build_vehicle_import_workbook(vin), "vehicles.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(denied.status_code, 200)
        self.assertFalse(denied.get_json()["success"])

        self.login("ops")
        first = self.client.post(
            "/api/vehicles/import",
            data={"file": (self._build_vehicle_import_workbook(vin), "vehicles.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(first.status_code, 200)
        result = first.get_json()
        self.assertTrue(result["success"], result)
        self.assertEqual(result["imported"], 1)
        self.assertEqual(result["failed"], 1)
        vehicle = self.db_row("SELECT car_type, status FROM vehicles WHERE vin=?", (vin,))
        self.assertIsNotNone(vehicle)
        self.assertEqual(vehicle["status"], "在库")

        second = self.client.post(
            "/api/vehicles/import",
            data={"file": (self._build_vehicle_import_workbook(vin), "vehicles.xlsx")},
            content_type="multipart/form-data",
        )
        result2 = second.get_json()
        self.assertTrue(result2["success"], result2)
        self.assertEqual(result2["imported"], 0)
        self.assertEqual(result2["skipped"], 1)
        self.assertEqual(
            self.db_value("SELECT COUNT(*) FROM vehicles WHERE vin=?", (vin,)),
            1,
        )

    # ---------------- API-13 指导价导入 ----------------
    def test_api13_guidance_price_import(self):
        def build_workbook():
            wb = Workbook()
            ws = wb.active
            ws.title = "新能源"
            ws.append(["车型", "整车", "以租代购方案"])
            ws.append(["API13导入车型", 10, "首付3万，3600三年"])
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return buf

        self.login("boss")
        response = self.client.post(
            "/api/model-guidance-prices/import",
            data={"file": (build_workbook(), "guidance.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200, response.get_json())
        result = response.get_json()
        self.assertTrue(result["success"], result)
        self.assertEqual(result["created"], 1)
        row = self.db_row(
            "SELECT guidance_price, lease_installment_price, fuel_type "
            "FROM model_guidance_prices WHERE car_type=?",
            ("API13导入车型",),
        )
        self.assertIsNotNone(row)
        self.assertEqual(row["guidance_price"], 100000)
        self.assertEqual(row["lease_installment_price"], 3600)
        self.assertEqual(row["fuel_type"], "新能源")

        again = self.client.post(
            "/api/model-guidance-prices/import",
            data={"file": (build_workbook(), "guidance.xlsx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(again.get_json()["updated"], 1)
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM model_guidance_prices WHERE car_type=?",
                ("API13导入车型",),
            ),
            1,
        )

    # ---------------- API-14 销售数据范围 ----------------
    def test_api14_sales_data_scope(self):
        self.create_second_sales()
        _, own_contract = self.create_active_contract(created_by="周销售")
        _, other_contract = self.create_active_contract(created_by="李销售")
        self.db_exec(
            "UPDATE customers SET name='周销售客户' "
            "WHERE id=(SELECT customer_id FROM contracts WHERE id=?)",
            (own_contract,),
        )
        self.db_exec(
            "UPDATE customers SET name='李销售客户' "
            "WHERE id=(SELECT customer_id FROM contracts WHERE id=?)",
            (other_contract,),
        )

        self.login("sales")
        library = self.client.get("/api/customer-library")
        self.assertEqual(library.status_code, 200, library.get_json())
        names = [row["customer_name"] for row in library.get_json()]
        self.assertIn("周销售客户", names)
        self.assertNotIn("李销售客户", names)

        contracts = self.client.get("/api/contracts")
        contract_ids = [row["id"] for row in contracts.get_json()]
        self.assertIn(own_contract, contract_ids)
        self.assertNotIn(other_contract, contract_ids)

        self.login("sales2")
        library2 = self.client.get("/api/customer-library")
        names2 = [row["customer_name"] for row in library2.get_json()]
        self.assertIn("李销售客户", names2)
        self.assertNotIn("周销售客户", names2)

        for role_user in ("boss", "fin"):
            self.login(role_user)
            all_rows = self.client.get("/api/customer-library")
            all_names = [row["customer_name"] for row in all_rows.get_json()]
            self.assertIn("周销售客户", all_names)
            self.assertIn("李销售客户", all_names)

    # ---------------- API-15 敏感字段脱敏 ----------------
    def test_api15_hidden_fields_by_role(self):
        self.create_vehicle("脱敏测试车型", purchase_price=98765)
        self.login("sales")
        sales_vehicles = self.client.get("/api/vehicles?page_size=50")
        self.assertEqual(sales_vehicles.status_code, 200)
        sales_row = next(
            r for r in sales_vehicles.get_json()["data"] if r["car_type"] == "脱敏测试车型"
        )
        self.assertNotIn("purchase_price", sales_row)

        self.login("fleet")
        fleet_vehicles = self.client.get("/api/vehicles?page_size=50")
        fleet_row = next(
            r for r in fleet_vehicles.get_json()["data"] if r["car_type"] == "脱敏测试车型"
        )
        self.assertNotIn("purchase_price", fleet_row)

        self.login("boss")
        boss_vehicles = self.client.get("/api/vehicles?page_size=50")
        boss_row = next(
            r for r in boss_vehicles.get_json()["data"] if r["car_type"] == "脱敏测试车型"
        )
        self.assertEqual(boss_row["purchase_price"], 98765)

        _, contract_id = self.create_active_contract(deposit=2000, collected=2000)
        self.login("fleet")
        fleet_contracts = self.client.get("/api/contracts")
        fleet_contract = next(r for r in fleet_contracts.get_json() if r["id"] == contract_id)
        self.assertNotIn("deposit", fleet_contract)
        self.assertNotIn("loan_amount", fleet_contract)
        self.login("boss")
        boss_contracts = self.client.get("/api/contracts")
        boss_contract = next(r for r in boss_contracts.get_json() if r["id"] == contract_id)
        self.assertEqual(boss_contract["deposit"], 2000)

    # ---------------- API-16 幂等与重复提交 ----------------
    def test_api16_idempotent_operations(self):
        car_type = "API16幂等车型"
        self.save_lease_guidance(car_type)
        vehicle_id, vin = self.create_vehicle(car_type)
        order_id = self.create_order(vin, car_type)
        self.submit_order(order_id)
        contract_id = self.deliver_order_pipeline(order_id, vehicle_id)

        # 重复出库 → 400 且提示无需重复操作
        self.login("fleet")
        redeliver = self.client.post(f"/api/vehicles/{vehicle_id}/deliver", json={})
        self.assertEqual(redeliver.status_code, 400, redeliver.get_json())
        self.assertIn("无需重复操作", redeliver.get_json()["message"])

        # 重复核销 → 400
        repayment_id = self.db_value(
            "SELECT id FROM repayments WHERE contract_id=? AND period=2", (contract_id,),
        )
        self.assertIsNotNone(repayment_id)
        self.login("ops")
        screenshot = self.client.post(
            f"/api/reconciliation/{repayment_id}/screenshot",
            json={"screenshot_path": "/uploads/rec2-customer.jpg", "reported_amount": 3000},
        )
        self.assertEqual(screenshot.status_code, 200, screenshot.get_json())
        self.login("fin")
        verified = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "BANKAPI16-001", "received_amount": 3000},
        )
        self.assertEqual(verified.status_code, 200, verified.get_json())
        reverify = self.client.post(
            f"/api/reconciliation/{repayment_id}/verify",
            json={"bank_serial": "BANKAPI16-002", "received_amount": 3000},
        )
        self.assertEqual(reverify.status_code, 400, reverify.get_json())
        self.assertEqual(
            self.db_value(
                "SELECT COUNT(*) FROM repayments WHERE id=? AND status='已还款'",
                (repayment_id,),
            ),
            1,
        )

        # 重复激活续租 → 400
        self.db_exec("UPDATE contracts SET end_date='2026-12-31' WHERE id=?", (contract_id,))
        renewal_id = self.create_and_activate_renewal(
            contract_id, billing_cycle="按月", monthly_rent=3000, loan_periods=1,
        )
        self.login("ops")
        reactivate = self.client.post(
            f"/api/renewals/{renewal_id}/activate",
            json={"contract_file": "/uploads/renewal-contract.pdf", "contract_number": "XZ-DUP"},
        )
        self.assertEqual(reactivate.status_code, 400, reactivate.get_json())

        # 重复退款 → 400
        vehicle2_id, vin2 = self.create_vehicle(car_type)
        order2_id = self.create_order(vin2, car_type, customer_phone="13200000000")
        self.submit_order(order2_id)
        self.login("sales")
        created = self.client.post(
            "/api/order-refunds",
            json={"sales_order_id": order2_id, "reason": "客户取消"},
        )
        refund_id = created.get_json()["id"]
        self.login("boss")
        self.client.post(f"/api/order-refunds/{refund_id}/approve", json={})
        self.login("fin")
        paid = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND-API16-01", "refund_paid_amount": 2000},
        )
        self.assertEqual(paid.status_code, 200, paid.get_json())
        repay = self.client.post(
            f"/api/order-refunds/{refund_id}/pay",
            json={"refund_serial": "REFUND-API16-02", "refund_paid_amount": 2000},
        )
        self.assertEqual(repay.status_code, 400, repay.get_json())
        self.assertEqual(
            self.db_value("SELECT status FROM vehicles WHERE id=?", (vehicle2_id,)),
            "在库",
        )


if __name__ == "__main__":
    unittest.main()
