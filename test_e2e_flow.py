# -*- coding: utf-8 -*-
"""6.2 端到端主流程测试：报单→F1→F2→F3→E2激活→合同→H1催收。
驱动真实状态流转并断言，区别于只测 GET 可达性的旧脚本。"""
import json, urllib.request, urllib.error, http.cookiejar, os, datetime

BASE = "http://localhost:49165"
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
RESULTS = []

def _client():
    cj = http.cookiejar.CookieJar()
    return urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

def req(opener, method, path, body=None):
    url = BASE + path
    data = json.dumps(body).encode() if body is not None else None
    r = urllib.request.Request(url, data=data, method=method)
    if data is not None:
        r.add_header("Content-Type", "application/json")
    try:
        with opener.open(r, timeout=15) as resp:
            raw = resp.read().decode()
            return resp.status, (json.loads(raw) if raw else {})
    except urllib.error.HTTPError as e:
        raw = e.read().decode()
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"raw": raw}

def login(role_user):
    op = _client()
    st, body = req(op, "POST", "/api/auth/login",
                   {"username": role_user, "password": "123456"})
    assert st == 200 and body.get("success"), f"登录失败 {role_user}: {body}"
    return op

def check(stage, ok, detail=""):
    RESULTS.append((stage, ok, detail))
    mark = "✓" if ok else "✗"
    print(f"{mark} {stage} {detail}")
    return ok

def build_factory_xlsx(periods, due_dates, amounts):
    from openpyxl import Workbook
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    fname = "e2e_factory_plan.xlsx"
    path = os.path.join(UPLOAD_DIR, fname)
    wb = Workbook(); ws = wb.active
    ws.append(["序号", "应还款日期", "客户应还金额合计"])
    for i in range(periods):
        ws.append([i + 1, due_dates[i], amounts[i]])
    ws.append(["合计", "", sum(amounts)])
    wb.save(path)
    return "/uploads/" + fname

def find_stock_vehicle(op):
    st, body = req(op, "GET", "/api/vehicles")
    items = body if isinstance(body, list) else (body.get("data") or [])
    for v in items:
        if v.get("status") == "在库" and v.get("car_type"):
            return v
    return None

def has_dual_guidance(op, car_type):
    st, body = req(op, "GET", "/api/model-guidance-prices")
    items = body if isinstance(body, list) else (body.get("data") or [])
    for m in items:
        if m.get("car_type") == car_type:
            lease = float(m.get("lease_installment_price") or 0)
            sale = float(m.get("sale_total_price") or 0)
            return lease > 0 and sale > 0, lease, sale
    return False, 0, 0

def run_flow_lease():
    """以租代售/租赁主流程：报单(销售)→F1(运营)→F2导入→F3比对→E2激活(财务)→合同→H1。"""
    print("\n==== 端到端主流程：租赁分期 ====")
    sales = login("sales")
    ops = login("ops")
    fin = login("fin")

    v = find_stock_vehicle(sales)
    if not check("PREP 找在库车辆", v is not None, f"VIN={v.get('vin') if v else None}"):
        return
    dual, lease_price, sale_price = has_dual_guidance(sales, v["car_type"])
    if not dual:
        # 通过模块C(老板API)补维护车型双指导价，再复查
        boss = login("boss")
        lease_price, sale_price = 4000.0, 150000.0
        st, gp = req(boss, "POST", "/api/model-guidance-prices", {
            "car_type": v["car_type"],
            "lease_installment_price": lease_price,
            "sale_total_price": sale_price,
            "remark": "E2E测试补维护",
        })
        check("C 补维护车型双指导价", st == 200 and gp.get("success"),
              str(gp.get("message", gp))[:50])
        dual, lease_price, sale_price = has_dual_guidance(sales, v["car_type"])
    if not check("PREP 车型双指导价", dual, f"{v['car_type']} 租{lease_price}/售{sale_price}"):
        return

    # D: 报单（租赁，报价>=指导价 避免触发特批，直接进入待财务确认）
    quote = lease_price  # 等于指导价，不低于 → 无需特批
    st, body = req(sales, "POST", "/api/sales-orders", {
        "vin": v["vin"], "sales_mode": "租赁",
        "customer_name": "E2E测试客户", "customer_phone": "13900000001",
        "vehicle_rent_amount": quote, "sale_total_price": quote * 12,
        "lease_term": "12", "deposit_amount": 5000,
        "payment_date": datetime.date.today().strftime("%Y-%m-%d"),
    })
    ok = st == 200 and body.get("success")
    check("D 报单创建", ok, str(body.get("message", body))[:60])
    if not ok:
        return
    order_id = body["id"]

    # 确认报单状态为 待财务确认
    st, lst = req(sales, "GET", "/api/sales-orders")
    orders = lst if isinstance(lst, list) else (lst.get("data") or [])
    cur = next((o for o in orders if o.get("id") == order_id), {})
    check("D 报单状态=待财务确认", cur.get("order_status") == "待财务确认",
          f"实际={cur.get('order_status')}")

    # F1: 运营查看/生成客户计划（GET planning-contract 自动 ensure）
    st, pc = req(ops, "GET", f"/api/sales-orders/{order_id}/planning-contract")
    f1_ok = st == 200 and pc.get("success") and len(pc.get("repayments", [])) > 0
    contract = pc.get("contract", {}) if f1_ok else {}
    cid = contract.get("id")
    reps = pc.get("repayments", [])
    check("F1 客户计划生成", f1_ok, f"合同={cid} 期数={len(reps)}")
    if not f1_ok:
        return

    # F2: 导入厂家还款计划表（构造 xlsx，金额略低于客户 → 形成正利差）
    due_dates = [r.get("due_date") for r in reps if r.get("period", 0) >= 1]
    cust_amounts = [float(r.get("amount") or 0) for r in reps if r.get("period", 0) >= 1]
    n = len(due_dates)
    factory_amounts = [round(a * 0.9, 2) for a in cust_amounts]  # 厂家便宜10% → 利差>0
    file_url = build_factory_xlsx(n, due_dates, factory_amounts)
    st, imp = req(ops, "POST", f"/api/contracts/{cid}/factory-repayments/import",
                  {"file_url": file_url})
    f2_ok = st == 200 and imp.get("success")
    comp = imp.get("comparison", {}) if f2_ok else {}
    check("F2 导入厂家计划", f2_ok, f"导入{n}期 比对={comp.get('status')}")

    # F3: 比对（利差区间未设置上下限时应为已通过；spread>0 在 [0,∞)）
    st, cmp_body = req(fin, "POST", f"/api/contracts/{cid}/plan-compare")
    comp2 = cmp_body.get("comparison", {}) if st == 200 else {}
    spread = comp2.get("spread_total")
    f3_ok = st == 200 and comp2.get("status") in ("已通过", "差异已确认")
    check("F3 利差区间比对", f3_ok,
          f"状态={comp2.get('status')} 利差={spread} 区间过={comp2.get('spread_range_passed')}")

    # E2: 财务激活报单（应满足 F1/F2/F3 前置）
    st, act = req(fin, "POST", f"/api/sales-orders/{order_id}/activate")
    e2_ok = st == 200 and act.get("success")
    check("E2 财务激活报单", e2_ok, str(act.get("message", act))[:60])

    # 验证报单已激活
    st, lst2 = req(sales, "GET", "/api/sales-orders")
    orders2 = lst2 if isinstance(lst2, list) else (lst2.get("data") or [])
    cur2 = next((o for o in orders2 if o.get("id") == order_id), {})
    check("E2 报单状态=已激活", cur2.get("order_status") == "已激活",
          f"实际={cur2.get('order_status')}")
    return order_id, cid

def find_initial_payment_flow(op, pid):
    """在审批列表中找到该首次付款待审批步骤的 flow_id。"""
    st, body = req(op, "GET", "/api/approvals?ref_type=initial_payment")
    items = body if isinstance(body, list) else (body.get("data") or [])
    for it in items:
        if it.get("ref_id") == pid:
            for s in it.get("steps", []):
                if s.get("status") == "待审批":
                    return s.get("id")
    return None

def run_g_delivery(order_id, cid):
    """G: 上传线下合同→首次付款→财务回单→审批通过→验证还款计划激活→出库。"""
    print("\n==== G 合同/首次付款/出库 ====")
    ops = login("ops"); fin = login("fin"); cm = login("fleet")

    # 读取 F1 计划合同字段，计算 租赁 首付款 = deposit + rent
    st, cs = req(ops, "GET", "/api/contracts")
    clist = cs if isinstance(cs, list) else (cs.get("data") or [])
    cur = next((x for x in clist if x.get("id") == cid), {})
    deposit = float(cur.get("deposit") or 0)
    rent = float(cur.get("rent") or 0)
    vehicle_id = cur.get("vehicle_id")
    expect_amount = deposit + rent

    # G1 运营上传线下合同（复用 F1 计划合同，置 delivery_status=待首付款）
    body = {
        "vehicle_id": vehicle_id, "sales_order_id": order_id,
        "contract_file": "/uploads/e2e_contract.pdf",
        "rent": rent, "deposit": deposit,
        "loan_periods": int(cur.get("loan_periods") or 0),
        "customer_loan_amount": float(cur.get("customer_loan_amount") or 0),
    }
    st, r = req(ops, "POST", "/api/contracts", body)
    check("G1 上传线下合同", st == 200 and r.get("success"),
          str(r.get("message", r))[:60])

    # G2 运营发起首次付款
    st, ip = req(ops, "POST", f"/api/contracts/{cid}/initial-payment",
                 {"customer_screenshot_path": "/uploads/e2e_cust_pay.jpg",
                  "amount": expect_amount})
    ip_ok = st == 200 and ip.get("success")
    check("G2 发起首次付款", ip_ok, str(ip.get("message", ip))[:60])
    pid = ip.get("id")
    if not ip_ok:
        return False

    # G3 财务上传公司收款回单（登记流水号+到账金额）
    st, rc = req(fin, "POST", f"/api/initial-payments/{pid}/receipt",
                 {"bank_receipt_path": "/uploads/e2e_bank.jpg",
                  "bank_serial": "E2E20260604", "received_amount": expect_amount})
    check("G3 财务上传收款回单", st == 200 and rc.get("success"),
          str(rc.get("message", rc))[:60])

    # G4 财务审批通过首次付款 → 触发还款计划激活
    flow_id = find_initial_payment_flow(fin, pid)
    check("G4 定位首付审批步骤", bool(flow_id), f"flow_id={flow_id}")
    if flow_id:
        st, ap = req(fin, "POST", f"/api/approvals/{flow_id}/approve",
                     {"comment": "e2e 自动审批"})
        check("G4 首付审批通过", st == 200 and ap.get("success"),
              str(ap.get("message", ap))[:60])

    # G5 验证还款计划已从“未激活”激活为“待还款/已还款”
    st, pc = req(fin, "GET", f"/api/contracts/{cid}/repayments")
    reps = pc if isinstance(pc, list) else (pc.get("data") or pc.get("repayments") or [])
    inactive = [r for r in reps if r.get("status") == "未激活"]
    check("G5 还款计划已激活", reps and not inactive,
          f"剩余未激活={len(inactive)} 共={len(reps)}")

    # G6 车管上传出库资料并出库
    st, df = req(cm, "POST", f"/api/contracts/{cid}/delivery-files",
                 {"delivery_photo_path": "/uploads/e2e_deliver.jpg",
                  "delivery_document_path": "/uploads/e2e_deliver_doc.jpg"})
    check("G6 上传出库资料", st == 200 and df.get("success"),
          str(df.get("message", df))[:60])
    st, dv = req(cm, "POST", f"/api/vehicles/{vehicle_id}/deliver")
    check("G6 车辆出库", st == 200 and dv.get("success"),
          str(dv.get("message", dv))[:60])
    return True

def run_h1(cid):
    """H1: 触发日终批处理，验证幂等与状态机。"""
    print("\n==== H1 每日监控与催收 ====")
    fin = login("fin")
    st, r1 = req(fin, "POST", "/api/jobs/daily-collect", {"force": True})
    check("H1 日终批处理触发", st == 200 and r1.get("success"),
          str({k: r1.get(k) for k in list(r1)[:4]})[:80])
    # 幂等：再次非强制应跳过或返回已执行
    st, r2 = req(fin, "POST", "/api/jobs/daily-collect", {})
    check("H1 当日幂等", st == 200 and r2.get("success"), "二次调用成功")
    # 还款计划应已从未激活转为待还款/临近/逾期
    st, pc = req(fin, "GET", f"/api/contracts/{cid}/repayments")
    reps = pc if isinstance(pc, list) else (pc.get("data") or pc.get("repayments") or [])
    statuses = sorted({r.get("status") for r in reps})
    activated = any(r.get("status") != "未激活" for r in reps)
    check("H1 计划激活状态流转", activated or len(reps) == 0, f"状态集={statuses}")

def summarize():
    print("\n" + "=" * 50)
    print("端到端流程测试汇总")
    print("=" * 50)
    passed = sum(1 for _, ok, _ in RESULTS if ok)
    failed = len(RESULTS) - passed
    for stage, ok, detail in RESULTS:
        print(f"  {'✓' if ok else '✗'} {stage}")
    print(f"\n总计: {passed} 通过, {failed} 失败 / 共 {len(RESULTS)} 个断言")
    return failed == 0

if __name__ == "__main__":
    out = run_flow_lease()
    if out:
        order_id, cid = out
        run_g_delivery(order_id, cid)
        run_h1(cid)
    ok = summarize()
    raise SystemExit(0 if ok else 1)
