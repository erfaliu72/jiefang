# -*- coding: utf-8 -*-
"""按 6.2 主流程图分阶段端到端测试。
覆盖三类业务（整车销售 / 经营租赁 / 以租代售）的全链路：
入库去重 → 报单/指导口径/价格特批 → 财务确认激活 → 上传线下合同 →
首次付款（足额/不足额+老板审批）→ 财务对账 → 车管出库 → 每期还款（等额/少还+滞纳金）→
退车（租赁）/ 过户（以租代售）/ 结清（销售）。

每个阶段独立记录断言结果，最后输出分阶段测试报告。
"""
import json, os, random, string, datetime
import urllib.request, urllib.error, http.cookiejar

BASE = os.environ.get("JJY_BASE", "http://127.0.0.1:49169")
RESULTS = []          # (stage_group, case, ok, detail)
CURRENT_GROUP = "-"


def set_group(name):
    global CURRENT_GROUP
    CURRENT_GROUP = name
    print(f"\n========== {name} ==========")


def check(case, ok, detail=""):
    RESULTS.append((CURRENT_GROUP, case, bool(ok), detail))
    print(f"  {'PASS' if ok else 'FAIL'}  {case}  {detail}")
    return ok


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
        with opener.open(r, timeout=20) as resp:
            raw = resp.read().decode()
            return resp.status, (json.loads(raw) if raw else {})
    except urllib.error.HTTPError as e:
        raw = e.read().decode()
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, {"raw": raw}


def login(user):
    op = _client()
    st, body = req(op, "POST", "/api/auth/login", {"username": user, "password": "123456"})
    assert st == 200 and body.get("success"), f"登录失败 {user}: {body}"
    return op


def rand_vin():
    chars = string.ascii_uppercase.replace("I", "").replace("O", "").replace("Q", "") + string.digits
    return "TEST" + "".join(random.choice(chars) for _ in range(13))


def as_list(body):
    return body if isinstance(body, list) else (body.get("data") or [])


def find_order(op, order_id):
    st, lst = req(op, "GET", "/api/sales-orders")
    return next((o for o in as_list(lst) if o.get("id") == order_id), {})


def find_contract(op, cid):
    st, lst = req(op, "GET", "/api/contracts")
    return next((x for x in as_list(lst) if x.get("id") == cid), {})


def pending_flow_id(op, ref_type, ref_id, role):
    st, body = req(op, "GET", f"/api/approvals?ref_type={ref_type}")
    for it in as_list(body):
        if it.get("ref_id") == ref_id:
            for s in it.get("steps", []):
                if s.get("status") == "待审批" and (role is None or s.get("required_role") == role):
                    return s.get("id")
    return None


# ---------------------------------------------------------------------------
# 角色客户端（全局复用）
# ---------------------------------------------------------------------------
boss = login("boss")
ops = login("ops")
fin = login("fin")
fleet = login("fleet")
sales = login("sales")

TODAY = datetime.date.today()
CAR_TYPE = "测试车型-J6P"          # 专用测试车型，避免污染真实车型口径
GUIDANCE = {
    "car_type": CAR_TYPE,
    "lease_installment_price": 5000.0,
    "sale_total_price": 300000.0,
    "lease_deposit_ratio": 0.10,
    "lease_repayment_ratio": 0.015,
    "sale_down_payment_ratio": 0.20,
    "sale_repayment_ratio": 0.02,
    "remark": "分阶段测试口径",
}


# ===========================================================================
# 阶段 A：车辆入库 + 去重
# ===========================================================================
def stage_a_intake():
    set_group("阶段A 车辆入库与去重")
    vin = rand_vin()
    st, r = req(fleet, "POST", "/api/vehicles", {
        "vin": vin, "plate_number": "陕A" + "".join(random.choice(string.digits) for _ in range(5)),
        "car_type": CAR_TYPE, "is_new": "新车",
    })
    check("A1 车管新车入库", st == 200 and r.get("success"), str(r.get("message", r))[:50])

    # 去重：同 VIN 再次入库应被拒绝
    st2, r2 = req(fleet, "POST", "/api/vehicles", {"vin": vin, "car_type": CAR_TYPE})
    check("A2 同VIN去重拦截", st2 != 200 or not r2.get("success"), str(r2.get("message", r2))[:40])

    # 校验入库后状态=在库
    st3, vehicles = req(sales, "GET", "/api/vehicles/list")
    v = next((x for x in as_list(vehicles) if x.get("vin") == vin), {})
    check("A3 入库后状态=在库", v.get("status") == "在库", f"status={v.get('status')}")
    return vin


def ensure_guidance():
    set_group("阶段A' 老板维护车型指导口径")
    st, r = req(boss, "POST", "/api/model-guidance-prices", GUIDANCE)
    check("A'1 维护双指导价+四比例", st == 200 and r.get("success"), str(r.get("message", r))[:50])


# ===========================================================================
# 阶段 B：销售报单 + 指导口径校验 + 价格特批
# ===========================================================================
def stage_b_order(vin, sales_mode, payload, expect_status):
    base = {
        "vin": vin, "sales_mode": sales_mode,
        "customer_name": f"测试客户{random.randint(1000,9999)}",
        "customer_phone": "139" + "".join(random.choice(string.digits) for _ in range(8)),
        "payment_date": TODAY.strftime("%Y-%m-%d"),
        "car_type": CAR_TYPE,
    }
    base.update(payload)
    st, r = req(sales, "POST", "/api/sales-orders", base)
    ok = st == 200 and r.get("success")
    check(f"B 报单创建[{sales_mode}]", ok, str(r.get("message", r))[:55])
    if not ok:
        return None
    oid = r["id"]
    cur = find_order(sales, oid)
    check(f"B 报单状态={expect_status}[{sales_mode}]",
          cur.get("order_status") == expect_status,
          f"实际={cur.get('order_status')}")
    return oid


def boss_price_approve(order_id, approve=True):
    fid = pending_flow_id(boss, "price_exception", order_id, "老板")
    if not check("B 定位价格特批步骤", bool(fid), f"flow={fid}"):
        return False
    if approve:
        st, r = req(boss, "POST", f"/api/approvals/{fid}/approve", {"comment": "测试通过特批"})
        ok = st == 200 and r.get("success")
        check("B 老板价格特批通过", ok, str(r.get("message", r))[:45])
        if ok:
            cur = find_order(sales, order_id)
            check("B 特批后转待财务确认", cur.get("order_status") == "待财务确认",
                  f"实际={cur.get('order_status')}")
        return ok
    else:
        st, r = req(boss, "POST", f"/api/approvals/{fid}/reject", {"comment": "测试驳回特批"})
        ok = st == 200 and r.get("success")
        check("B 老板价格特批驳回", ok, str(r.get("message", r))[:45])
        if ok:
            cur = find_order(sales, order_id)
            check("B 驳回后报单作废/价格特批驳回",
                  cur.get("order_status") in ("已作废", "价格特批驳回"),
                  f"实际={cur.get('order_status')}")
        return ok


# ===========================================================================
# 阶段 C：财务确认激活报单（分期类需先有客户计划）
# ===========================================================================
def stage_c_activate(order_id, sales_mode):
    if sales_mode != "整车销售":
        st, pc = req(ops, "GET", f"/api/sales-orders/{order_id}/planning-contract")
        ok = st == 200 and pc.get("success") and len(pc.get("repayments", [])) > 0
        check(f"C 运营生成客户还款计划[{sales_mode}]", ok,
              f"期数={len(pc.get('repayments', []))}")
    st, r = req(fin, "POST", f"/api/sales-orders/{order_id}/activate")
    ok = st == 200 and r.get("success")
    check(f"C 财务确认激活报单[{sales_mode}]", ok, str(r.get("message", r))[:45])
    if ok:
        cur = find_order(sales, order_id)
        check(f"C 报单状态=已激活[{sales_mode}]", cur.get("order_status") == "已激活",
              f"实际={cur.get('order_status')}")
    return ok


# ===========================================================================
# 阶段 D：运营上传线下合同
# ===========================================================================
def stage_d_contract(order_id, sales_mode, vin, contract_overrides=None):
    cur = find_order(ops, order_id)
    vehicle_id = cur.get("vehicle_id")
    body = {
        "vehicle_id": vehicle_id, "sales_order_id": order_id,
        "contract_file": "/uploads/test_contract.pdf",
        "start_date": TODAY.strftime("%Y-%m-%d"),
    }
    if contract_overrides:
        body.update(contract_overrides)  # 允许覆盖 start_date 以构造逾期场景
    st, r = req(ops, "POST", "/api/contracts", body)
    ok = st == 200 and r.get("success")
    check(f"D 上传线下合同[{sales_mode}]", ok, str(r.get("message", r))[:45])
    if not ok:
        return None, vehicle_id
    cid = r["id"]
    c = find_contract(ops, cid)
    check(f"D 合同状态=待首付款[{sales_mode}]", c.get("delivery_status") == "待首付款",
          f"实际={c.get('delivery_status')}")
    return cid, vehicle_id


# ===========================================================================
# 阶段 E：首次付款 + 财务对账（足额 / 不足额）
# ===========================================================================
def stage_e_initial_payment(cid, sales_mode, shortage=False):
    c = find_contract(ops, cid)
    if sales_mode == "整车销售":
        expect = float(c.get("total_price") or c.get("down_payment") or 0)
    elif sales_mode == "以租代售":
        expect = float(c.get("down_payment") or 0)
    else:
        expect = float(c.get("deposit") or 0) + float(c.get("rent") or 0)
    check(f"E 首付应收金额计算[{sales_mode}]", expect > 0, f"应收={expect}")

    # 运营发起首付
    st, ip = req(ops, "POST", f"/api/contracts/{cid}/initial-payment", {
        "customer_screenshot_path": "/uploads/test_cust_pay.jpg",
        "amount": expect,
    })
    ok = st == 200 and ip.get("success")
    check(f"E 运营发起首次付款[{sales_mode}]", ok, str(ip.get("message", ip))[:45])
    if not ok:
        return False
    pid = ip["id"]

    received = round(expect * 0.6, 2) if shortage else expect
    receipt = {
        "bank_receipt_path": "/uploads/test_bank.jpg",
        "bank_serial": "TEST" + "".join(random.choice(string.digits) for _ in range(8)),
        "received_amount": received,
    }
    if shortage:
        receipt["shortage_reason"] = "客户资金周转，承诺补足"
        receipt["promised_repay_date"] = (TODAY + datetime.timedelta(days=15)).strftime("%Y-%m-%d")
    st, rc = req(fin, "POST", f"/api/initial-payments/{pid}/receipt", receipt)
    check(f"E 财务上传收款回单[{sales_mode}{'-不足' if shortage else ''}]",
          st == 200 and rc.get("success"), str(rc.get("message", rc))[:45])

    # 财务审批首付步骤
    fid = pending_flow_id(fin, "initial_payment", pid, "财务")
    check(f"E 定位首付审批步骤[{sales_mode}]", bool(fid), f"flow={fid}")
    if fid:
        st, ap = req(fin, "POST", f"/api/approvals/{fid}/approve", {"comment": "测试首付审核"})
        check(f"E 财务首付审核[{sales_mode}]", st == 200 and ap.get("success"),
              str(ap.get("message", ap))[:55])

    if shortage:
        # 老板审批不足额出库
        c2 = find_contract(ops, cid)
        check("E 不足额转老板审批", c2.get("delivery_status") == "首付不足待审批",
              f"实际={c2.get('delivery_status')}")
        bfid = pending_flow_id(boss, "initial_payment_shortage", pid, "老板")
        check("E 定位不足额老板审批步骤", bool(bfid), f"flow={bfid}")
        if bfid:
            st, ba = req(boss, "POST", f"/api/approvals/{bfid}/approve", {"comment": "同意不足额出库"})
            check("E 老板同意不足额出库", st == 200 and ba.get("success"),
                  str(ba.get("message", ba))[:45])

    c3 = find_contract(ops, cid)
    check(f"E 首付完成后合同=待出库[{sales_mode}]", c3.get("delivery_status") == "待出库",
          f"实际={c3.get('delivery_status')}")
    return c3.get("delivery_status") == "待出库"


# ===========================================================================
# 阶段 F：车管出库
# ===========================================================================
def stage_f_deliver(cid, vehicle_id, sales_mode):
    st, df = req(fleet, "POST", f"/api/contracts/{cid}/delivery-files", {
        "delivery_photo_path": "/uploads/test_deliver.jpg",
        "delivery_document_path": "/uploads/test_deliver_doc.jpg",
    })
    check(f"F 车管上传出库资料[{sales_mode}]", st == 200 and df.get("success"),
          str(df.get("message", df))[:45])
    st, dv = req(fleet, "POST", f"/api/vehicles/{vehicle_id}/deliver")
    ok = st == 200 and dv.get("success")
    check(f"F 车管确认出库[{sales_mode}]", ok, str(dv.get("message", dv))[:45])
    if ok:
        st, vehicles = req(sales, "GET", "/api/vehicles/list")
        v = next((x for x in as_list(vehicles) if x.get("id") == vehicle_id), {})
        expect_status = {"整车销售": "已售", "经营租赁": "租赁中", "以租代售": "以租代售"}[sales_mode]
        check(f"F 出库后车辆状态={expect_status}[{sales_mode}]",
              v.get("status") == expect_status, f"实际={v.get('status')}")
    return ok


# ===========================================================================
# 阶段 G：每期还款对账（等额 / 少还+滞纳金）
# ===========================================================================
def stage_g_repayments(cid, sales_mode, partial_last=False):
    st, reps = req(fin, "GET", f"/api/contracts/{cid}/repayments")
    reps = as_list(reps) if not isinstance(reps, list) else reps
    periods = [r for r in reps if (r.get("period") or 0) >= 1]
    check(f"G 还款计划已激活[{sales_mode}]",
          periods and all(r.get("status") != "未激活" for r in periods),
          f"期数={len(periods)}")
    if not periods:
        return

    # 租赁首期租金随首付款一并核销，故仅对"未还款"的期逐期对账
    pending = [r for r in periods if r.get("status") != "已还款"]
    prepaid = len(periods) - len(pending)
    if prepaid:
        check(f"G 首付已核销首期租金[{sales_mode}]", True, f"首付随收 {prepaid} 期")
    n = len(pending)
    confirmed = 0
    fail_msgs = []
    for idx, rp in enumerate(pending):
        rid = rp["id"]
        amount = float(rp.get("amount") or 0)
        last = idx == n - 1
        received = round(amount * 0.5, 2) if (partial_last and last) else amount
        # 上传截图(运营)+回单(财务)，满足对账门禁
        req(ops, "POST", f"/api/reconciliation/{rid}/screenshot",
            {"screenshot_path": "/uploads/test_rep_shot.jpg"})
        req(fin, "POST", f"/api/reconciliation/{rid}/receipt",
            {"bank_receipt_path": "/uploads/test_rep_bank.jpg"})
        st, vr = req(fin, "POST", f"/api/reconciliation/{rid}/verify", {
            "bank_serial": "REP" + "".join(random.choice(string.digits) for _ in range(7)),
            "received_amount": received,
        })
        if st == 200 and vr.get("success"):
            confirmed += 1
        else:
            fail_msgs.append(f"P{rp.get('period')}:{vr.get('message', vr)}")
    detail = f"成功核销 {confirmed}/{n} 期(首付已核销{prepaid}期)"
    if fail_msgs:
        detail += " | " + "; ".join(fail_msgs[:3])
    check(f"G 逐期对账核销[{sales_mode}]", confirmed >= (n - 1 if partial_last else n), detail)

    if partial_last:
        # 触发日终批处理累计滞纳金
        req(fin, "POST", "/api/jobs/daily-collect", {"force": True})
        st, reps2 = req(fin, "GET", f"/api/contracts/{cid}/repayments")
        reps2 = as_list(reps2) if not isinstance(reps2, list) else reps2
        last_rp = [r for r in reps2 if (r.get("period") or 0) >= 1][-1]
        check("G 少还期差额挂账为应收",
              float(last_rp.get("paid_amount") or 0) < float(last_rp.get("amount") or 0),
              f"已收={last_rp.get('paid_amount')} 应收={last_rp.get('amount')}")


# ===========================================================================
# 阶段 H：退车流程（租赁）
# ===========================================================================
def stage_h_return(vehicle_id):
    set_group("阶段H 退车流程（租赁）")
    st, r = req(sales, "POST", "/api/return-inspections", {
        "vehicle_id": vehicle_id, "return_reason": "到期退车",
    })
    ok = st == 200 and r.get("success")
    check("H1 销售发起退车", ok, str(r.get("message", r))[:45])
    if not ok:
        return
    rid = r["id"]
    st, rf = req(fleet, "POST", f"/api/return-inspections/{rid}/fleet", {
        "tool_triangle": 1, "doc_keys": 1, "mileage": "120000",
        "body_tire_clean": "正常", "needs_repair": 0,
    })
    check("H2 车管验车", st == 200 and rf.get("success"), str(rf.get("message", rf))[:40])
    st, ro = req(ops, "POST", f"/api/return-inspections/{rid}/operator", {
        "rent_late_fee": 0, "deposit_paid": 5000, "total_deduction": 0, "actual_refund": 5000,
    })
    check("H3 运营填写退车数据", st == 200 and ro.get("success"), str(ro.get("message", ro))[:40])
    st, rfi = req(fin, "POST", f"/api/return-inspections/{rid}/finance", {})
    check("H4 财务复核", st == 200 and rfi.get("success"), str(rfi.get("message", rfi))[:40])
    st, rb = req(boss, "POST", f"/api/return-inspections/{rid}/boss-approve", {"comment": "同意退款"})
    check("H5 老板领导审批", st == 200 and rb.get("success"), str(rb.get("message", rb))[:40])
    st, rp = req(fin, "POST", f"/api/return-inspections/{rid}/pay", {})
    ok = st == 200 and rp.get("success")
    check("H6 财务出款完成", ok, str(rp.get("message", rp))[:40])
    if ok:
        st, vehicles = req(sales, "GET", "/api/vehicles/list")
        v = next((x for x in as_list(vehicles) if x.get("id") == vehicle_id), {})
        check("H7 退车后车辆回库/待维修",
              v.get("status") in ("在库", "待维修"), f"实际={v.get('status')}")


# ===========================================================================
# 阶段 I：过户流程（以租代售）
# ===========================================================================
def stage_i_transfer(cid, vehicle_id):
    set_group("阶段I 过户流程（以租代售）")
    st, r = req(ops, "POST", f"/api/contracts/{cid}/ownership-transfer", {"settle_type": "natural_settle"})
    ok = st == 200 and r.get("success")
    check("I1 运营发起过户", ok, str(r.get("message", r))[:45])
    if not ok:
        return
    tid = r["id"]
    st, rc = req(ops, "POST", f"/api/ownership-transfers/{tid}/complete", {
        "new_owner_name": "测试受让人", "new_owner_id_card": "610000199001011234",
        "transfer_doc_path": "/uploads/test_transfer.pdf",
    })
    ok = st == 200 and rc.get("success")
    check("I2 运营登记过户完成", ok, str(rc.get("message", rc))[:45])
    if ok:
        c = find_contract(ops, cid)
        check("I3 合同状态=已结清", c.get("contract_status") == "已结清",
              f"实际={c.get('contract_status')}")
        st, vehicles = req(sales, "GET", "/api/vehicles/list")
        v = next((x for x in as_list(vehicles) if x.get("id") == vehicle_id), {})
        check("I4 车辆状态=已售/已过户", v.get("status") == "已售/已过户",
              f"实际={v.get('status')}")


# ===========================================================================
# 阶段 J：滞纳金减免（销售发起 → 老板审批 → 财务执行）
# ===========================================================================
def stage_j_waiver(cid):
    set_group("阶段J 滞纳金减免")
    st, r = req(sales, "POST", f"/api/contracts/{cid}/waivers", {
        "waiver_kind": "late_fee", "target_period_list": "all",
        "reason": "客户特殊情况，申请减免滞纳金",
    })
    ok = st == 200 and r.get("success")
    check("J1 销售发起减免申请", ok, str(r.get("message", r))[:45])
    if not ok:
        return
    wid = r["waiver_id"]
    st, ra = req(boss, "POST", f"/api/waivers/{wid}/approve", {"decision": "approve", "comment": "同意减免"})
    check("J2 老板审批减免通过", st == 200 and ra.get("success"), str(ra.get("message", ra))[:45])
    st, re = req(fin, "POST", f"/api/waivers/{wid}/execute", {})
    check("J3 财务执行减免", st == 200 and re.get("success"), str(re.get("message", re))[:55])


# ===========================================================================
# 业务编排
# ===========================================================================
def run_sale():
    """整车销售：A→B(无特批)→C→D→E(足额)→F→结清。"""
    vin = stage_a_intake()
    set_group("阶段B-F 整车销售全链路")
    oid = stage_b_order(vin, "整车销售", {"sale_total_price": 280000, "deposit_amount": 280000},
                        expect_status="待财务确认")
    if not oid:
        return
    if not stage_c_activate(oid, "整车销售"):
        return
    cid, vid = stage_d_contract(oid, "整车销售", vin,
                                {"total_price": 280000, "down_payment": 280000})
    if not cid:
        return
    if not stage_e_initial_payment(cid, "整车销售", shortage=False):
        return
    if stage_f_deliver(cid, vid, "整车销售"):
        c = find_contract(ops, cid)
        check("整车销售出库后=已结清", c.get("contract_status") == "已结清",
              f"实际={c.get('contract_status')}")


def run_lease():
    """经营租赁：A→B(达标无特批)→C→D→E(足额)→F→G(等额)→H 退车。"""
    vin = stage_a_intake()
    set_group("阶段B-G 经营租赁全链路")
    # 押金=总价*10%、每期>=总价*1.5% → 达标，无需特批
    oid = stage_b_order(vin, "经营租赁", {
        "sale_total_price": 300000, "vehicle_rent_amount": 5000,
        "deposit_amount": 30000, "lease_term": "3",
    }, expect_status="待财务确认")
    if not oid:
        return
    if not stage_c_activate(oid, "经营租赁"):
        return
    cid, vid = stage_d_contract(oid, "经营租赁", vin, {
        "rent": 5000, "deposit": 30000, "loan_periods": 3,
    })
    if not cid:
        return
    if not stage_e_initial_payment(cid, "经营租赁", shortage=False):
        return
    if not stage_f_deliver(cid, vid, "经营租赁"):
        return
    set_group("阶段G 经营租赁每期还款")
    stage_g_repayments(cid, "经营租赁", partial_last=False)
    stage_h_return(vid)


def run_lease_to_sale():
    """以租代售：A→B(低于口径→老板特批)→C→D→E(不足额+老板审批)→F→G(少还+滞纳金)→J 减免→I 过户。"""
    vin = stage_a_intake()
    set_group("阶段B-G 以租代售全链路（含价格特批+不足额）")
    # 首付比例 10% < 指导 20% → 触发价格特批
    oid = stage_b_order(vin, "以租代售", {
        "sale_total_price": 300000, "vehicle_rent_amount": 6000,
        "deposit_amount": 30000, "lease_term": "3",
    }, expect_status="待价格特批")
    if not oid:
        return
    if not boss_price_approve(oid, approve=True):
        return
    if not stage_c_activate(oid, "以租代售"):
        return
    past_start = (TODAY - datetime.timedelta(days=120)).strftime("%Y-%m-%d")
    cid, vid = stage_d_contract(oid, "以租代售", vin, {
        "down_payment": 30000, "rent": 6000, "loan_periods": 3, "total_price": 300000,
        "start_date": past_start,
    })
    if not cid:
        return
    if not stage_e_initial_payment(cid, "以租代售", shortage=True):
        return
    if not stage_f_deliver(cid, vid, "以租代售"):
        return
    set_group("阶段G 以租代售每期还款（末期少还）")
    stage_g_repayments(cid, "以租代售", partial_last=True)
    stage_j_waiver(cid)
    # 末期少还，过户应被拦截（未结清）
    set_group("阶段I 过户前置校验")
    st, r = req(ops, "POST", f"/api/contracts/{cid}/ownership-transfer", {"settle_type": "natural_settle"})
    check("I0 未结清时过户被拦截", not (st == 200 and r.get("success")),
          str(r.get("message", r))[:50])


def run_price_reject():
    """价格特批驳回分支：报单低于口径 → 老板驳回 → 报单作废。"""
    vin = stage_a_intake()
    set_group("分支 价格特批驳回")
    oid = stage_b_order(vin, "以租代售", {
        "sale_total_price": 300000, "vehicle_rent_amount": 6000,
        "deposit_amount": 15000, "lease_term": "3",
    }, expect_status="待价格特批")
    if not oid:
        return
    boss_price_approve(oid, approve=False)


def summarize():
    print("\n" + "=" * 60)
    print("分阶段流程测试报告汇总")
    print("=" * 60)
    groups = {}
    for grp, case, ok, detail in RESULTS:
        groups.setdefault(grp, []).append((case, ok, detail))
    total_pass = sum(1 for _, _, ok, _ in RESULTS if ok)
    total = len(RESULTS)
    for grp, items in groups.items():
        p = sum(1 for _, ok, _ in items if ok)
        print(f"\n[{grp}]  {p}/{len(items)} 通过")
        for case, ok, detail in items:
            if not ok:
                print(f"    FAIL  {case}  {detail}")
    print(f"\n总计：{total_pass}/{total} 通过，{total - total_pass} 失败")
    return total_pass, total


if __name__ == "__main__":
    ensure_guidance()
    run_sale()
    run_lease()
    run_lease_to_sale()
    run_price_reject()
    p, t = summarize()
    raise SystemExit(0 if p == t else 1)
