#!/usr/bin/env python3
"""6.2 完整流程测试"""
import urllib.request
import urllib.parse
import json

BASE_URL = "http://localhost:49165"
cookies = {}

def api_get(path):
    url = BASE_URL + path
    req = urllib.request.Request(url)
    cookie_str = '; '.join([f"{k}={v}" for k, v in cookies.items()])
    if cookie_str:
        req.add_header('Cookie', cookie_str)
    with urllib.request.urlopen(req) as response:
        for header in response.headers.get_all('Set-Cookie') or []:
            parts = header.split(';')[0].split('=')
            if len(parts) == 2:
                cookies[parts[0].strip()] = parts[1].strip()
        return json.loads(response.read())

def api_post(path, data):
    url = BASE_URL + path
    encoded_data = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(url, data=encoded_data, headers={'Content-Type': 'application/json'})
    cookie_str = '; '.join([f"{k}={v}" for k, v in cookies.items()])
    if cookie_str:
        req.add_header('Cookie', cookie_str)
    with urllib.request.urlopen(req) as response:
        for header in response.headers.get_all('Set-Cookie') or []:
            parts = header.split(';')[0].split('=')
            if len(parts) == 2:
                cookies[parts[0].strip()] = parts[1].strip()
        return json.loads(response.read())

def login(username, password):
    return api_post("/api/auth/login", {"username": username, "password": password})

def main():
    """主测试流程"""
    print("金聚源租车管理系统 - 6.2 完整流程测试")
    print("=" * 50)

    results = {}

    # 模块 A - 用户与权限
    print("\n=== 模块 A - 用户与权限 ===")
    try:
        result = login("boss", "123456")
        assert result["success"], "登录失败"
        assert result["user"]["role"] == "老板", "角色错误"
        assert "dashboard" in result["user"]["pages"], "页面权限缺失"
        print("✓ A1 登录成功")
        print("✓ A2 角色权限正确")
        results["A"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["A"] = "失败"

    # 模块 B - 车辆资产管理
    print("\n=== 模块 B - 车辆资产管理 ===")
    try:
        vehicles = api_get("/api/vehicles")
        assert len(vehicles) > 0, "无车辆数据"
        v = vehicles[0]
        assert "status" in v, "车辆状态字段缺失"
        assert "lock_status" in v, "锁定状态字段缺失"
        assert v["lock_status"] in ["未锁", "锁车流程中", "车辆已锁", "开锁流程中"], f"锁状态异常: {v['lock_status']}"
        print(f"✓ B1 车辆查询成功，共{len(vehicles)}辆")
        print(f"✓ B2 车辆状态: {v['status']}，锁定: {v['lock_status']}")
        results["B"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["B"] = "失败"

    # 模块 C - 指导价管理
    print("\n=== 模块 C - 指导价管理 ===")
    try:
        prices = api_get("/api/model-guidance-prices")
        print(f"✓ C1 指导价查询成功，共{len(prices)}条")
        stats = api_get("/api/stats")
        print(f"✓ C2 老板看板数据获取成功")
        results["C"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["C"] = "失败"

    # 模块 D - 销售报单
    print("\n=== 模块 D - 销售报单 ===")
    try:
        orders = api_get("/api/sales-orders")
        valid_statuses = ["草稿", "待价格特批", "待财务确认", "已激活", "已作废"]
        for order in orders:
            if order.get("order_status"):
                assert order["order_status"] in valid_statuses, f"状态异常: {order['order_status']}"
        print(f"✓ D1 报单查询成功，共{len(orders)}条")
        print(f"✓ D2 报单状态字典正确")
        results["D"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["D"] = "失败"

    # 模块 F - 还款计划
    print("\n=== 模块 F - 还款计划 ===")
    try:
        orders = api_get("/api/sales-orders")
        activated = [o for o in orders if o.get("order_status") == "已激活"]
        if activated:
            oid = activated[0]["id"]
            repayments = api_get(f"/api/repayments?order_id={oid}")
            factory = api_get(f"/api/factory-repayments?order_id={oid}")
            print(f"✓ F1 客户计划: {len(repayments)}期")
            print(f"✓ F2 厂家计划: {len(factory)}期")
            if repayments:
                statuses = set([rp.get("status") for rp in repayments])
                print(f"✓ F3 还款状态: {statuses}")
        else:
            print("⚠ 无已激活报单")
        results["F"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["F"] = "失败"

    # 模块 H - 每日监控
    print("\n=== 模块 H - 每日监控与催收 ===")
    try:
        r = api_post("/api/jobs/daily-collect", {})
        assert "success" in r, "日终批处理异常"
        print(f"✓ H1 日终批处理: {r.get('message', '')[:50]}")
        lock_reqs = api_get("/api/lock-requests")
        print(f"✓ H4 锁车/解锁查询: {len(lock_reqs)}条")
        results["H"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["H"] = "失败"

    # 模块 J - 滞纳金
    print("\n=== 模块 J - 滞纳金 ===")
    try:
        ledger = api_get("/api/late-fee-ledger")
        waivers = api_get("/api/waivers")
        print(f"✓ J1 滞纳金台账: {len(ledger)}条")
        print(f"✓ J3 减免申请: {len(waivers)}条")
        results["J"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["J"] = "失败"

    # 模块 L - 发票
    print("\n=== 模块 L - 发票申请 ===")
    try:
        invoices = api_get("/api/invoice-requests")
        print(f"✓ L1 发票申请: {len(invoices)}条")
        results["L"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["L"] = "失败"

    # 模块 M - 退车
    print("\n=== 模块 M - 退车流程 ===")
    try:
        inspections = api_get("/api/return-inspections")
        print(f"✓ M 退车记录: {len(inspections)}条")
        results["M"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["M"] = "失败"

    # 模块 N - 过户
    print("\n=== 模块 N - 过户流程 ===")
    try:
        transfers = api_get("/api/ownership-transfers")
        print(f"✓ N 过户记录: {len(transfers)}条")
        results["N"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["N"] = "失败"

    # 模块 O - 车辆返利
    print("\n=== 模块 O - 车辆返利 ===")
    try:
        rebates = api_get("/api/vehicle-rebates")
        print(f"✓ O1 车辆返利: {len(rebates)}条")
        results["O"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["O"] = "失败"

    # 模块 R - 收款公司
    print("\n=== 模块 R - 收款公司主数据 ===")
    try:
        companies = api_get("/api/payment-companies")
        print(f"✓ R1 收款公司: {len(companies)}家")
        results["R"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["R"] = "失败"

    # 模块 S - 客户黑名单
    print("\n=== 模块 S - 客户黑名单 ===")
    try:
        blacklist = api_get("/api/customer-blacklist")
        valid_levels = ["警示", "拒绝"]
        for item in blacklist:
            if item.get("level"):
                assert item["level"] in valid_levels, f"等级异常: {item['level']}"
        print(f"✓ S1 黑名单: {len(blacklist)}条")
        print(f"✓ S2 等级字段正确")
        results["S"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["S"] = "失败"

    # 模块 Q - 老板看板
    print("\n=== 模块 Q - 老板穿透看板 ===")
    try:
        stats = api_get("/api/stats")
        assert "total_vehicles" in stats, "字段缺失"
        print(f"✓ Q1 看板数据: {stats}")
        results["Q"] = "通过"
    except Exception as e:
        print(f"✗ 异常: {e}")
        results["Q"] = "失败"

    # 汇总
    print("\n" + "=" * 50)
    print("测试汇总")
    print("=" * 50)

    passed = sum(1 for v in results.values() if v == "通过")
    failed = len(results) - passed

    for module, result in results.items():
        status = "✓" if result == "通过" else "✗"
        print(f"模块 {module}: {status} {result}")

    print(f"\n总计: {passed}通过, {failed}失败")

    return failed == 0

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
