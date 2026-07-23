#!/usr/bin/env python3
"""清空金聚源车辆管理系统所有业务数据

用法（本地 SQLite）:
  python3 clear_all_data.py

用法（外网 RDS MySQL，在 ECS 上执行）:
  JJY_DB_HOST=你的RDS地址 JJY_DB_PASSWORD=密码 python3 clear_all_data.py

保留系统配置表: users, role_pages, role_actions, role_field_permissions,
                model_guidance_prices, model_guidance_price_history, receiving_companies
清空所有业务数据表，并重置自增ID。
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from database import get_db

# ── 业务数据表（按外键依赖逆序，避免冲突）──
BUSINESS_TABLES = [
    'reconciliation_allocations',
    'late_fee_ledger',
    'urge_records',
    'lock_requests',
    'contract_fee_items',
    'contract_initial_payments',
    'factory_repayments',
    'repayments',
    'waivers',
    'invoice_requests',
    'return_inspections',
    'ownership_transfers',
    'vehicle_rebates',
    'attachments',
    'approval_flows',
    'daily_job_runs',
    'customer_prepayments',
    'receivables',
    'customer_blacklist',
    'customers',
    'sales_orders',
    'contracts',
    'vehicles',
    'audit_logs',
]

# ── 保留的系统配置表 ──
SYSTEM_TABLES = [
    'users',
    'role_pages',
    'role_actions',
    'role_field_permissions',
    'model_guidance_prices',
    'model_guidance_price_history',
    'receiving_companies',
]


def confirm():
    print('╔══════════════════════════════════════════════════════════╗')
    print('║   ⚠️  即将清空所有业务数据！此操作不可撤销！            ║')
    print('╠══════════════════════════════════════════════════════════╣')
    print(f'║  保留的表 ({len(SYSTEM_TABLES)} 张):')
    for t in SYSTEM_TABLES:
        print(f'║    ✅ {t}')
    print('║')
    print(f'║  清空的表 ({len(BUSINESS_TABLES)} 张):')
    for t in BUSINESS_TABLES:
        print(f'║    ❌ {t}')
    print('╚══════════════════════════════════════════════════════════╝')
    print()
    r = input('确认清空所有业务数据？(输入 YES 确认，其他任意键取消): ')
    return r == 'YES'


def table_exists(c, table, is_mysql):
    if is_mysql:
        c.execute("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema=DATABASE() AND table_name=?", (table,))
        return c.fetchone()[0] > 0
    else:
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table,))
        return c.fetchone() is not None


def clear_data():
    conn = get_db()
    c = conn.cursor()

    is_mysql = False
    try:
        c.execute("SELECT VERSION()")
        is_mysql = True
    except Exception:
        is_mysql = False

    db_type = 'MySQL' if is_mysql else 'SQLite'
    print(f'🔍 数据库类型: {db_type}')

    # 禁用外键检查
    if is_mysql:
        c.execute("SET FOREIGN_KEY_CHECKS = 0")
    else:
        c.execute("PRAGMA foreign_keys = OFF")

    counts = {}
    try:
        for table in BUSINESS_TABLES:
            if not table_exists(c, table, is_mysql):
                print(f'  ⏭️  {table} — 表不存在，跳过')
                continue

            c.execute(f"SELECT COUNT(*) FROM {table}")
            row_count = c.fetchone()[0]
            counts[table] = row_count

            if row_count == 0:
                print(f'  ⏭️  {table} — 已为空')
                continue

            c.execute(f"DELETE FROM {table}")
            print(f'  ✅  {table} — 删除了 {row_count} 条记录')

            # 重置自增 ID
            if is_mysql:
                c.execute(f"ALTER TABLE {table} AUTO_INCREMENT = 1")
            else:
                c.execute(f"DELETE FROM sqlite_sequence WHERE name='{table}'")

        conn.commit()

        # 统计保留表
        kept_counts = {}
        for table in SYSTEM_TABLES:
            try:
                if table_exists(c, table, is_mysql):
                    c.execute(f"SELECT COUNT(*) FROM {table}")
                    kept_counts[table] = c.fetchone()[0]
                else:
                    kept_counts[table] = 0
            except Exception:
                kept_counts[table] = 0

        # 恢复外键检查
        if is_mysql:
            c.execute("SET FOREIGN_KEY_CHECKS = 1")
        else:
            c.execute("PRAGMA foreign_keys = ON")

    except Exception as e:
        conn.rollback()
        print(f'\n❌ 清空失败: {e}')
        return False
    finally:
        conn.close()

    total_deleted = sum(counts.values())
    cleared_count = len([t for t in counts if counts[t] > 0])
    print()
    print('═' * 50)
    print(f'📊 清空完成！')
    print(f'   🗑️  清空 {cleared_count} 张业务表，共删除 {total_deleted} 条记录')
    print(f'   📦 保留 {len(kept_counts)} 张系统配置表:')
    for t, cnt in kept_counts.items():
        print(f'       ✅ {t} ({cnt} 条)')
    print('═' * 50)

    # 上传文件目录清理提示
    upload_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    if os.path.exists(upload_dir):
        files = [f for f in os.listdir(upload_dir) if os.path.isfile(os.path.join(upload_dir, f))]
        if files:
            print(f'\n💡 uploads/ 目录下还有 {len(files)} 个上传文件未清理')
            print(f'   如需清理，请执行: rm -rf {upload_dir}/*')

    return True


if __name__ == '__main__':
    if not confirm():
        print('❌ 已取消')
        sys.exit(0)
    ok = clear_data()
    sys.exit(0 if ok else 1)
