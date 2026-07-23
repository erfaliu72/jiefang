#!/usr/bin/env python3
"""
清空 RDS MySQL 所有业务数据
1. 先杀应用进程释放数据库连接
2. 按依赖顺序 DELETE 所有业务表
3. 清理上传文件
4. 重启应用
"""
import pymysql
import os
import subprocess
import signal
import time

# RDS 连接信息
DB_CONF = dict(
    host='rm-2vc15as5od6ay3gl4zo.mysql.cn-chengdu.rds.aliyuncs.com',
    port=3306,
    user='dbroot',
    password='fijSec-tuccyt-9dohxu',
    database='jinjuyuan',
    charset='utf8mb4',
)

BUSINESS_TABLES = [
    'reconciliation_allocations', 'late_fee_ledger', 'contract_fee_items',
    'contract_initial_payments', 'factory_repayments', 'waivers',
    'invoice_requests', 'return_inspections', 'ownership_transfers',
    'vehicle_rebates', 'attachments', 'urge_records', 'lock_requests',
    'approval_flows', 'daily_job_runs', 'customer_prepayments',
    'receivables', 'customer_blacklist', 'customers', 'sales_orders',
    'contracts', 'vehicles', 'audit_logs',
]

KEEP_TABLES = [
    'users', 'role_pages', 'role_actions', 'role_field_permissions',
    'model_guidance_prices', 'model_guidance_price_history',
    'receiving_companies',
]

def stop_app():
    """杀掉 gunicorn 进程释放数据库连接"""
    for cmd in ['pkill -f gunicorn', 'pkill -f wsgi:app']:
        subprocess.run(cmd, shell=True, stderr=subprocess.DEVNULL)
    time.sleep(2)
    print('  ✅ App stopped')


def clear_database():
    """清空所有业务表"""
    conn = pymysql.connect(**DB_CONF)
    c = conn.cursor()
    c.execute('SET FOREIGN_KEY_CHECKS=0')
    ok, fail = 0, 0
    for t in BUSINESS_TABLES:
        try:
            c.execute('SELECT COUNT(*) FROM %s' % t)
            cnt = c.fetchone()[0]
            if cnt > 0:
                c.execute('DELETE FROM %s' % t)
                conn.commit()
                print('  ✅ %s: deleted %d rows' % (t, cnt))
                ok += 1
            else:
                print('  ⏭️  %s: empty' % t)
        except Exception as e:
            print('  ❌ %s: %s' % (t, e))
            conn.rollback()
            fail += 1
    c.execute('SET FOREIGN_KEY_CHECKS=1')

    # 验证保留的表
    print()
    for t in KEEP_TABLES:
        try:
            c.execute('SELECT COUNT(*) FROM %s' % t)
            print('  📦 %s: %d rows (kept)' % (t, c.fetchone()[0]))
        except:
            print('  ⏭️  %s: not found' % t)

    conn.commit()
    conn.close()
    print('\n  ✅ Done: %d ok, %d failed' % (ok, fail))


def clean_uploads():
    """清理上传文件"""
    for d in ['/opt/jinjuyuan/uploads', '/opt/jinjuyuan/generated-contracts']:
        if os.path.exists(d):
            for f in os.listdir(d):
                fp = os.path.join(d, f)
                if os.path.isfile(fp):
                    os.remove(fp)
            print('  ✅ Cleaned %s' % d)


def restart_app():
    """重启应用"""
    # 尝试 systemd，否则手动启动
    r = subprocess.run('systemctl restart jinjuyuan', shell=True,
                       stderr=subprocess.DEVNULL)
    if r.returncode != 0:
        # 手动启动
        env = os.environ.copy()
        env.update({
            'JJY_DB_HOST': DB_CONF['host'],
            'JJY_DB_PORT': str(DB_CONF['port']),
            'JJY_DB_USER': DB_CONF['user'],
            'JJY_DB_PASSWORD': DB_CONF['password'],
            'JJY_DB_NAME': DB_CONF['database'],
        })
        subprocess.Popen(
            ['gunicorn', '-b', '0.0.0.0:8081', '-w', '2', '--timeout', '120',
             'wsgi:app'],
            cwd='/opt/jinjuyuan',
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    time.sleep(4)

    # 验证
    r = subprocess.run(
        'curl -s -o /dev/null -w %{http_code} http://127.0.0.1:8081/',
        shell=True, capture_output=True, text=True, timeout=10)
    print('  ✅ App restarted, HTTP status:', r.stdout)


if __name__ == '__main__':
    print('Step 1: Stop app...')
    stop_app()

    print('Step 2: Clear database...')
    clear_database()

    print('Step 3: Clean uploads...')
    clean_uploads()

    print('Step 4: Restart app...')
    restart_app()

    print('\n✅ ALL COMPLETE!')
    print('   Visit: http://8.156.91.120:8081')
