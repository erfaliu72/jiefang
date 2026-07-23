"""清理生产数据库中的指导价数据和指定VIN车辆"""
import paramiko, json, time

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

# 先写一个清理脚本到服务器
clean_script = r'''
import sqlite3, os
os.chdir('/opt/jinjuyuan')
conn = sqlite3.connect('jinjuyuan.db')
conn.row_factory = sqlite3.Row
c = conn.cursor()

# 1. 查看指导价相关表
c.execute("SELECT name FROM sqlite_master WHERE type='table' AND (name LIKE '%guidance%' OR name LIKE '%price%')")
print("指导价相关表:", [r[0] for r in c.fetchall()])

# 2. 查看 model_guidance_prices
c.execute("SELECT COUNT(*) FROM model_guidance_prices")
print("model_guidance_prices 当前行数:", c.fetchone()[0])

# 3. 清空 model_guidance_prices
c.execute("DELETE FROM model_guidance_prices")
print("已清空 model_guidance_prices, 影响行数:", c.rowcount)

# 4. 清空 history 表
c.execute("DELETE FROM model_guidance_price_history")
print("已清空 model_guidance_price_history, 影响行数:", c.rowcount)

# 5. 查看并删除指定VIN车辆
c.execute("SELECT id, vin, plate_number, car_type, guidance_price FROM vehicles WHERE vin='LFNA4LDC6TAE31340'")
v = c.fetchone()
if v:
    print("找到车辆:", dict(v))
    vehicle_id = v['id']
    # 删除关联数据
    for tbl in ['vehicle_guidance_price_history', 'sales_orders', 'contracts', 'repayments',
                'factory_repayments', 'attachments', 'contract_fee_items',
                'reconciliation_allocations', 'contract_initial_payments',
                'lock_requests', 'return_inspections', 'urge_records']:
        c.execute("DELETE FROM {} WHERE vehicle_id=?".format(tbl), (vehicle_id,))
        if c.rowcount:
            print("  清理表 {}: {} 行".format(tbl, c.rowcount))
    c.execute("DELETE FROM vehicles WHERE id=?", (vehicle_id,))
    print("已删除车辆, 影响行数:", c.rowcount)
else:
    print("未找到 VIN=LFNA4LDC6TAE31340 的车辆")

conn.commit()

# 6. 验证
c.execute("SELECT COUNT(*) FROM model_guidance_prices")
print("model_guidance_prices 剩余行数:", c.fetchone()[0])
c.execute("SELECT id, vin FROM vehicles WHERE vin='LFNA4LDC6TAE31340'")
print("车辆是否存在:", c.fetchone() is not None)

conn.close()
'''

# 写脚本到服务器
sftp = ssh.open_sftp()
with sftp.open('/tmp/clean_data.py', 'w') as f:
    f.write(clean_script)
sftp.close()

stdin, stdout, stderr = ssh.exec_command('/opt/jinjuyuan-venv/bin/python3 /tmp/clean_data.py')
print(stdout.read().decode())
err = stderr.read().decode().strip()
if err:
    print('STDERR:', err[:1000])

# 重启 gunicorn
print('Restarting gunicorn...')
ssh.exec_command('pkill -f gunicorn')
time.sleep(2)
ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'nohup gunicorn -b 0.0.0.0:8081 -w 1 --threads 8 --timeout 120 wsgi:application '
    '>/tmp/gunicorn.log 2>&1 &'
)
time.sleep(5)
stdin, stdout, stderr = ssh.exec_command('curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8081/')
print('HTTP:', stdout.read().decode().strip())

ssh.close()
print('Done!')
