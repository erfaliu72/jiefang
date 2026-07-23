"""通过SFTP上传脚本到服务器，清理指导价数据和指定VIN车辆"""
import paramiko, time

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

sftp = ssh.open_sftp()

# 写清理脚本
clear_script = (
    'import os, sqlite3\n'
    'os.chdir("/opt/jinjuyuan")\n'
    'c = sqlite3.connect("jinjuyuan.db").cursor()\n\n'
    '# 查看当前数据\n'
    'c.execute("SELECT car_type, guidance_price FROM model_guidance_prices")\n'
    'rows = c.fetchall()\n'
    'print(f"当前指导价 {len(rows)} 条:")\n'
    'for r in rows:\n'
    '    print(f"  {r[0]} -> {r[1]}")\n\n'
    '# 清除\n'
    'c.execute("DELETE FROM model_guidance_prices")\n'
    'print(f"已删除 model_guidance_prices: {c.rowcount} 条")\n'
    'c.execute("DELETE FROM model_guidance_price_history")\n'
    'print(f"已删除 model_guidance_price_history: {c.rowcount} 条")\n\n'
    '# 删除车辆\n'
    'c.execute("SELECT id FROM vehicles WHERE vin=\'LFNA4LDC6TAE31340\'")\n'
    'v = c.fetchone()\n'
    'if v:\n'
    '    vid = v[0]\n'
    '    for tbl in ["vehicle_guidance_price_history","sales_orders","contracts","repayments","factory_repayments"]:\n'
    '        c.execute(f"DELETE FROM {tbl} WHERE vehicle_id={vid}")\n'
    '    c.execute(f"DELETE FROM vehicles WHERE id={vid}")\n'
    '    print(f"已删除车辆 id={vid}")\n'
    'else:\n'
    '    print("车辆 LFNA4LDC6TAE31340 不存在或已删除")\n\n'
    'c.connection.commit()\n\n'
    '# 验证\n'
    'c.execute("SELECT COUNT(*) FROM model_guidance_prices")\n'
    'c2 = sqlite3.connect("jinjuyuan.db").cursor()\n'
    'c2.execute("SELECT COUNT(*) FROM model_guidance_prices")\n'
    'print(f"操作后 model_guidance_prices: {c2.fetchone()[0]} 条")\n'
    'c2.execute("SELECT COUNT(*) FROM vehicles WHERE vin=\'LFNA4LDC6TAE31340\'")\n'
    'print(f"车辆是否存在: {c2.fetchone()[0]}")\n'
    'c2.connection.close()\n'
)

with sftp.open('/tmp/clear_guidance.py', 'w') as f:
    f.write(clear_script)
sftp.close()
print('✅ 脚本已上传')

# 执行清理
stdin, stdout, stderr = ssh.exec_command('/opt/jinjuyuan-venv/bin/python3 /tmp/clear_guidance.py')
out = stdout.read().decode()
err = stderr.read().decode().strip()
print('输出:', out)
if err:
    print('错误:', err[:300])

# 重启
print('重启服务...')
time.sleep(1)
ssh.exec_command('pkill -f gunicorn')
time.sleep(3)
ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'nohup gunicorn -b 0.0.0.0:8081 -w 1 --threads 8 --timeout 120 wsgi:application '
    '>/tmp/gunicorn.log 2>&1 &'
)
time.sleep(5)

# 验证API
import json
stdin, stdout, stderr = ssh.exec_command(
    'curl -s -c /tmp/boss_cookie -X POST -H "Content-Type: application/json" '
    '-d \'{"username":"boss","password":"123456"}\' '
    'http://127.0.0.1:8081/api/auth/login > /dev/null && '
    'curl -s -b /tmp/boss_cookie http://127.0.0.1:8081/api/model-guidance-prices'
)
try:
    data = json.loads(stdout.read().decode())
    if isinstance(data, list):
        print(f'\nAPI 返回 {len(data)} 条')
        for p in data:
            print(f'  {p.get("car_type")} -> {p.get("guidance_price")}')
    else:
        print('API响应:', str(data)[:300])
except Exception as e:
    print(f'解析失败: {e}')

ssh.close()
print('✅ 完成')
