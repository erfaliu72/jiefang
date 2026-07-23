"""验证远端服务状态"""
import paramiko, json

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

# 1. 登录
stdin, stdout, stderr = ssh.exec_command(
    'curl -s -c /tmp/boss_cookie -X POST -H "Content-Type: application/json" '
    '-d \'{"username":"boss","password":"123456"}\' '
    'http://127.0.0.1:8081/api/auth/login'
)
login = json.loads(stdout.read().decode())
print(f'Login: {"OK" if login.get("success") else "FAIL"}')

# 2. Dashboard
stdin, stdout, stderr = ssh.exec_command(
    'curl -s -b /tmp/boss_cookie http://127.0.0.1:8081/api/dashboard/stats'
)
try:
    d = json.loads(stdout.read().decode())
    print(f'Dashboard: total_vehicles={d.get("total_vehicles")}, in_stock={d.get("in_stock")}')
except:
    print('Dashboard: 500')

# 3. Guidance prices
stdin, stdout, stderr = ssh.exec_command(
    'curl -s -b /tmp/boss_cookie http://127.0.0.1:8081/api/model-guidance-prices'
)
try:
    d = json.loads(stdout.read().decode())
    print(f'Guidance prices: {len(d)} items')
    for p in d:
        print(f'  {p["car_type"]} -> {p["guidance_price"]}')
except Exception as e:
    print(f'Guidance error: {e}')

# 4. Vehicles list
stdin, stdout, stderr = ssh.exec_command(
    'curl -s -b /tmp/boss_cookie "http://127.0.0.1:8081/api/vehicles?page=1&page_size=1"'
)
try:
    d = json.loads(stdout.read().decode())
    count = len(d) if isinstance(d, list) else (d.get("total") if isinstance(d, dict) else "?")
    print(f'Vehicles: {count}')
except:
    print('Vehicles: error')

ssh.close()
