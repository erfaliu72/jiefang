"""部署脚本：上传文件 → 清缓存 → systemctl restart（保留远端现有数据）"""
import paramiko, time, json

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

sftp = ssh.open_sftp()

# 上传文件
for local_name, remote_path in [
    ('app.py', '/opt/jinjuyuan/app.py'),
    ('database.py', '/opt/jinjuyuan/database.py'),
    ('wsgi.py', '/opt/jinjuyuan/wsgi.py'),
    ('templates/index.html', '/opt/jinjuyuan/templates/index.html'),
]:
    sftp.put(f'/Users/liuyuanchang/code/jiefang/{local_name}', remote_path)
    print(f'✅ {local_name}')
sftp.close()

# 停止
ssh.exec_command('pkill -9 -f gunicorn 2>/dev/null')
time.sleep(2)

# 清缓存
ssh.exec_command(
    'find /opt/jinjuyuan -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null; '
    'find /opt/jinjuyuan -name "*.pyc" -delete 2>/dev/null; '
    'find /opt/jinjuyuan -name "*.py" -exec touch {} + 2>/dev/null'
)
print('🧹 Cache cleared')

# 先停系统服务
time.sleep(1)
ssh.exec_command('pkill -9 -f gunicorn 2>/dev/null')
time.sleep(2)

# systemctl 启动
time.sleep(2)
stdin, stdout, stderr = ssh.exec_command('systemctl daemon-reload && systemctl restart jinjuyuan && sleep 8 && systemctl is-active jinjuyuan')
print(f'  systemctl: {stdout.read().decode().strip()}')

time.sleep(3)
stdin, stdout, stderr = ssh.exec_command('curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8081/')
print(f'✅ HTTP: {stdout.read().decode().strip()}')


# 验证服务正常
stdin, stdout, stderr = ssh.exec_command('curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8081/api/vehicles?page=1\\&page_size=1')
print(f'📊 车辆API: {stdout.read().decode().strip()}')
print('✅ Done')
