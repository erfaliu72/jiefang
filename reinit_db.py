import paramiko, time

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

print('1. Stopping gunicorn...')
ssh.exec_command('pkill -f gunicorn')
time.sleep(2)

print('2. Removing empty db file...')
stdin, stdout, stderr = ssh.exec_command('rm -f /opt/jinjuyuan/jinjuyuan.db && echo "OK"')
print('  Result:', stdout.read().decode().strip())

print('3. Initializing database...')
stdin, stdout, stderr = ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'python3 /dev/stdin << "PYEOF"\n'
    'from database import init_db\n'
    'init_db()\n'
    'print("init_db OK")\n'
    'PYEOF'
)
print('  ', stdout.read().decode().strip())
err = stderr.read().decode().strip()
if err: print('  ERR:', err)

print('4. Seeding data...')
stdin, stdout, stderr = ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'python3 /dev/stdin << "PYEOF"\n'
    'from database import seed_data\n'
    'seed_data()\n'
    'print("seed_data OK")\n'
    'PYEOF'
)
print('  ', stdout.read().decode().strip())
err = stderr.read().decode().strip()
if err: print('  ERR:', err)

print('5. Verifying...')
stdin, stdout, stderr = ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'python3 /dev/stdin << "PYEOF"\n'
    'from database import get_db\n'
    'conn = get_db()\n'
    'c = conn.cursor()\n'
    'c.execute("SELECT id, username, role, display_name FROM users")\n'
    'for u in c.fetchall():\n'
    '    print(dict(u))\n'
    'c.execute("SELECT COUNT(*) FROM vehicles")\n'
    'print("Vehicles:", c.fetchone()[0])\n'
    'conn.close()\n'
    'PYEOF'
)
print(stdout.read().decode())
err = stderr.read().decode().strip()
if err: print('  ERR:', err)

print('6. Starting gunicorn...')
time.sleep(2)
ssh.exec_command(
    'cd /opt/jinjuyuan && source /opt/jinjuyuan-venv/bin/activate && '
    'nohup gunicorn -b 0.0.0.0:8081 -w 1 --threads 8 --timeout 120 wsgi:application '
    '>/tmp/gunicorn.log 2>&1 &'
)
time.sleep(5)

stdin, stdout, stderr = ssh.exec_command('curl -s -o /dev/null -w "%{http_code}" http://127.0.0.1:8081/')
print('  HTTP status:', stdout.read().decode().strip())

ssh.close()
print('Done!')
