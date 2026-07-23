import paramiko

host = '8.156.91.120'
password = 'fijSec-tuccyt-9dohxu'

ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(host, username='root', password=password, timeout=15)

# Write Python script to server via heredoc
stdin, stdout, stderr = ssh.exec_command(
    "cat > /tmp/clean.py << 'ENDPY'\n"
    "import sqlite3, os\n"
    "os.chdir('/opt/jinjuyuan')\n"
    "c = sqlite3.connect('jinjuyuan.db').cursor()\n"
    "r1 = c.execute('DELETE FROM model_guidance_prices').rowcount\n"
    "c.connection.commit()\n"
    "r2 = c.execute('DELETE FROM model_guidance_price_history').rowcount\n"
    "c.connection.commit()\n"
    "c.execute('SELECT COUNT(*) FROM model_guidance_prices')\n"
    "gp = c.fetchone()[0]\n"
    "c.execute('SELECT COUNT(*) FROM model_guidance_price_history')\n"
    "gh = c.fetchone()[0]\n"
    "c.execute('SELECT id, vin FROM vehicles WHERE vin=?')\n"
    "v = c.fetchone()\n"
    "print(f'DELETED guidance_prices: {r1}')\n"
    "print(f'DELETED guidance_history: {r2}')\n"
    "print(f'Remaining guidance_prices: {gp}')\n"
    "print(f'Remaining guidance_history: {gh}')\n"
    "if v:\n"
    "  vid = v[0]\n"
    "  for tbl in ['vehicle_guidance_price_history','sales_orders','contracts','repayments','factory_repayments']:\n"
    "    c.execute(f'DELETE FROM {tbl} WHERE vehicle_id=?', (vid,))\n"
    "  c.execute('DELETE FROM vehicles WHERE id=?', (vid,))\n"
    "  c.connection.commit()\n"
    "  print(f'DELETED vehicle id={vid} vin={v[1]}')\n"
    "else:\n"
    "  print('Vehicle already deleted')\n"
    "c.connection.close()\n"
    "ENDPY\n"
    "echo SCRIPT_OK"
)
out = stdout.read().decode()
print('Write:', out)

stdin, stdout, stderr = ssh.exec_command(
    'cd /opt/jinjuyuan && /opt/jinjuyuan-venv/bin/python3 /tmp/clean.py'
)
out = stdout.read().decode()
err = stderr.read().decode().strip()
print('OUTPUT:', out)
if err:
    print('ERR:', err)

ssh.close()
