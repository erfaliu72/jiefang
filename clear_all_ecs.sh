#!/bin/bash
# Self-contained clear script for ECS
# Save and run: bash clear_all_ecs.sh

APP_DIR="/opt/jinjuyuan"
VENV_DIR="/opt/jinjuyuan-venv"
RDS_HOST="rm-2vc15as5od6ay3gl4zo.mysql.cn-chengdu.rds.aliyuncs.com"
RDS_USER="dbroot"
RDS_PASS="fijSec-tuccyt-9dohxu"
RDS_DB="jinjuyuan"

echo "=== Step 1: Stop app ==="
pkill -f gunicorn 2>/dev/null
pkill -f wsgi:app 2>/dev/null
sleep 2

echo "=== Step 2: Clear database ==="
source "$VENV_DIR/bin/activate"
python3 << PYEOF
import pymysql
conn = pymysql.connect(host="$RDS_HOST", port=3306, user="$RDS_USER", password="$RDS_PASS", database="$RDS_DB", charset='utf8mb4')
c = conn.cursor()
c.execute('SET FOREIGN_KEY_CHECKS=0')
tables = ['reconciliation_allocations','late_fee_ledger','contract_fee_items','contract_initial_payments','factory_repayments','waivers','invoice_requests','return_inspections','ownership_transfers','vehicle_rebates','attachments','urge_records','lock_requests','approval_flows','daily_job_runs','customer_prepayments','receivables','customer_blacklist','customers','sales_orders','contracts','vehicles','audit_logs']
for t in tables:
    c.execute(f'SELECT COUNT(*) FROM {t}')
    cnt = c.fetchone()[0]
    if cnt:
        c.execute(f'DELETE FROM {t}')
        print(f'  OK {t}: {cnt}')
    else:
        print(f'  - {t}: empty')
c.execute('SET FOREIGN_KEY_CHECKS=1')
conn.commit()
for t in ['users', 'model_guidance_prices','receiving_companies']:
    c.execute(f'SELECT COUNT(*) FROM {t}')
    print(f'  KEPT {t}: {c.fetchone()[0]}')
conn.close()
print('CLEAR_DONE')
PYEOF

echo "=== Step 3: Clean uploads ==="
rm -rf "$APP_DIR/uploads/*" "$APP_DIR/generated-contracts/*" 2>/dev/null
echo "UPLOADS_CLEANED"

echo "=== Step 4: Restart app ==="
# Write env file proper format
cat > "$APP_DIR/.env" << 'ENVEOF'
JJY_DB_HOST=rm-2vc15as5od6ay3gl4zo.mysql.cn-chengdu.rds.aliyuncs.com
JJY_DB_PORT=3306
JJY_DB_USER=dbroot
JJY_DB_PASSWORD=fijSec-tuccyt-9dohxu
JJY_DB_NAME=jinjuyuan
JJY_SECRET_KEY=jinjuyuan-prod-2026-changeme
ENVEOF

systemctl restart jinjuyuan 2>/dev/null
RESTART_RESULT=$?
if [ $RESTART_RESULT -ne 0 ]; then
    cd "$APP_DIR"
    source "$VENV_DIR/bin/activate"
    export JJY_DB_HOST="$RDS_HOST"
    export JJY_DB_PORT=3306
    export JJY_DB_USER="$RDS_USER"
    export JJY_DB_PASSWORD="$RDS_PASS"
    export JJY_DB_NAME="$RDS_DB"
    nohup gunicorn -b 0.0.0.0:8081 -w 1 --threads 8 --timeout 120 wsgi:application >/tmp/jinjuyuan.log 2>&1 &
    echo "App started manually"
fi

sleep 5

echo "=== Step 5: Verify ==="
curl -s -o /dev/null -w "HTTP: %{http_code}\n" http://127.0.0.1:8081/
source "$VENV_DIR/bin/activate"
python3 << PYEOF
import pymysql
conn = pymysql.connect(host="$RDS_HOST", port=3306, user="$RDS_USER", password="$RDS_PASS", database="$RDS_DB", charset='utf8mb4')
c = conn.cursor()
for t in ['vehicles', 'contracts', 'sales_orders', 'audit_logs', 'users']:
    c.execute(f'SELECT COUNT(*) FROM {t}')
    print(f'  {t}: {c.fetchone()[0]}')
conn.close()
PYEOF

echo "=== ALL DONE ==="
echo "Visit: http://8.156.91.120:8081"
