#!/usr/bin/env python3
"""Connect to RDS and verify data"""
import os, sys
os.environ['JJY_DB_HOST'] = 'rm-2vc15as5od6ay3gl4zo.mysql.cn-chengdu.rds.aliyuncs.com'
os.environ['JJY_DB_PORT'] = '3306'
os.environ['JJY_DB_USER'] = 'dbroot'
os.environ['JJY_DB_PASSWORD'] = 'fijSec-tuccyt-9dohxu'
os.environ['JJY_DB_NAME'] = 'jinjuyuan'

# Must change to app dir for database import
sys.path.insert(0, '/opt/jinjuyuan')
os.chdir('/opt/jinjuyuan')

from database import get_db, USE_MYSQL
print('MYSQL mode:', USE_MYSQL)
conn = get_db()
c = conn.cursor()
c.execute('SELECT COUNT(*) FROM users')
print('users:', c.fetchone()[0])
c.execute('SELECT COUNT(*) FROM vehicles')
print('vehicles:', c.fetchone()[0])
c.execute('SELECT COUNT(*) FROM contracts')
print('contracts:', c.fetchone()[0])
c.execute('SELECT COUNT(*) FROM audit_logs')
print('audit_logs:', c.fetchone()[0])
conn.close()
print('VERIFY_OK')
