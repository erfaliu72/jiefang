#!/usr/bin/env python3
"""SQLite -> MySQL 数据迁移。
在 ECS 上运行：先用改造后的 database.init_db() 在 RDS 建好表结构，
再把本地带过来的 jinjuyuan.db 所有业务数据搬到 RDS。
依赖环境变量 JJY_DB_HOST/PORT/USER/PASSWORD/NAME。
"""
import os
import sys
import sqlite3
import pymysql

SQLITE_PATH = sys.argv[1] if len(sys.argv) > 1 else 'jinjuyuan.db'

MYSQL_CONF = {
    'host': os.environ['JJY_DB_HOST'],
    'port': int(os.environ.get('JJY_DB_PORT', '3306')),
    'user': os.environ.get('JJY_DB_USER', 'root'),
    'password': os.environ.get('JJY_DB_PASSWORD', ''),
    'database': os.environ.get('JJY_DB_NAME', 'jinjuyuan'),
    'charset': 'utf8mb4',
}


def main():
    src = sqlite3.connect(SQLITE_PATH)
    src.row_factory = sqlite3.Row
    sc = src.cursor()

    # 取所有业务表（排除 sqlite 内部表）
    sc.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
    tables = [r[0] for r in sc.fetchall()]
    print(f"待迁移表: {len(tables)} 张")

    dst = pymysql.connect(**MYSQL_CONF)
    dc = dst.cursor()
    dc.execute("SET FOREIGN_KEY_CHECKS=0")
    dc.execute("SET SESSION sql_mode=''")

    total_rows = 0
    for t in tables:
        # 目标表必须已由 init_db 建好；取目标表实际列
        try:
            dc.execute(f"SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=%s AND TABLE_NAME=%s",
                       (MYSQL_CONF['database'], t))
            mysql_cols = [r[0] for r in dc.fetchall()]
        except Exception as e:
            print(f"  [跳过] {t}: 目标表不存在 ({e})")
            continue
        if not mysql_cols:
            print(f"  [跳过] {t}: 目标无此表")
            continue

        sc.execute(f"SELECT * FROM {t}")
        rows = sc.fetchall()
        if not rows:
            print(f"  {t}: 0 行")
            continue

        src_cols = [d[0] for d in sc.description]
        # 只迁移两边都有的列
        use_cols = [c for c in src_cols if c in mysql_cols]
        col_list = ','.join(f'`{c}`' for c in use_cols)
        placeholders = ','.join(['%s'] * len(use_cols))
        # 先清空目标表，保证可重跑
        dc.execute(f"DELETE FROM `{t}`")
        sql = f"INSERT INTO `{t}` ({col_list}) VALUES ({placeholders})"
        data = [[row[c] for c in use_cols] for row in rows]
        dc.executemany(sql, data)
        total_rows += len(data)
        print(f"  {t}: {len(data)} 行")

    dc.execute("SET FOREIGN_KEY_CHECKS=1")
    dst.commit()
    dst.close()
    src.close()
    print(f"迁移完成，共 {total_rows} 行")


if __name__ == '__main__':
    main()
