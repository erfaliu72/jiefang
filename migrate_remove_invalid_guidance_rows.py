#!/usr/bin/env python3
"""删除无法校验、且无业务引用的数字开头历史指导价。

执行前备份数据库到 /private/tmp；价格历史和审计记录保留。
默认只预览，传入 --apply 后才实际删除。
"""

import argparse
import shutil
import sqlite3
from datetime import datetime

from database import DATABASE


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--apply", action="store_true", help="执行删除，默认只预览")
    args = parser.parse_args()

    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            """
            SELECT p.id, p.car_type, p.is_new
            FROM model_guidance_prices p
            WHERE p.car_type GLOB '[0-9]*'
              AND NOT EXISTS (
                  SELECT 1 FROM vehicles v
                  WHERE (v.is_deleted IS NULL OR v.is_deleted=0)
                    AND v.car_type=p.car_type
              )
              AND NOT EXISTS (
                  SELECT 1 FROM sales_orders o WHERE o.car_type=p.car_type
              )
              AND NOT EXISTS (
                  SELECT 1 FROM finance_plans f WHERE f.car_type=p.car_type
              )
            ORDER BY p.id
            """
        ).fetchall()
        print(f"待删除无效历史指导价：{len(rows)} 条")
        for row in rows:
            print(f"  #{row['id']} {row['car_type']} / {row['is_new'] or '新车'}")
        if not args.apply:
            print("预览完成；使用 --apply 执行删除。")
            return

        stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        backup_path = f"/private/tmp/jinjuyuan.db.before_invalid_guidance_cleanup_{stamp}.bak"
        shutil.copy2(DATABASE, backup_path)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for row in rows:
            conn.execute(
                """
                INSERT INTO audit_logs (action, target_type, target_id, detail, operator, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    "清理非标准车型指导价",
                    "model_guidance_price",
                    row["id"],
                    f"删除无库存、订单和方案引用的非标准指导价：{row['car_type']}（{row['is_new'] or '新车'}）",
                    "系统",
                    now,
                ),
            )
        conn.executemany("DELETE FROM model_guidance_prices WHERE id=?", [(row["id"],) for row in rows])
        conn.commit()
        print(f"清理完成，数据库备份：{backup_path}")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
