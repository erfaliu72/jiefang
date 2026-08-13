#!/usr/bin/env python3
"""一次性数据迁移：车型指导价按基准车型（去厢型）合并（20260807 方案一）。

把 model_guidance_prices 中带厢型后缀的变体行（如 解放J6F全柴190LNG厢货/底盘/高栏）
合并到基准车型行（如 解放J6F全柴190LNG）。合并规则：金额列逐列取非0最大值。

- 自动备份 DB 到 jinjuyuan.db.bak.before_base_merge_<ts>
- 变体行删除前写入 merge_backup.json（可回滚）
- 幂等：重复执行全 no-op

用法: python3 migrate_guidance_merge_base.py
"""
import os
import json
import shutil
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'jinjuyuan.db')

# 从 app.py 复用归一化函数（不触发 Flask app 初始化）
import importlib.util
_spec = importlib.util.spec_from_file_location('app_module', os.path.join(BASE_DIR, 'app.py'))
# 注意：直接 import app 会初始化 Flask app。改为复制归一化逻辑以保证一致。
# 若 app.py import 安全也可直接 import；这里用独立实现 + 测试比对。


def strip_condition_prefix(car_type):
    if not car_type:
        return car_type or ''
    for prefix in ('新车', '二手车'):
        if car_type.startswith(prefix):
            return car_type[len(prefix):]
    return car_type


BOX_SUFFIX_WORDS = ('厢货', '宽体', '高栏', '冷藏', '平板', '底盘')
TAILGATE_SUFFIXES = (
    ('有尾板', '有尾板'),
    ('带尾板', '有尾板'),
    ('无尾板', '无尾板'),
    ('尾板', '有尾板'),
)


def strip_box_suffix(car_type):
    if not car_type:
        return car_type or ''
    s = str(car_type).strip()
    if ' / ' in s:
        parts = [p for p in s.split(' / ') if p]
        if len(parts) >= 2 and parts[-1] in BOX_SUFFIX_WORDS:
            return ' / '.join(parts[:-1])
        return s
    tailgate_suffix = ''
    for suffix, normalized in TAILGATE_SUFFIXES:
        if s.endswith(suffix) and len(s) > len(suffix):
            s = s[:-len(suffix)]
            tailgate_suffix = normalized
            break
    for w in BOX_SUFFIX_WORDS:
        if s.endswith(w) and len(s) > len(w):
            s = s[:-len(w)]
            break
    return s + tailgate_suffix


def normalize_base_car_type(car_type):
    base = strip_box_suffix(strip_condition_prefix(car_type))
    if not base:
        return ''
    if base.endswith(('有尾板', '无尾板')):
        return base
    return base + '无尾板'


# 参与合并的金额列
MONEY_COLS = [
    'guidance_price', 'lease_installment_price', 'sale_total_price',
    'lease_deposit_guidance', 'box_standard_price', 'box_wide_price',
    'box_high_rail_price', 'box_refrigerated_price', 'box_flatbed_price',
    'tail_plate_price', 'chassis_base_price', 'landing_price',
]
# 保留的行字段（文本类取非空最长）
TEXT_COLS = ['remark', 'product_code', 'fuel_type', 'interest_free_plan',
             'rent_to_buy_plan', 'min_loan_plan', 'lease_plan']


def main():
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup = f"{DB_PATH}.bak.before_base_merge_{ts}"
    shutil.copy2(DB_PATH, backup)
    print(f"[migrate] DB 已备份到: {backup}")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    c.execute("SELECT * FROM model_guidance_prices")
    rows = [dict(r) for r in c.fetchall()]
    print(f"[migrate] 迁移前 model_guidance_prices 行数: {len(rows)}")

    # 分组：base 归一化后 != car_type 的为变体行
    variants = []
    base_rows = {}
    for r in rows:
        base = normalize_base_car_type(r['car_type'])
        if base != r['car_type']:
            variants.append(r)
        else:
            base_rows.setdefault((base, r['is_new']), r)

    print(f"[migrate] 变体行: {len(variants)}, 基准行: {len(base_rows)}")
    if not variants:
        print("[migrate] 无变体行，无需迁移（幂等确认）")
        conn.close()
        return

    # 变体行备份
    backup_path = os.path.join(BASE_DIR, 'merge_backup.json')
    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(variants, f, ensure_ascii=False, indent=2, default=str)
    print(f"[migrate] 变体行备份到: {backup_path}")

    # 按 (base, is_new) 归组合并
    groups = {}
    for v in variants:
        base = normalize_base_car_type(v['car_type'])
        groups.setdefault((base, v['is_new']), []).append(v)

    merge_report = []
    for (base, is_new), vlist in groups.items():
        target = base_rows.get((base, is_new))
        merged = dict(target) if target else {
            'car_type': base, 'is_new': is_new,
            'guidance_price': 0, 'lease_installment_price': 0, 'sale_total_price': 0,
            'lease_deposit_guidance': 0,
            'box_standard_price': 0, 'box_wide_price': 0, 'box_high_rail_price': 0,
            'box_refrigerated_price': 0, 'box_flatbed_price': 0, 'tail_plate_price': 0,
            'chassis_base_price': 0, 'landing_price': 0,
            'remark': '', 'product_code': '', 'fuel_type': '',
            'interest_free_plan': '', 'rent_to_buy_plan': '', 'min_loan_plan': '', 'lease_plan': '',
        }
        for v in vlist:
            for col in MONEY_COLS:
                merged[col] = max(float(merged.get(col) or 0), float(v.get(col) or 0))
            for col in TEXT_COLS:
                a = (merged.get(col) or '').strip()
                b = (v.get(col) or '').strip()
                if b and len(b) > len(a):
                    merged[col] = b

        # 写库：UPDATE 或 INSERT
        if target:
            set_clause = ', '.join(f"{col}=?" for col in MONEY_COLS + TEXT_COLS)
            vals = [merged[col] for col in MONEY_COLS + TEXT_COLS]
            c.execute(f"UPDATE model_guidance_prices SET {set_clause} WHERE car_type=? AND is_new=?",
                      vals + [base, is_new])
            action = 'UPDATE'
        else:
            cols = ['car_type', 'is_new'] + MONEY_COLS + TEXT_COLS
            placeholders = ', '.join('?' for _ in cols)
            vals = [merged[col] for col in cols]
            c.execute(f"INSERT INTO model_guidance_prices ({', '.join(cols)}) VALUES ({placeholders})", vals)
            action = 'INSERT'

        merge_report.append({
            'action': action, 'base': base, 'is_new': is_new,
            'from': [v['car_type'] for v in vlist],
            'merged': {col: merged[col] for col in MONEY_COLS if merged.get(col)},
        })

    # 删除变体行
    variant_ids = [v['id'] for v in variants]
    c.execute(f"DELETE FROM model_guidance_prices WHERE id IN ({','.join('?' for _ in variant_ids)})", variant_ids)

    conn.commit()

    # 打印报告
    print("\n[merge] 合并结果:")
    for rep in merge_report:
        print(f"  [{rep['action']}] {rep['base']} ({rep['is_new']})")
        print(f"      ← {', '.join(rep['from'])}")
        if rep['merged']:
            print(f"      值: {rep['merged']}")

    c.execute("SELECT COUNT(*) FROM model_guidance_prices")
    print(f"\n[migrate] 迁移后行数: {c.fetchone()[0]}")
    conn.close()
    print("[migrate] 完成 ✓")


if __name__ == '__main__':
    main()
