from flask import Flask, jsonify, request, send_from_directory, session
from flask_cors import CORS
from database import get_db, init_db, seed_data
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from functools import wraps
from werkzeug.utils import secure_filename
from openpyxl import load_workbook
import os, uuid
import json
import re
import threading, time as _time

app = Flask(__name__)
app.secret_key = os.environ.get('JJY_SECRET_KEY', 'jinjuyuan-secret-2024')
CORS(app, supports_credentials=True)

# 上传文件目录
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)
CONTRACT_OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'generated-contracts')
os.makedirs(CONTRACT_OUTPUT_DIR, exist_ok=True)

_SCHEDULER_STARTED = False

# ======================== PRD 角色权限矩阵 ========================
# 每个角色可访问的页面 — 车辆列表全员可见
ROLE_PAGES = {
    '老板': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'bills', 'reconciliation', 'risk', 'profit', 'settings'],
    '运营': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'reconciliation', 'risk', 'invoice'],
    '财务': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'bills', 'reconciliation', 'profit'],
    '车管': ['dashboard', 'assets', 'completion_history', 'approvals'],
    '销售': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'risk'],
}

# 每个角色可执行的操作
ROLE_ACTIONS = {
    '老板': ['*'],
    '运营': ['view_contracts', 'view_overdue', 'lock_vehicle', 'execute_lock', 'confirm_repayment', 'initiate_return', 'view_orders'],
    '财务': ['view_contracts', 'confirm_repayment', 'confirm_factory', 'view_bills', 'view_profit', 'upload_receipt', 'collect_payment', 'verify_return', 'upload_initial_receipt', 'activate_order'],
    '车管': ['add_vehicle', 'update_vehicle', 'return_inspect', 'deliver_vehicle'],
    '销售': ['create_contract', 'view_contracts', 'upload_screenshot', 'view_overdue', 'initiate_return', 'request_lock', 'initiate_initial_payment', 'create_order'],
}

# AA-AR 财务敏感字段（对所有非财务/非老板角色隐藏）——20260805 字段精简后已无残留
_AA_AR_FIELDS = [
    'purchase_price', 'tax_rate',
]

ROLE_HIDDEN_FIELDS = {
    '销售': {
        'vehicles': ['purchase_price', 'tax_rate', 'estimated_residual_value'] + _AA_AR_FIELDS,
        'contracts': [
            'loan_amount', 'monthly_payment', 'factory_guarantee_deposit', 'paid_principal',
            'loan_balance', 'collected_deposit', 'collected_rent', 'expected_profit_floor',
            'expected_profit_ceiling', 'snapshot_guidance_price', 'snapshot_invoice_price',
        ],
        'factory_repayments': ['amount'],
        'vehicle_rebates': ['*'],
        'profit': ['*'],
        'customer_blacklist': ['*'],
    },
    '运营': {
        'vehicles': ['purchase_price', 'tax_rate'] + _AA_AR_FIELDS,
        'contracts': ['snapshot_guidance_price', 'snapshot_invoice_price'],
        'factory_repayments': ['amount'],
        'vehicle_rebates': ['*'],
        'profit': ['*'],
        'customer_blacklist': ['*'],
    },
    '车管': {
        'vehicles': [
            'purchase_price', 'tax_rate', 'estimated_residual_value', 'paid_principal',
            'loan_balance', 'collected_deposit', 'collected_rent',
        ] + _AA_AR_FIELDS,
        'contracts': [
            'loan_amount', 'monthly_payment', 'factory_guarantee_deposit', 'paid_principal',
            'loan_balance', 'collected_deposit', 'collected_rent', 'deposit', 'down_payment',
            'expected_profit_floor', 'expected_profit_ceiling', 'snapshot_guidance_price',
            'snapshot_invoice_price',
        ],
        'sales_orders': ['snapshot_guidance_price', 'snapshot_lease_installment_price'],
        'factory_repayments': ['amount'],
        'vehicle_rebates': ['*'],
        'profit': ['*'],
        'customer_blacklist': ['*'],
    },
}

FEE_PRIORITY_ORDER = {
    'insurance_fee': 1,
    'service_fee': 1,
    'maintenance_fee': 2,
    'accident_fee': 2,
    'penalty_fee': 2,
    'late_fee': 3,
    'rent': 4,
}

PROFIT_CEILING_SENTINEL = 999999999

# ======================== 通用审批流程配置 ========================
APPROVAL_CONFIGS = {
    'price_exception': [
        {'step': 1, 'role': '老板', 'label': '价格审批'},
    ],
    'order_exception': [
        {'step': 1, 'role': '老板', 'label': '报单异常审批'},
    ],
    'sale_payment': [
        {'step': 1, 'role': '财务', 'label': '财务确认报单'},
    ],
    'initial_payment': [
        {'step': 1, 'role': '财务', 'label': '财务审核收款'},
    ],
    'initial_payment_shortage': [
        {'step': 1, 'role': '老板', 'label': '首次付款不足出库审批'},
    ],
    'return_stock': [
        {'step': 1, 'role': '老板', 'label': '领导审批'},
    ],
    'invoice': [
        {'step': 1, 'role': '老板', 'label': '发票审批'},
    ],
}

# ======================== 20260804 大改版常量 ========================
# 租赁月供指导价：箱型 5 选 1 → model_guidance_prices 列名
BOX_TYPE_MONTHLY_PRICE_COLUMNS = {
    '厢货': 'box_standard_price',
    '宽体': 'box_wide_price',
    '高栏': 'box_high_rail_price',
    '冷藏': 'box_refrigerated_price',
    '平板': 'box_flatbed_price',
}
GUIDANCE_BOX_TYPES = tuple(BOX_TYPE_MONTHLY_PRICE_COLUMNS.keys())
TAIL_PLATE_SURCHARGE = 300  # 尾板加价默认值（可用 model_guidance_prices.tail_plate_price 覆盖）


def strip_condition_prefix(car_type):
    """去掉车型名的成色前缀（新车/二手车），返回基准车型名。"""
    if not car_type:
        return car_type or ''
    for prefix in ('新车', '二手车'):
        if car_type.startswith(prefix):
            return car_type[len(prefix):]
    return car_type


# 厢型后缀词表。尾板属于基准车型，不在这里剥离。
BOX_SUFFIX_WORDS = ('厢货', '宽体', '高栏', '冷藏', '平板', '底盘')
TAILGATE_SUFFIXES = (
    ('有尾板', '有尾板'),
    ('带尾板', '有尾板'),
    ('无尾板', '无尾板'),
    ('尾板', '有尾板'),
)


def strip_box_suffix(car_type):
    """去掉车型名末尾厢型，保留并统一尾板标识。

    基准车型的尾板是身份维度：'冷藏带尾板' 归一为 '有尾板'，而不是被删掉。
    """
    if not car_type:
        return car_type or ''
    s = str(car_type).strip()
    # 1) 拼接格式：'XX / YY / 厢货' → 'XX / YY'
    if ' / ' in s:
        parts = [p for p in s.split(' / ') if p]
        if len(parts) >= 2 and parts[-1] in BOX_SUFFIX_WORDS:
            return ' / '.join(parts[:-1])
        return s
    # 2) 紧凑格式：先取尾板，再去除厢型，最后恢复尾板。
    tailgate_suffix = ''
    for suffix, normalized in TAILGATE_SUFFIXES:
        if s.endswith(suffix) and len(s) > len(suffix):
            s = s[:-len(suffix)]
            tailgate_suffix = normalized
            break
    # 兼容旧前端曾把“无尾板/有尾板”截成单个“无/有”后又回传的脏值。
    if tailgate_suffix == '无尾板' and s.endswith('无'):
        s = s[:-1]
    elif tailgate_suffix == '有尾板' and s.endswith('有'):
        s = s[:-1]
    for w in BOX_SUFFIX_WORDS:
        if s.endswith(w) and len(s) > len(w):
            s = s[:-len(w)]
            break
    return s + tailgate_suffix


def normalize_base_car_type(car_type):
    """组合归一化：去成色、去厢型，尾板统一为基准车型固定维度。

    历史车型未显式保存尾板时按“无尾板”兼容，避免旧库存和指导价失配。
    """
    base = strip_box_suffix(strip_condition_prefix(car_type))
    if not base:
        return ''
    if base.endswith('带尾板'):
        return base[:-len('带尾板')] + '有尾板'
    if base.endswith('有尾板') or base.endswith('无尾板'):
        return base
    # 兼容旧前端将尾板后缀剥离后留下的单字尾标识。
    if base.endswith('有'):
        return base[:-1] + '有尾板'
    if base.endswith('无'):
        return base[:-1] + '无尾板'
    return base + '无尾板'


def normalize_vehicle_condition(value):
    """将车辆成色统一为定价维度值。"""
    return value if value in ('新车', '二手车') else '新车'


def extract_guidance_box_type(car_type):
    """从历史拼接车型中提取指导价厢型，供旧接口兼容和方案范围校验使用。"""
    raw = strip_condition_prefix(car_type or '')
    if ' / ' in raw:
        parts = [part.strip() for part in raw.split(' / ') if part.strip()]
        if parts and parts[-1] in GUIDANCE_BOX_TYPES:
            return parts[-1]
    raw = raw.strip()
    for suffix in ('带尾板', '尾板'):
        if raw.endswith(suffix) and len(raw) > len(suffix):
            raw = raw[:-len(suffix)]
            break
    for box_type in GUIDANCE_BOX_TYPES:
        if raw.endswith(box_type) and len(raw) > len(box_type):
            return box_type
    return ''


def normalize_tailgate(value):
    """将采购模板尾板值统一为 SKU/报价使用的“有/无”。"""
    raw = str(value or '').strip()
    if not raw:
        return '无'
    if raw in ('无', '不带尾板', '无尾板', '否'):
        return '无'
    return '有' if raw in ('有', '带尾板', '有尾板', '是') else '无'


def energy_type_for_vehicle(vehicle):
    """按燃料形式识别车型能源类型；旧数据缺燃料形式时用规格作保守推断。"""
    vehicle = vehicle or {}
    fuel = str(vehicle.get('fuel_form') or '').strip()
    if '纯电' in fuel:
        return '纯电'
    if '混动' in fuel:
        return '混动'
    if fuel:
        return '燃油车'
    if (vehicle.get('battery_brand') or vehicle.get('battery_capacity') or vehicle.get('battery_model')):
        return '纯电'
    if (vehicle.get('horsepower') or vehicle.get('gear_position')):
        return '燃油车'
    return ''


BASE_MODEL_FIELDS_BY_ENERGY = {
    '纯电': ('brand', 'product_series', 'battery_brand', 'battery_capacity'),
    '混动': ('brand', 'product_series', 'horsepower', 'gear_position'),
    '燃油车': ('brand', 'product_series', 'horsepower', 'gear_position'),
}
ENERGY_EXCLUSIVE_FIELDS = {
    '纯电': ('horsepower', 'gear_position'),
    '混动': ('battery_brand', 'battery_capacity'),
    '燃油车': ('battery_brand', 'battery_capacity'),
}


def base_model_validation_errors(vehicle, require_identity=False):
    """校验基准车型能源口径；返回可直接给接口展示的错误信息。"""
    vehicle = vehicle or {}
    has_identity_input = any(
        str(vehicle.get(field) or '').strip()
        for field in (
            'fuel_form', 'brand', 'product_series', 'battery_brand', 'battery_capacity',
            'horsepower', 'gear_position', 'tailgate',
        )
    )
    if not has_identity_input and not require_identity:
        return []
    energy_type = energy_type_for_vehicle(vehicle)
    if not energy_type:
        return ['请填写燃料形式，或补全可识别的车型规格']
    missing = [
        field for field in BASE_MODEL_FIELDS_BY_ENERGY[energy_type]
        if not str(vehicle.get(field) or '').strip()
    ]
    if missing:
        labels = {
            'brand': '品牌', 'product_series': '品系', 'battery_brand': '电池品牌',
            'battery_capacity': '电池度数', 'horsepower': '马力', 'gear_position': '档位',
        }
        return [f"{energy_type}基准车型缺少“{'、'.join(labels[field] for field in missing)}”"]
    invalid_fields = [
        field for field in ENERGY_EXCLUSIVE_FIELDS[energy_type]
        if str(vehicle.get(field) or '').strip()
    ]
    if invalid_fields:
        labels = {
            'battery_brand': '电池品牌', 'battery_capacity': '电池度数',
            'horsepower': '马力', 'gear_position': '档位',
        }
        return [f"{energy_type}车型不应填写“{'、'.join(labels[field] for field in invalid_fields)}”"]
    return []


def base_model_tailgate_label(value):
    return '有尾板' if normalize_tailgate(value) == '有' else '无尾板'


def has_base_model_input(vehicle):
    """是否已开始填写基准车型字段，用于区分旧车型兼容数据与新建规格数据。"""
    vehicle = vehicle or {}
    return any(
        str(vehicle.get(field) or '').strip()
        for field in (
            'fuel_form', 'brand', 'product_series', 'battery_brand',
            'battery_capacity', 'horsepower', 'gear_position', 'tailgate',
        )
    )


def normalize_base_model_tailgate(vehicle):
    """新建或编辑的规格车统一落库尾板“有/无”；旧数据不做批量改写。"""
    if has_base_model_input(vehicle):
        vehicle['tailgate'] = normalize_tailgate(vehicle.get('tailgate'))
    return vehicle


def validate_manual_vehicle_base_model(vehicle):
    """手工新增/编辑时阻止能源字段串用，导入保留为可追踪的校验状态。"""
    if not has_base_model_input(vehicle):
        return []
    return base_model_validation_errors(vehicle, require_identity=True)


def guidance_base_candidates(car_type):
    """指导价查询兼容旧的未显式尾板基准车型。"""
    base = normalize_base_car_type(car_type)
    candidates = [base]
    if base.endswith('无尾板'):
        legacy = base[:-len('无尾板')]
        if legacy:
            candidates.append(legacy)
    return candidates


def sales_order_energy_category(fuel_form):
    """把采购模板燃料形式映射为报单界面的能源分类。"""
    return energy_type_for_vehicle({'fuel_form': fuel_form})


def vehicle_order_snapshot(vehicle):
    """从库存 VIN 生成报单车辆快照，报价相关字段不能由报单表单覆盖。"""
    if not vehicle:
        return {}
    vehicle = dict(vehicle)
    return {
        'is_new': (vehicle.get('condition') or '新车').strip() or '新车',
        'vehicle_brand': (vehicle.get('brand') or '').strip(),
        'vehicle_category': sales_order_energy_category(vehicle.get('fuel_form')),
        'vehicle_cab': (vehicle.get('cab_type') or vehicle.get('cab_style') or '').strip(),
        'vehicle_engine_battery': (
            vehicle.get('battery_brand') or vehicle.get('engine_spec') or ''
        ).strip(),
        'vehicle_power_battery': (
            vehicle.get('battery_capacity') or vehicle.get('horsepower') or ''
        ).strip(),
        'vehicle_gearbox': (
            vehicle.get('gear_position') or vehicle.get('gearbox_spec') or ''
        ).strip(),
        'vehicle_box_type': (
            vehicle.get('box_type') or vehicle.get('vehicle_box_type') or ''
        ).strip(),
        'vehicle_color': (vehicle.get('vehicle_color') or '').strip(),
        'tail_plate': normalize_tailgate(vehicle.get('tailgate')),
        'car_type': (vehicle.get('car_type') or '').strip(),
        'plate_number': (vehicle.get('plate_number') or '').strip(),
        'car_purchase_amount': parse_money(vehicle.get('purchase_price')),
    }


def resolve_lease_guidance(conn, car_type, vehicle_box_type=None, tail_plate=None, is_new=None):
    """读取车型租赁指导价：押金指导价 + 箱型月供指导价（含尾板加价）。
    返回 {deposit_guidance, monthly_guidance, box_type, tail_plate_surcharge}
    is_new: '新车'/'二手车'，按成色取对应行；None 时取任意一行（兼容老数据）。
    """
    result = {'deposit_guidance': 0.0, 'monthly_guidance': 0.0, 'box_type': vehicle_box_type or '', 'tail_plate_surcharge': 0.0}
    if not car_type:
        return result
    candidates = guidance_base_candidates(car_type)
    sql = """SELECT lease_deposit_guidance,
               box_standard_price, box_wide_price, box_high_rail_price,
               box_refrigerated_price, box_flatbed_price, tail_plate_price
        FROM model_guidance_prices
        WHERE car_type IN ({})""".format(', '.join('?' for _ in candidates))
    params = list(candidates)
    if is_new:
        sql += " AND is_new=?"
        params.append(is_new)
    sql += " ORDER BY CASE car_type WHEN ? THEN 0 ELSE 1 END LIMIT 1"
    params.append(candidates[0])
    row = conn.execute(sql, params).fetchone()
    if not row:
        return result
    deposit_guidance = parse_money(row['lease_deposit_guidance'])
    box_col = BOX_TYPE_MONTHLY_PRICE_COLUMNS.get(vehicle_box_type or '')
    box_price = parse_money(row[box_col]) if box_col and box_col in row.keys() else 0.0
    tail_surcharge = parse_money(row['tail_plate_price']) if parse_money(row['tail_plate_price']) > 0 else TAIL_PLATE_SURCHARGE
    tail_surcharge = tail_surcharge if normalize_tailgate(tail_plate) == '有' else 0.0
    result.update({
        'deposit_guidance': deposit_guidance,
        'monthly_guidance': round(box_price + tail_surcharge, 2),
        'tail_plate_surcharge': tail_surcharge,
    })
    return result


def resolve_finance_plan(conn, plan_id, car_type=None, condition=None, box_type=None):
    """读取以租代售金融方案；可校验其是否属于指定新车子型号。
    不存在、已停用或不在生效日期内返回 None。"""
    if not plan_id:
        return None
    row = conn.execute(
        """SELECT * FROM finance_plans WHERE id=? AND status='启用'
           AND (effective_date IS NULL OR effective_date = '' OR effective_date <= date('now'))
           AND (expiry_date IS NULL OR expiry_date = '' OR expiry_date >= date('now'))""",
        (plan_id,)
    ).fetchone()
    if not row:
        return None
    plan = dict(row)
    if car_type and normalize_base_car_type(plan['car_type']) != normalize_base_car_type(car_type):
        return None
    if condition and normalize_vehicle_condition(plan.get('condition')) != normalize_vehicle_condition(condition):
        return None
    if box_type and (plan.get('box_type') or '').strip() != box_type.strip():
        return None
    return plan


def has_active_finance_plan_for_vehicle(conn, car_type, condition, box_type):
    """是否存在可供指定车辆使用的有效以租代售方案。"""
    normalized_car_type = normalize_base_car_type(car_type)
    normalized_condition = normalize_vehicle_condition(condition)
    normalized_box_type = (box_type or '').strip()
    if not normalized_car_type or not normalized_box_type:
        return False

    rows = conn.execute(
        """SELECT car_type, condition, box_type FROM finance_plans
           WHERE status='启用'
             AND (effective_date IS NULL OR effective_date = '' OR effective_date <= date('now'))
             AND (expiry_date IS NULL OR expiry_date = '' OR expiry_date >= date('now'))"""
    ).fetchall()
    return any(
        normalize_base_car_type(row['car_type']) == normalized_car_type
        and normalize_vehicle_condition(row['condition']) == normalized_condition
        and (row['box_type'] or '').strip() == normalized_box_type
        for row in rows
    )


def normalize_approval_step_label(ref_type, step_order, required_role, step_label):
    """兼容历史审批流文案，确保页面展示使用当前标准名称。"""
    if ref_type == 'initial_payment' and step_order == 1 and required_role == '财务':
        return '财务审核收款'
    if ref_type == 'initial_payment_shortage':
        return '首次付款不足出库审批'
    if ref_type == 'return_stock' and step_order == 1 and required_role == '运营':
        return '运营查车辆数据'
    if ref_type == 'return_stock' and required_role == '老板':
        return '领导审批'
    if ref_type == 'price_exception':
        return '价格审批'
    if ref_type == 'order_exception':
        return '报单异常审批'
    return step_label


def create_approval_flow(conn, ref_type, ref_id):
    """创建审批流程，返回 batch_no"""
    config = APPROVAL_CONFIGS.get(ref_type, [])
    batch_no = f"{ref_type}_{ref_id}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    for step_cfg in config:
        conn.execute(
            "INSERT INTO approval_flows (ref_type, ref_id, batch_no, step_order, required_role, step_label) VALUES (?,?,?,?,?,?)",
            (ref_type, ref_id, batch_no, step_cfg['step'], step_cfg['role'], step_cfg['label'])
        )
    return batch_no


def migrate_legacy_pending_approvals(conn):
    """旧版本曾有合同/常规审批；新流程取消合同财务审批，只保留必要特批。"""
    conn.execute("UPDATE sales_orders SET order_status='待价格特批' WHERE order_status='待老板价格审批'")
    conn.execute("UPDATE sales_orders SET order_status='已作废' WHERE order_status IN ('价格特批驳回', '财务驳回')")
    conn.execute("""
        UPDATE sales_orders
        SET order_status='待财务确认'
        WHERE order_status='已转合同'
          AND COALESCE(finance_confirmed_at, '')=''
    """)
    conn.execute("""
        UPDATE contracts
        SET delivery_status='待首付款'
        WHERE delivery_status IN ('待审批', '审批中')
          AND contract_status='执行中'
    """)
    conn.execute("""
        UPDATE approval_flows
        SET status='已取消',
            comment=COALESCE(NULLIF(comment, ''), '系统迁移：合同不再需要财务审批'),
            acted_at=COALESCE(acted_at, datetime('now','localtime'))
        WHERE status='待审批'
          AND (
            ref_type='contract_delivery'
            OR (ref_type='initial_payment' AND required_role!='财务')
            OR (ref_type='initial_payment_shortage' AND required_role!='老板')
            OR (ref_type='sale_payment' AND required_role!='财务')
            OR (ref_type='lock_request')
            OR (ref_type='return_stock' AND required_role!='老板')
          )
    """)
    conn.execute("""
        INSERT INTO approval_flows (ref_type, ref_id, batch_no, step_order, required_role, step_label, status, created_at)
        SELECT 'sale_payment', so.id,
               'sale_payment_' || so.id || '_seed',
               1,
               '财务',
               '财务确认报单',
               '待审批',
               datetime('now','localtime')
        FROM sales_orders so
        WHERE so.order_status = '待财务确认'
          AND NOT EXISTS (
              SELECT 1
              FROM approval_flows af
              WHERE af.ref_type = 'sale_payment'
                AND af.ref_id = so.id
                AND af.status = '待审批'
          )
    """)


def is_price_below_guidance(price, guidance_price):
    """指导价为 0 或未维护时不触发老板特批；缺失阻塞由报单入口处理。"""
    return float(guidance_price or 0) > 0 and float(price or 0) < float(guidance_price or 0)


def normalize_sales_mode(value):
    """仅支持 租赁 / 以租代售 两种模式（整车销售已下线）。"""
    mode = (value or '').strip()
    if mode in ('以租代购', '以租代售'):
        return '以租代售'
    return '租赁'


def contract_type_from_sales_mode(value):
    return normalize_sales_mode(value)


def resolve_guidance_prices_for_vehicle(conn, vehicle):
    """返回车型租赁参考价；历史 guidance_price 仅作展示。按车辆成色（新车/二手车）匹配对应行。"""
    car_type = (vehicle['car_type'] if vehicle and 'car_type' in vehicle.keys() else '') or ''
    legacy_price = parse_money(vehicle['guidance_price'] if vehicle and 'guidance_price' in vehicle.keys() else 0)
    is_new = (vehicle['condition'] if vehicle and 'condition' in vehicle.keys() else '') or ''
    result = {
        'lease_installment_price': 0.0,
        'legacy_guidance_price': legacy_price,
        'source': '单车指导价' if legacy_price > 0 else '',
    }
    if car_type:
        candidates = guidance_base_candidates(car_type)
        sql = """SELECT guidance_price, lease_installment_price
            FROM model_guidance_prices WHERE car_type IN ({})""".format(
                ','.join('?' for _ in candidates)
            )
        params = list(candidates)
        if is_new in ('新车', '二手车'):
            sql += " AND is_new=?"
            params.append(is_new)
        row = conn.execute(sql, params).fetchone()
        if row:
            lease_price = parse_money(row['lease_installment_price'])
            legacy_model_price = parse_money(row['guidance_price'])
            result.update({
                'lease_installment_price': lease_price,
                'legacy_guidance_price': legacy_model_price or legacy_price,
                'source': '车型指导价',
            })
    return result


def calculate_guidance_check(conn, vehicle, sales_mode, sale_total_price, lease_quote, upfront_amount,
                             vehicle_box_type=None, tail_plate=None, finance_plan_id=None):
    """按业务模式计算指导价核对结果（20260804 改版）。

    租赁：押金 + 月供双维度均不得低于指导价（押金=车型押金指导价；月供=箱型5选1+尾板+300/0）。
    以租代售：无价格特批，只校验金融方案是否已选。
    below/missing 为异常项，供报单层合并成老板一次审批。
    """
    guidance = resolve_guidance_prices_for_vehicle(conn, vehicle)
    mode = contract_type_from_sales_mode(sales_mode)
    installment = parse_money(lease_quote)
    deposit_quote = parse_money(upfront_amount)
    missing = []
    below = []

    if mode == '租赁':
        car_type = (vehicle or {}).get('car_type') if isinstance(vehicle, dict) else (vehicle['car_type'] if vehicle else '')
        is_new = ''
        if vehicle:
            if isinstance(vehicle, dict):
                is_new = vehicle.get('condition') or vehicle.get('is_new') or ''
            else:
                try:
                    is_new = vehicle['condition'] or ''
                except (IndexError, KeyError):
                    is_new = ''
        g = resolve_lease_guidance(conn, car_type, vehicle_box_type, tail_plate, is_new=is_new or None)
        result = {
            'mode': mode,
            'base_price': guidance['lease_installment_price'] or guidance['legacy_guidance_price'],
            'quote_price': installment,
            'guidance_price': g['monthly_guidance'],
            'guidance': g,
            'checks': [],
            'missing': missing,
            'below': below,
            'needs_approval': False,
        }
        if g['deposit_guidance'] <= 0 or g['monthly_guidance'] <= 0:
            missing.append('缺少租赁指导价（押金或箱型月供）')
        if is_price_below_guidance(deposit_quote, g['deposit_guidance']):
            below.append(f"押金 ¥{round(deposit_quote,2)} < 指导押金 ¥{round(g['deposit_guidance'],2)}")
        if is_price_below_guidance(installment, g['monthly_guidance']):
            below.append(f"月供 ¥{round(installment,2)} < 指导月供 ¥{round(g['monthly_guidance'],2)}")
        return result

    if mode == '以租代售':
        car_type = (vehicle or {}).get('car_type') if isinstance(vehicle, dict) else (vehicle['car_type'] if vehicle else '')
        condition = ''
        if vehicle:
            if isinstance(vehicle, dict):
                condition = vehicle.get('condition') or vehicle.get('is_new') or ''
            else:
                try:
                    condition = vehicle['condition'] or ''
                except (IndexError, KeyError):
                    condition = ''
        plan = resolve_finance_plan(
            conn, finance_plan_id, car_type=car_type,
            condition=condition or '新车', box_type=vehicle_box_type,
        )
        if not plan:
            missing.append('请选择与车辆新车厢型一致的以租代售金融方案')
        return {
            'mode': mode,
            'base_price': guidance['lease_installment_price'] or guidance['legacy_guidance_price'],
            'quote_price': installment,
            'guidance_price': 0,
            'guidance': {'plan': plan},
            'checks': [],
            'missing': missing,
            'below': below,
            'needs_approval': False,
        }
    # 兜底：未知模式按租赁处理（normalize_sales_mode 已保证仅 租赁/以租代售）
    return {
        'mode': mode,
        'base_price': guidance['lease_installment_price'] or guidance['legacy_guidance_price'],
        'quote_price': installment,
        'guidance_price': 0,
        'guidance': guidance,
        'checks': [],
        'missing': missing,
        'below': below,
        'needs_approval': False,
    }


# 车辆维度字段 → 数据字典 category 映射（SKU 改造：字典动态化）
VEHICLE_DIM_CATEGORY_MAP = {
    'brand': 'brand',                # 品牌
    'product_series': 'product_series',  # 品系
    'box_type': 'box_type',          # 厢型
    'battery_capacity': 'battery_capacity',  # 电池度数
    'horsepower': 'horsepower',      # 马力
    'gear_position': 'gear_position',  # 档位
    'battery_brand': 'battery_brand',  # 电池品牌
}

# 数据字典只维护基准车型与 SKU 所需维度。颜色、驾驶室仅为车辆实例/
# 报单快照字段；历史的 engine_battery / power_battery / gearbox / box_dimension
# 也不再给操作人员维护。
DICT_CATEGORY_ENERGY_LIMITS = {
    'brand': (),
    'product_series': (),
    'box_type': (),
    'battery_brand': ('纯电',),
    'battery_capacity': ('纯电',),
    'horsepower': ('混动', '燃油车'),
    'gear_position': ('混动', '燃油车'),
}
RETIRED_DICT_CATEGORIES = {
    'box_dimension', 'engine_battery', 'power_battery', 'gearbox',
    'vehicle_color', 'cab_type',
}

# 可选维度（字典缺失时仅 warning 不阻塞）；基准车型不完整或未维护指导价才阻塞报单。
WARNING_DIM_CATEGORIES = set(VEHICLE_DIM_CATEGORY_MAP.values())


def dict_value_exists(conn, category, value, energy_type=''):
    """字典中是否存在可用于指定能源类型的启用值。"""
    if not category or not value:
        return True
    params = [category, value]
    where = "category=? AND value=? AND status='启用'"
    allowed_energies = DICT_CATEGORY_ENERGY_LIMITS.get(category, ())
    if allowed_energies:
        if energy_type not in allowed_energies:
            return False
        where += " AND energy_type=?"
        params.append(energy_type)
    else:
        where += " AND COALESCE(energy_type, '')=''"
    return bool(conn.execute(
        f"SELECT 1 FROM data_dictionaries WHERE {where} LIMIT 1",
        params
    ).fetchone())


def validate_vehicle_dict(conn, car_type, vehicle=None):
    """校验车辆数据字典字段，返回 (validation_status, validation_message)。

    20260810 SKU 改造：
    - car_type 不在租赁指导价字典时，若存在有效的同车型、成色、厢型以租代售方案 → warning
    - 无租赁指导价且无有效以租代售方案 → invalid（阻塞报单）
    - 品牌/品系/厢型等维度值不在 data_dictionaries → warning（不阻塞，提示老板补充字典）
    """
    blocking_errors = []
    warnings = []
    dictionary_warnings = []
    if car_type:
        candidates = guidance_base_candidates(car_type)
        row = conn.execute(
            "SELECT 1 FROM model_guidance_prices WHERE car_type IN ({}) LIMIT 1".format(
                ','.join('?' for _ in candidates)
            ),
            candidates
        ).fetchone()
        if not row:
            condition = normalize_vehicle_condition(
                vehicle.get('condition') if hasattr(vehicle, 'get') else ''
            )
            box_type = (
                (vehicle.get('vehicle_box_type') or vehicle.get('box_type') or '').strip()
                if hasattr(vehicle, 'get') else ''
            )
            if has_active_finance_plan_for_vehicle(conn, car_type, condition, box_type):
                warnings.append('租赁指导价未维护，仅支持按已维护的以租代售方案报单')
            else:
                blocking_errors.append(f"车型「{car_type}」不在已维护的指导价字典中，请先维护车型指导价")

    # 基准车型的能源维度完整性与互斥性（invalid）
    if vehicle:
        identity_errors = base_model_validation_errors(vehicle)
        if identity_errors:
            blocking_errors.extend(identity_errors)

        # 维度值字典检查（warning 级），仅校验当前能源类型适用维度。
        energy_type = energy_type_for_vehicle(vehicle)
        for field, category in VEHICLE_DIM_CATEGORY_MAP.items():
            val = (vehicle.get(field) or '').strip() if hasattr(vehicle, 'get') else ''
            limits = DICT_CATEGORY_ENERGY_LIMITS.get(category, ())
            if limits and energy_type not in limits:
                continue
            if val and not dict_value_exists(conn, category, val, energy_type):
                dictionary_warnings.append(f"「{val}」不在{category}字典中")
        if dictionary_warnings:
            warnings.append('数据字典待补充：' + '；'.join(dictionary_warnings))

    if blocking_errors:
        return ('invalid', '；'.join(blocking_errors + warnings))
    if warnings:
        return ('warning', '；'.join(warnings))
    return ('valid', '')


def guidance_exception_reason(result):
    parts = []
    if result.get('guidance_price', 0) <= 0:
        parts.append('缺少指导价')
    elif result.get('quote_price', 0) > 0 and result.get('guidance_price', 0) > 0 and result['quote_price'] < result['guidance_price']:
        parts.append(f"报价低于指导价 ¥{result['quote_price']} < ¥{result['guidance_price']}")
    return '；'.join(parts)


def expire_sales_order_drafts(conn):
    conn.execute("""
        UPDATE sales_orders
        SET order_status='已作废',
            voided_at=COALESCE(voided_at, datetime('now','localtime')),
            void_reason=COALESCE(NULLIF(void_reason, ''), '草稿超过7天自动作废'),
            voided_by=COALESCE(NULLIF(voided_by, ''), '系统')
        WHERE order_status='草稿'
          AND expires_at IS NOT NULL
          AND expires_at < datetime('now','localtime')
    """)


def role_values_from_db(conn, table, key_col, role, fallback):
    rows = conn.execute(
        f"SELECT {key_col} FROM {table} WHERE role=? ORDER BY id ASC",
        (role,)
    ).fetchall()
    values = [row[key_col] for row in rows]
    return values or fallback.get(role, [])


def role_pages_for(conn, role):
    return role_values_from_db(conn, 'role_pages', 'page_key', role, ROLE_PAGES)


def role_actions_for(conn, role):
    return role_values_from_db(conn, 'role_actions', 'action_key', role, ROLE_ACTIONS)


def role_hidden_fields_for(conn, role):
    rows = conn.execute("""
        SELECT resource_key, field_key
        FROM role_field_permissions
        WHERE role=? AND COALESCE(can_view, 1)=0
        ORDER BY id ASC
    """, (role,)).fetchall()
    data = {}
    for row in rows:
        data.setdefault(row['resource_key'], []).append(row['field_key'])
    return data or ROLE_HIDDEN_FIELDS.get(role, {})


def redact_for_role(conn, role, resource_key, payload):
    hidden = set(role_hidden_fields_for(conn, role).get(resource_key, []))
    if not hidden:
        return payload
    if '*' in hidden:
        return [] if isinstance(payload, list) else {}

    def redact_row(row):
        item = dict(row)
        for field in hidden:
            item.pop(field, None)
        return item

    if isinstance(payload, list):
        return [redact_row(row) for row in payload]
    if isinstance(payload, dict):
        return redact_row(payload)
    return payload


def normalize_return_inspection(row):
    if not row:
        return None
    data = dict(row)
    for key in [
        'tool_triangle', 'tool_vest', 'tool_extinguisher', 'tool_wedge', 'tool_jack',
        'tool_kit', 'tent_pole', 'car_wash_fee', 'body_ad_clean',
        'doc_license', 'doc_keys', 'boss_approved', 'finance_approved', 'paid_out',
    ]:
        data[key] = bool(data.get(key))
    return data


def build_return_inspection_payload(conn, rid):
    c = conn.cursor()
    c.execute("""
        SELECT ri.*, v.vin as vehicle_vin, v.plate_number as vehicle_plate_number, v.car_type as vehicle_car_type,
               v.company as vehicle_company, c.company as contract_company, c.yard as contract_yard,
               c.lease_bank_name, c.lease_bank_card_no, c.contract_status, c.delivery_status as contract_delivery_status,
               cu.name as lease_customer_name, cu.phone as lease_customer_phone
        FROM return_inspections ri
        LEFT JOIN vehicles v ON v.id = ri.vehicle_id
        LEFT JOIN contracts c ON c.id = ri.contract_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE ri.id=?
    """, (rid,))
    row = c.fetchone()
    return normalize_return_inspection(row)


def return_inspection_status_for_step(row):
    if not row:
        return '待登记'
    if row.get('status') == '已驳回待销售修改':
        return '已驳回待销售修改'
    if row.get('paid_out'):
        return '已完成'
    if row.get('boss_approved'):
        return '待出款'
    if row.get('finance_approved'):
        return '待领导审批'
    if row.get('operator_status') == '已填写':
        return '待财务复核'
    if row.get('fleet_status') == '已填写':
        return '待运营填写'
    if row.get('sales_status') == '已登记':
        return '待车管验车'
    return row.get('status') or '待登记'


def update_return_inspection_status(conn, rid):
    c = conn.cursor()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        return
    status = return_inspection_status_for_step(dict(row))
    c.execute("UPDATE return_inspections SET status=? WHERE id=?", (status, rid))


def resolve_guidance_price_for_vehicle(conn, vehicle):
    """历史调用兼容：返回租赁月供参考价优先的单值（sale_total_price 已废弃）。"""
    prices = resolve_guidance_prices_for_vehicle(conn, vehicle)
    price = prices['lease_installment_price'] or prices['legacy_guidance_price']
    return price, prices['source'] or '指导价'


def unresolved_guidance_vehicle_count(conn):
    """统计未设置指导价的车型数（按基准车型 + 成色去重，20260807 方案一）。

    底盘车（box_type='底盘'）不参与统计。
    """
    c = conn.cursor()
    c.execute("""
        SELECT car_type, COALESCE(NULLIF(condition, ''), '新车') as cond, COUNT(*) as cnt
        FROM vehicles
        WHERE (is_deleted IS NULL OR is_deleted = 0)
          AND COALESCE(box_type, '') != '底盘'
          AND COALESCE(status, '') IN ('在库', '报单锁定中')
        GROUP BY car_type, cond
    """)
    vehicle_groups = [dict(r) for r in c.fetchall()]
    c.execute("""
        SELECT car_type, is_new,
               guidance_price, lease_deposit_guidance,
               box_standard_price, box_wide_price, box_high_rail_price,
               box_refrigerated_price, box_flatbed_price
        FROM model_guidance_prices
    """)
    guidance_rows = [dict(r) for r in c.fetchall()]

    def _configured(r):
        return any(parse_money(r.get(k)) > 0 for k in (
            'guidance_price', 'lease_deposit_guidance',
            'box_standard_price', 'box_wide_price', 'box_high_rail_price',
            'box_refrigerated_price', 'box_flatbed_price'))

    configured_keys = set()
    for r in guidance_rows:
        base = normalize_base_car_type(r['car_type'])
        if _configured(r):
            configured_keys.add((base, r['is_new']))

    missing = 0
    for v in vehicle_groups:
        key = (normalize_base_car_type(v['car_type']), v['cond'])
        if key not in configured_keys:
            missing += 1
    return missing


def get_approval_status(conn, ref_type, ref_id):
    """获取最新一轮审批流程状态"""
    c = conn.cursor()
    c.execute("SELECT * FROM approval_flows WHERE ref_type=? AND ref_id=? ORDER BY batch_no DESC, step_order ASC", (ref_type, ref_id))
    rows = [dict(r) for r in c.fetchall()]
    for row in rows:
        row['step_label'] = normalize_approval_step_label(
            row.get('ref_type'),
            row.get('step_order'),
            row.get('required_role'),
            row.get('step_label', '')
        )
    if not rows:
        return {'steps': [], 'current_step': 0, 'status': 'none', 'batch_no': ''}
    latest_batch = rows[0]['batch_no']
    steps = [r for r in rows if r['batch_no'] == latest_batch]
    current_step = 0
    overall_status = '已完成'
    for s in steps:
        if s['status'] == '已驳回':
            overall_status = '已驳回'
            break
        if s['status'] == '待审批':
            if current_step == 0:
                current_step = s['step_order']
            overall_status = '审批中'
    return {'steps': steps, 'current_step': current_step, 'status': overall_status, 'batch_no': latest_batch}


def get_current_user():
    """从session获取当前登录用户"""
    user_id = session.get('user_id')
    if not user_id:
        return None
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, username, display_name, role FROM users WHERE id=? AND is_active=1", (user_id,))
    user = c.fetchone()
    conn.close()
    return dict(user) if user else None


def get_initial_payment_amount(contract):
    """根据合同类型计算首次付款默认金额。"""
    contract_type = contract['contract_type'] if contract else ''
    def value(key):
        return contract[key] if contract and key in contract.keys() else 0
    if contract_type == '以租代售':
        return float(value('down_payment') or 0)
    return float(value('deposit') or 0) + float(value('rent') or 0)


def initial_payment_label(contract):
    if not contract:
        return '首付款审核'
    if contract['contract_type'] == '以租代售':
        return '首付款审核'
    return '押金及首次支付审核'


def finalize_initial_payment(conn, payment_id, operator_name, now, allow_shortage=False):
    c = conn.cursor()
    c.execute("""
        SELECT ip.*, c.contract_type, c.customer_id, c.deposit, c.down_payment, c.rent,
               c.sales_order_id, c.vehicle_id
        FROM contract_initial_payments ip
        JOIN contracts c ON c.id = ip.contract_id
        WHERE ip.id=?
    """, (payment_id,))
    payment = c.fetchone()
    if not payment:
        raise ValueError('首次付款记录不存在')

    contract_id = payment['contract_id']
    expected_amount = get_initial_payment_amount(payment)
    received_amount = parse_money(payment['received_amount'])
    shortage_amount = round(max(0, expected_amount - received_amount), 2)
    if shortage_amount > 0 and not allow_shortage:
        return {
            'needs_shortage_approval': True,
            'shortage_amount': shortage_amount,
            'expected_amount': expected_amount,
            'received_amount': received_amount,
        }

    remaining_received = received_amount
    deposit_due = parse_money(payment['deposit']) if payment['contract_type'] == '租赁' else 0
    down_due = parse_money(payment['down_payment']) if payment['contract_type'] == '以租代售' else 0
    rent_due = parse_money(payment['rent']) if payment['contract_type'] == '租赁' else 0
    deposit_paid = min(remaining_received, deposit_due) if deposit_due > 0 else 0
    remaining_received = round(max(0, remaining_received - deposit_paid), 2)
    down_paid = min(remaining_received, down_due) if down_due > 0 else 0
    remaining_received = round(max(0, remaining_received - down_paid), 2)
    rent_paid = min(remaining_received, rent_due) if rent_due > 0 else 0

    updates = ["delivery_status='待出库'"]
    params = []
    if deposit_due > 0:
        updates.append("deposit_status=?")
        params.append('已收' if shortage_amount <= 0 or deposit_paid >= deposit_due else '部分已收')
        updates.append("collected_deposit=?")
        params.append(deposit_paid if shortage_amount > 0 else deposit_due)
    if down_due > 0:
        updates.append("down_payment_status=?")
        params.append('已收' if shortage_amount <= 0 or down_paid >= down_due else '部分已收')
    if payment['contract_type'] == '租赁' and rent_due > 0:
        updates.append("collected_rent=COALESCE(collected_rent,0)+?")
        params.append(rent_paid if shortage_amount > 0 else rent_due)
    params.append(contract_id)
    c.execute(f"UPDATE contracts SET {', '.join(updates)} WHERE id=?", params)

    c.execute("""
            UPDATE repayments
            SET status=CASE
                    WHEN ? > 0 AND MIN(amount, ?) < amount THEN '部分核销'
                    ELSE '已还款'
                END,
                paid_amount=CASE WHEN ? > 0 THEN MIN(amount, ?) ELSE amount END,
                verified_amount=CASE WHEN COALESCE(verified_amount, 0)=0 THEN CASE WHEN ? > 0 THEN MIN(amount, ?) ELSE amount END ELSE verified_amount END,
                paid_at=?,
                screenshot_path=COALESCE(screenshot_path, ?),
                bank_receipt_path=COALESCE(bank_receipt_path, ?),
                bank_serial=COALESCE(bank_serial, ?),
                verified_by=COALESCE(verified_by, ?),
                verified_at=COALESCE(verified_at, ?),
                waterfall_summary=COALESCE(waterfall_summary, ?),
                remark=COALESCE(remark, CASE WHEN ?='以租代售' THEN '首付款' ELSE '押金' END)
            WHERE contract_id=? AND period=0
        """, (
            shortage_amount,
            deposit_paid if payment['contract_type'] == '租赁' else down_paid,
            shortage_amount,
            deposit_paid if payment['contract_type'] == '租赁' else down_paid,
            shortage_amount,
            deposit_paid if payment['contract_type'] == '租赁' else down_paid,
            datetime.now().strftime('%Y-%m-%d'),
            payment['customer_screenshot_path'],
            payment['bank_receipt_path'],
            payment['bank_serial'] or f"INITIAL-{payment_id}",
            operator_name,
            now,
            'initial_payment 不足额出库，差额挂账应收' if shortage_amount > 0 else 'initial_payment 押金/首付款独立核销',
            payment['contract_type'],
            contract_id,
        ))

    if payment['contract_type'] == '租赁' and rent_due > 0:
        c.execute("""
            UPDATE repayments
            SET status=CASE
                    WHEN ? > 0 AND MIN(amount, ?) < amount THEN '部分核销'
                    ELSE '已还款'
                END,
                paid_amount=CASE WHEN ? > 0 THEN MIN(amount, ?) ELSE amount END,
                verified_amount=CASE WHEN COALESCE(verified_amount, 0)=0 THEN CASE WHEN ? > 0 THEN MIN(amount, ?) ELSE amount END ELSE verified_amount END,
                paid_at=?,
                screenshot_path=COALESCE(screenshot_path, ?),
                bank_receipt_path=COALESCE(bank_receipt_path, ?),
                bank_serial=COALESCE(bank_serial, ?),
                verified_by=COALESCE(verified_by, ?),
                verified_at=COALESCE(verified_at, ?),
                waterfall_summary=COALESCE(waterfall_summary, ?),
                remark=COALESCE(remark, '首次付款审核自动核销首期租金')
            WHERE id=(
                SELECT id FROM (
                    SELECT id FROM repayments
                    WHERE contract_id=? AND period>=1 AND status!='已还款'
                    ORDER BY period ASC LIMIT 1
                ) _sub
            )
        """, (
            shortage_amount,
            rent_paid,
            shortage_amount,
            rent_paid,
            shortage_amount,
            rent_paid,
            datetime.now().strftime('%Y-%m-%d'),
            payment['customer_screenshot_path'],
            payment['bank_receipt_path'],
            payment['bank_serial'] or f"INITIAL-{payment_id}",
            operator_name,
            now,
            'initial_payment 不足额出库，差额挂账应收' if shortage_amount > 0 else 'initial_payment 首期租金自动核销',
            contract_id,
        ))
        # 首次付款是在本次审核日入账；若计划被提前生成并误跑了日终，
        # 仅冲回本次实际收款日及之后的错误滞纳金，历史真实逾期仍保留。
        first_rent_row = c.execute("""
            SELECT id FROM repayments
            WHERE contract_id=? AND period=1
            ORDER BY id ASC LIMIT 1
        """, (contract_id,)).fetchone()
        if first_rent_row:
            reverse_late_fee_accruals_from_payment_date(
                conn,
                first_rent_row['id'],
                now[:10],
                '首次付款审核自动核销首期租金',
            )

    if shortage_amount > 0:
        create_or_update_receivable(
            conn,
            contract_id,
            'initial_payment_shortfall',
            shortage_amount,
            initial_payment_id=payment_id,
            promised_repay_date=payment['promised_repay_date'],
            reason=payment['shortage_reason'] or '首次付款不足',
            status='待归还',
            created_by=operator_name,
        )
        c.execute("""
            UPDATE contract_initial_payments
            SET status='已通过', shortage_amount=?, shortage_status='挂账中', approved_by=?, approved_at=?
            WHERE id=?
        """, (shortage_amount, operator_name, now, payment_id))
    else:
        c.execute("UPDATE contract_initial_payments SET status='已通过', shortage_amount=0, shortage_status='无欠款', approved_by=?, approved_at=? WHERE id=?",
                  (operator_name, now, payment_id))

    c.execute("""
        UPDATE repayments
        SET status='待还款'
        WHERE contract_id=? AND status='未激活'
    """, (contract_id,))

    extra_amount = round(received_amount - expected_amount, 2)
    if extra_amount > 0:
        c.execute("""
            INSERT INTO customer_prepayments
                (contract_id, customer_id, amount, source_repayment_id, source_bank_serial, balance, remark)
            VALUES (?, ?, ?, NULL, ?, ?, ?)
        """, (
            contract_id,
            payment['customer_id'],
            extra_amount,
            payment['bank_serial'] or f"INITIAL-{payment_id}",
            extra_amount,
            '首次付款多收款挂账，待客户确认抵充月份',
        ))

    return {
        'needs_shortage_approval': False,
        'shortage_amount': shortage_amount,
        'expected_amount': expected_amount,
        'received_amount': received_amount,
        'contract_id': contract_id,
    }


def parse_money(value, default=0):
    try:
        if value in (None, ''):
            return float(default)
        if isinstance(value, str):
            value = value.replace(',', '').strip()
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def parse_optional_money(value):
    if value in (None, ''):
        return None
    try:
        if isinstance(value, str):
            value = value.replace(',', '').strip()
        return float(value)
    except (TypeError, ValueError):
        return None


def normalize_expected_profit_ceiling(value):
    amount = parse_optional_money(value)
    if amount is None or amount <= 0 or amount >= PROFIT_CEILING_SENTINEL:
        return None
    return amount


def normalize_date(value):
    if not value:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value)[:10]


def merge_attachment_csv(*values):
    merged = []
    seen = set()
    for value in values:
        for url in str(value or '').split(','):
            clean = url.strip()
            if clean and clean not in seen:
                merged.append(clean)
                seen.add(clean)
    return ','.join(merged)


def upload_url_to_path(file_url):
    """把 /uploads/xxx 或相对路径转换为本地绝对路径。"""
    if not file_url:
        return ''
    clean_url = str(file_url).strip()
    if clean_url.startswith('/uploads/'):
        return os.path.join(UPLOAD_DIR, os.path.basename(clean_url))
    if os.path.isabs(clean_url):
        return clean_url
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), clean_url.lstrip('/'))


def create_or_update_receivable(conn, contract_id, receivable_type, amount, **kwargs):
    amount = round(parse_money(amount), 2)
    if amount <= 0:
        return None
    c = conn.cursor()
    repayment_id = kwargs.get('repayment_id')
    initial_payment_id = kwargs.get('initial_payment_id')
    sales_order_id = kwargs.get('sales_order_id')
    existing = None
    if sales_order_id:
        existing = c.execute("""
            SELECT id, paid_amount
            FROM receivables
            WHERE sales_order_id=? AND receivable_type=?
            ORDER BY id DESC LIMIT 1
        """, (sales_order_id, receivable_type)).fetchone()
    elif initial_payment_id:
        existing = c.execute("""
            SELECT id, paid_amount
            FROM receivables
            WHERE initial_payment_id=? AND receivable_type=?
            ORDER BY id DESC LIMIT 1
        """, (initial_payment_id, receivable_type)).fetchone()
    elif repayment_id:
        existing = c.execute("""
            SELECT id, paid_amount
            FROM receivables
            WHERE repayment_id=? AND receivable_type=?
            ORDER BY id DESC LIMIT 1
        """, (repayment_id, receivable_type)).fetchone()
    if existing:
        paid = parse_money(existing['paid_amount'])
        status = '已结清' if paid >= amount else ('逾期应收' if kwargs.get('status') == '逾期应收' else '待归还')
        c.execute("""
            UPDATE receivables
            SET amount=?, due_date=?, promised_repay_date=?, reason=?, status=?
            WHERE id=?
        """, (
            amount,
            kwargs.get('due_date') or '',
            kwargs.get('promised_repay_date') or '',
            kwargs.get('reason') or '',
            status,
            existing['id'],
        ))
        return existing['id']
    c.execute("""
        INSERT INTO receivables
            (contract_id, repayment_id, initial_payment_id, receivable_type, source_period,
             amount, due_date, promised_repay_date, reason, status, created_by, sales_order_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        contract_id,
        repayment_id,
        initial_payment_id,
        receivable_type,
        kwargs.get('source_period'),
        amount,
        kwargs.get('due_date') or '',
        kwargs.get('promised_repay_date') or '',
        kwargs.get('reason') or '',
        kwargs.get('status') or '待归还',
        kwargs.get('created_by') or '系统',
        sales_order_id,
    ))
    return c.lastrowid


def settle_receivable_payment(conn, receivable_id, amount):
    c = conn.cursor()
    row = c.execute("SELECT amount, paid_amount FROM receivables WHERE id=?", (receivable_id,)).fetchone()
    if not row:
        return 0
    outstanding = round(max(0, parse_money(row['amount']) - parse_money(row['paid_amount'])), 2)
    paid_now = round(parse_money(amount), 2)
    if paid_now <= 0 or paid_now > outstanding:
        raise ValueError(f'归还金额必须大于0且不超过剩余应收 ¥{outstanding}')
    paid = round(parse_money(row['paid_amount']) + paid_now, 2)
    due = parse_money(row['amount'])
    status = '已结清' if paid >= due else '部分归还'
    c.execute("""
        UPDATE receivables
        SET paid_amount=?, status=?, settled_at=CASE WHEN ?='已结清' THEN datetime('now','localtime') ELSE settled_at END
        WHERE id=?
    """, (paid, status, status, receivable_id))
    return paid_now


def apply_receivable_payment_to_repayments(conn, receivable_id, amount, operator_name, bank_serial='', screenshot_path=''):
    """把挂账应收实收同步到原分期，避免应收已结清而计划仍未结清。"""
    c = conn.cursor()
    receivable = c.execute("""
        SELECT rv.*, c.contract_type
        FROM receivables rv
        JOIN contracts c ON c.id=rv.contract_id
        WHERE rv.id=?
    """, (receivable_id,)).fetchone()
    if not receivable:
        return []

    remaining = round(parse_money(amount), 2)
    if remaining <= 0:
        return []
    contract_id = receivable['contract_id']
    if receivable['receivable_type'] == 'period_shortfall' and receivable['repayment_id']:
        rows = c.execute("""
            SELECT id, period, amount, COALESCE(paid_amount,0) AS paid_amount, due_date
            FROM repayments WHERE id=?
        """, (receivable['repayment_id'],)).fetchall()
    elif receivable['receivable_type'] == 'initial_payment_shortfall':
        # 首次付款不足只补首次付款覆盖的项目：押金/首付款，再到租赁首期租金。
        rows = c.execute("""
            SELECT id, period, amount, COALESCE(paid_amount,0) AS paid_amount, due_date
            FROM repayments
            WHERE contract_id=?
              AND (period=0 OR (?='租赁' AND period=1))
            ORDER BY period ASC
        """, (contract_id, receivable['contract_type'])).fetchall()
    else:
        return []

    applied_rows = []
    rent_added = 0
    deposit_paid = None
    for row in rows:
        if remaining <= 0:
            break
        outstanding = round(max(0, parse_money(row['amount']) - parse_money(row['paid_amount'])), 2)
        if outstanding <= 0:
            continue
        allocated = min(remaining, outstanding)
        new_paid = round(parse_money(row['paid_amount']) + allocated, 2)
        status = repayment_status_after_allocation(row['due_date'], new_paid, parse_money(row['amount']))
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        c.execute("""
            UPDATE repayments
            SET paid_amount=?,
                verified_amount=COALESCE(verified_amount,0)+?,
                status=?,
                paid_at=CASE WHEN ?='已还款' THEN ? ELSE paid_at END,
                bank_serial=COALESCE(NULLIF(bank_serial,''), ?),
                screenshot_path=COALESCE(NULLIF(screenshot_path,''), ?),
                verified_by=COALESCE(NULLIF(verified_by,''), ?),
                verified_at=COALESCE(NULLIF(verified_at,''), ?),
                waterfall_summary=COALESCE(NULLIF(waterfall_summary,''), '挂账应收归还同步核销')
            WHERE id=?
        """, (
            new_paid, allocated, status, status, datetime.now().strftime('%Y-%m-%d'),
            bank_serial, screenshot_path, operator_name, now, row['id'],
        ))
        if row['period'] == 0:
            deposit_paid = new_paid
        elif row['period'] >= 1:
            rent_added = round(rent_added + allocated, 2)
        applied_rows.append({'repayment_id': row['id'], 'period': row['period'], 'amount': allocated, 'status': status})
        remaining = round(remaining - allocated, 2)

    if rent_added > 0:
        c.execute("UPDATE contracts SET collected_rent=COALESCE(collected_rent,0)+? WHERE id=?", (rent_added, contract_id))
    if deposit_paid is not None:
        if receivable['contract_type'] == '租赁':
            contract = c.execute("SELECT deposit FROM contracts WHERE id=?", (contract_id,)).fetchone()
            deposit_due = parse_money(contract['deposit']) if contract else 0
            c.execute("""
                UPDATE contracts
                SET collected_deposit=?,
                    deposit_status=CASE WHEN ? >= ? THEN '已收' ELSE '部分已收' END
                WHERE id=?
            """, (deposit_paid, deposit_paid, deposit_due, contract_id))
        elif receivable['contract_type'] == '以租代售':
            contract = c.execute("SELECT down_payment FROM contracts WHERE id=?", (contract_id,)).fetchone()
            down_payment_due = parse_money(contract['down_payment']) if contract else 0
            c.execute("""
                UPDATE contracts
                SET down_payment_status=CASE WHEN ? >= ? THEN '已收' ELSE '部分已收' END
                WHERE id=?
            """, (deposit_paid, down_payment_due, contract_id))
    return applied_rows


def repayment_status_after_allocation(due_date, paid_amount, expected_amount):
    if paid_amount >= expected_amount:
        return '已还款'
    if paid_amount > 0:
        return '部分核销'
    if due_date and due_date < datetime.now().strftime('%Y-%m-%d'):
        diff = (datetime.now().date() - datetime.strptime(due_date, '%Y-%m-%d').date()).days
        return f'逾期{diff}日' if diff >= 1 else '逾期'
    return '待还款'


def sync_period_shortfall_receivable(conn, repayment_id, repayment_outstanding, due_date='', reason_suffix=''):
    """让少还挂账始终等于分期当前未还余额，供补款、减免等反向业务复用。"""
    c = conn.cursor()
    rows = c.execute("""
        SELECT id, COALESCE(paid_amount, 0) AS paid_amount, reason
        FROM receivables
        WHERE repayment_id=?
          AND receivable_type='period_shortfall'
          AND status NOT IN ('已结清', '已取消')
        ORDER BY id ASC
    """, (repayment_id,)).fetchall()
    if not rows:
        return

    outstanding = round(max(0, parse_money(repayment_outstanding)), 2)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for row in rows:
        paid = parse_money(row['paid_amount'])
        new_amount = round(paid + outstanding, 2)
        if outstanding <= 0:
            status = '已结清' if paid > 0 else '已取消'
        elif paid >= new_amount:
            status = '已结清'
        else:
            status = '逾期应收' if due_date and due_date < datetime.now().strftime('%Y-%m-%d') else '待归还'
        reason = row['reason'] or ''
        if reason_suffix and reason_suffix not in reason:
            reason = f"{reason}；{reason_suffix}".strip('；')
        c.execute("""
            UPDATE receivables
            SET amount=?,
                status=?,
                settled_at=CASE WHEN ? IN ('已结清', '已取消') THEN COALESCE(settled_at, ?) ELSE NULL END,
                reason=?
            WHERE id=?
        """, (new_amount, status, status, now, reason, row['id']))


def settle_contract_shortfall_receivables(conn, contract_id):
    """合同被一次性结清时关闭由分期/首次付款不足衍生出的重复挂账视图。"""
    c = conn.cursor()
    c.execute("""
        UPDATE receivables
        SET paid_amount=amount,
            status='已结清',
            settled_at=COALESCE(settled_at, datetime('now','localtime'))
        WHERE contract_id=?
          AND receivable_type IN ('period_shortfall', 'initial_payment_shortfall')
          AND status NOT IN ('已结清', '已取消')
    """, (contract_id,))


def close_covered_initial_payment_shortfalls(conn, contract_id, operator_name='系统', reason=''):
    """首付款差额已由押金/首期租金实际覆盖时，同步关闭重复的挂账应收。"""
    c = conn.cursor()
    receivables = c.execute("""
        SELECT id
        FROM receivables
        WHERE contract_id=?
          AND receivable_type='initial_payment_shortfall'
          AND status NOT IN ('已结清', '已取消')
          AND amount > COALESCE(paid_amount, 0)
    """, (contract_id,)).fetchall()
    if not receivables:
        return []

    contract = c.execute(
        "SELECT contract_type FROM contracts WHERE id=?",
        (contract_id,),
    ).fetchone()
    if not contract:
        return []
    required_periods = [0, 1] if contract['contract_type'] == '租赁' else [0]
    placeholders = ','.join('?' for _ in required_periods)
    repayments = c.execute(f"""
        SELECT period, amount, COALESCE(paid_amount, 0) AS paid_amount
        FROM repayments
        WHERE contract_id=? AND period IN ({placeholders})
    """, [contract_id, *required_periods]).fetchall()
    by_period = {row['period']: row for row in repayments}
    if any(
        period not in by_period
        or parse_money(by_period[period]['paid_amount']) < parse_money(by_period[period]['amount'])
        for period in required_periods
    ):
        return []

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    closed_ids = []
    for receivable in receivables:
        c.execute("""
            UPDATE receivables
            SET paid_amount=amount,
                status='已结清',
                settled_at=COALESCE(settled_at, ?),
                verified_by=COALESCE(NULLIF(verified_by, ''), ?),
                verified_at=COALESCE(verified_at, ?)
            WHERE id=?
        """, (now, operator_name, now, receivable['id']))
        closed_ids.append(receivable['id'])
        log_audit(
            conn,
            '自动结清首次付款不足',
            'receivable',
            receivable['id'],
            f'押金及首期租金已全额核销，关闭重复挂账'
            + (f'；{reason}' if reason else ''),
            operator_name,
        )
    return closed_ids


def sync_fee_item_status(conn, fee_item_id):
    c = conn.cursor()
    c.execute("SELECT amount_due, amount_paid FROM contract_fee_items WHERE id=?", (fee_item_id,))
    row = c.fetchone()
    if not row:
        return
    due = parse_money(row['amount_due'])
    paid = parse_money(row['amount_paid'])
    if paid <= 0:
        status = '待支付'
    elif paid >= due:
        status = '已支付'
    else:
        status = '部分支付'
    c.execute("UPDATE contract_fee_items SET status=? WHERE id=?", (status, fee_item_id))


def normalize_extra_alloc_periods(value):
    if value in (None, '', []):
        return []
    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            if isinstance(parsed, list):
                value = parsed
            else:
                value = [part.strip() for part in value.split(',')]
        except ValueError:
            value = [part.strip() for part in value.split(',')]
    if not isinstance(value, list):
        value = [value]
    result = []
    for item in value:
        try:
            period = int(item)
        except (TypeError, ValueError):
            continue
        if period > 0 and period not in result:
            result.append(period)
    return result


def find_next_adjustable_repayment(c, contract_id, current_period):
    c.execute("""
        SELECT id, amount, COALESCE(paid_amount, 0) AS paid_amount
        FROM repayments
        WHERE contract_id=?
          AND period>?
          AND status NOT IN ('已还款', '预抵')
        ORDER BY period ASC
        LIMIT 1
    """, (contract_id, current_period))
    return c.fetchone()


def apply_waterfall_allocation(conn, repayment_id, received_amount, operator_name, extra_alloc_periods=None):
    """核销客户回款，并按新规则调整客户还款计划表。

    费用项仍按 T+7 顺序优先分配；进入租金分期后：
    - 少于本期标准租金：实收记本期，差额挂本期应收，不顺延；
    - 多于本期标准租金：必须指定抵扣期数，系统写入对应期的预抵/部分核销。
    """
    c = conn.cursor()
    c.execute("SELECT * FROM repayments WHERE id=?", (repayment_id,))
    repayment = c.fetchone()
    if not repayment:
        raise ValueError('还款记录不存在')

    contract_id = repayment['contract_id']
    expected_amount = parse_money(repayment['amount'])
    received_amount = parse_money(received_amount, expected_amount)
    if received_amount <= 0:
        raise ValueError('到账金额必须大于0')

    remaining = received_amount
    allocation_lines = []
    allocation_rows = []

    c.execute("""
        SELECT *
        FROM contract_fee_items
        WHERE contract_id=? AND amount_due > COALESCE(amount_paid, 0)
        ORDER BY
            CASE fee_type
                WHEN 'insurance_fee' THEN 1
                WHEN 'service_fee' THEN 1
                WHEN 'maintenance_fee' THEN 2
                WHEN 'accident_fee' THEN 2
                WHEN 'penalty_fee' THEN 2
                WHEN 'late_fee' THEN 3
                ELSE 99
            END,
            COALESCE(due_date, created_at) ASC,
            id ASC
    """, (contract_id,))
    fee_items = c.fetchall()

    for item in fee_items:
        if remaining <= 0:
            break
        outstanding = max(0, parse_money(item['amount_due']) - parse_money(item['amount_paid']))
        if outstanding <= 0:
            continue
        allocated = min(remaining, outstanding)
        c.execute(
            "UPDATE contract_fee_items SET amount_paid = COALESCE(amount_paid, 0) + ? WHERE id=?",
            (allocated, item['id'])
        )
        sync_fee_item_status(conn, item['id'])
        allocation_rows.append((
            repayment_id, contract_id, item['id'], item['fee_type'], allocated,
            item['description'] or '', operator_name
        ))
        allocation_lines.append(f"{item['fee_type']} ¥{round(allocated, 2)}")
        remaining = round(remaining - allocated, 2)

    paid_amount = parse_money(repayment['paid_amount'])
    rent_outstanding = max(0, expected_amount - paid_amount)
    rent_allocated = min(remaining, rent_outstanding)
    if rent_allocated > 0:
        c.execute("""
            UPDATE contracts
            SET collected_rent = COALESCE(collected_rent, 0) + ?
            WHERE id=?
        """, (rent_allocated, contract_id))
        allocation_rows.append((
            repayment_id, contract_id, None, 'rent', rent_allocated,
            f"第{repayment['period']}期租金", operator_name
        ))
        allocation_lines.append(f"rent ¥{round(rent_allocated, 2)}")
        paid_amount = round(paid_amount + rent_allocated, 2)
        remaining = round(remaining - rent_allocated, 2)

    receivable_id = None
    shortfall = 0
    if rent_allocated < rent_outstanding:
        shortfall = round(rent_outstanding - rent_allocated, 2)
        receivable_id = create_or_update_receivable(
            conn,
            contract_id,
            'period_shortfall',
            shortfall,
            repayment_id=repayment_id,
            source_period=repayment['period'],
            due_date=repayment['due_date'],
            promised_repay_date=repayment['due_date'],
            reason=f"第{repayment['period']}期少还",
            status='逾期应收' if repayment['due_date'] and repayment['due_date'] < datetime.now().strftime('%Y-%m-%d') else '待归还',
            created_by=operator_name,
        )
        allocation_lines.append(f"shortfall receivable ¥{shortfall}")

    extra_allocated_rows = []
    if remaining > 0:
        periods = normalize_extra_alloc_periods(extra_alloc_periods)
        if not periods:
            raise ValueError('本期多还，请先确认抵扣月份')
        placeholders = ','.join(['?'] * len(periods))
        c.execute(f"""
            SELECT id, period, amount, COALESCE(paid_amount, 0) AS paid_amount, status
            FROM repayments
            WHERE contract_id=?
              AND period IN ({placeholders})
              AND period>?
            ORDER BY period ASC
        """, [contract_id, *periods, repayment['period']])
        target_rows = c.fetchall()
        found_periods = {row['period'] for row in target_rows}
        missing = [p for p in periods if p not in found_periods]
        if missing:
            raise ValueError(f'抵扣期数不存在或不是后续期数: {",".join(str(p) for p in missing)}')

        future_capacity = round(sum(max(0, parse_money(row['amount']) - parse_money(row['paid_amount'])) for row in target_rows), 2)
        if remaining > future_capacity:
            raise ValueError(f'多还金额超过所选抵扣期数剩余应还，最多可抵扣 ¥{future_capacity}')

        extra_remaining = remaining
        for row in target_rows:
            if extra_remaining <= 0:
                break
            due = max(0, parse_money(row['amount']) - parse_money(row['paid_amount']))
            allocated = min(extra_remaining, due)
            if allocated <= 0:
                continue
            new_paid = round(parse_money(row['paid_amount']) + allocated, 2)
            # 规则：被抵扣期如被一整期完全覆盖 -> 直接“已还款”，与来源期共用流水号/还款截图；
            # 仅覆盖一部分（不足一整期）-> “预抵”，差额留待下次还款继续补足。
            fully_covered = new_paid >= parse_money(row['amount'])
            new_status = '已还款' if fully_covered else '预抵'
            source_serial = repayment['bank_serial'] if 'bank_serial' in repayment.keys() else ''
            source_screenshot = repayment['screenshot_path'] if 'screenshot_path' in repayment.keys() else ''
            source_receipt = repayment['bank_receipt_path'] if 'bank_receipt_path' in repayment.keys() else ''
            c.execute("""
                UPDATE repayments
                SET paid_amount=?,
                    verified_amount=COALESCE(verified_amount, 0)+?,
                    status=?,
                    paid_at=?,
                    bank_serial=COALESCE(NULLIF(bank_serial, ''), ?),
                    screenshot_path=COALESCE(NULLIF(screenshot_path, ''), ?),
                    bank_receipt_path=COALESCE(NULLIF(bank_receipt_path, ''), ?),
                    verified_by=COALESCE(NULLIF(verified_by, ''), ?),
                    verified_at=COALESCE(NULLIF(verified_at, ''), ?),
                    extra_alloc_confirmed_by=?,
                    extra_alloc_confirmed_at=?,
                    waterfall_summary=COALESCE(NULLIF(waterfall_summary, ''), ?)
                WHERE id=?
            """, (
                new_paid,
                allocated,
                new_status,
                datetime.now().strftime('%Y-%m-%d'),
                source_serial,
                source_screenshot,
                source_receipt,
                operator_name,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                operator_name,
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                f'多还抵扣来源第{repayment["period"]}期',
                row['id'],
            ))
            c.execute("""
                UPDATE contracts
                SET collected_rent = COALESCE(collected_rent, 0) + ?
                WHERE id=?
            """, (allocated, contract_id))
            allocation_rows.append((
                repayment_id, contract_id, None, 'rent_prepay', allocated,
                f"多还抵扣第{row['period']}期", operator_name
            ))
            extra_allocated_rows.append({'period': row['period'], 'amount': allocated, 'status': new_status})
            allocation_lines.append(f"period {row['period']} prepay ¥{round(allocated, 2)}")
            extra_remaining = round(extra_remaining - allocated, 2)
        remaining = extra_remaining

    if allocation_rows:
        c.executemany("""
            INSERT INTO reconciliation_allocations
                (repayment_id, contract_id, fee_item_id, allocation_type, allocated_amount, note, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, allocation_rows)

    status = repayment_status_after_allocation(repayment['due_date'], paid_amount, expected_amount)
    summary = '；'.join(allocation_lines) if allocation_lines else '未分配'
    c.execute("""
        UPDATE repayments
        SET paid_amount=?,
            verified_amount=COALESCE(verified_amount, 0) + ?,
            status=?,
            waterfall_summary=?,
            paid_at=CASE WHEN ?='已还款' THEN ? ELSE paid_at END,
            extra_alloc_periods=?
        WHERE id=?
    """, (
        paid_amount, received_amount, status, summary,
        status, datetime.now().strftime('%Y-%m-%d'),
        json.dumps(normalize_extra_alloc_periods(extra_alloc_periods), ensure_ascii=False) if extra_allocated_rows else repayment['extra_alloc_periods'] if 'extra_alloc_periods' in repayment.keys() else None,
        repayment_id
    ))

    # 少还应收只是该期剩余未还金额的催收视图。后续补款后要同步结清，
    # 否则同一笔差额会在应收和分期里长期重复出现并被重复计息。
    if shortfall <= 0 and paid_amount >= expected_amount:
        c.execute("""
            UPDATE receivables
            SET paid_amount=amount,
                status='已结清',
                settled_at=COALESCE(settled_at, datetime('now','localtime'))
            WHERE repayment_id=?
              AND receivable_type='period_shortfall'
              AND status NOT IN ('已结清', '已取消')
        """, (repayment_id,))

    closed_initial_shortfalls = close_covered_initial_payment_shortfalls(
        conn,
        contract_id,
        operator_name,
        f'第{repayment["period"]}期回款核销后复核',
    )
    if closed_initial_shortfalls:
        allocation_lines.append(
            f"initial payment shortfall closed #{','.join(str(item) for item in closed_initial_shortfalls)}"
        )
        summary = '；'.join(allocation_lines)
        c.execute(
            "UPDATE repayments SET waterfall_summary=? WHERE id=?",
            (summary, repayment_id),
        )

    return {
        'status': status,
        'paid_amount': paid_amount,
        'verified_amount': round(parse_money(repayment['verified_amount']) + received_amount, 2),
        'rent_allocated': rent_allocated,
        'remaining_unallocated': remaining,
        'shortfall_receivable_id': receivable_id,
        'shortfall_amount': shortfall,
        'extra_allocated_rows': extra_allocated_rows,
        'summary': summary,
    }


def parse_factory_plan_sheet(local_path):
    workbook = load_workbook(local_path, data_only=False)
    sheet = workbook.active
    header_row = None
    column_map = {}

    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        cells = [str(v).replace('\n', '').strip() if v is not None else '' for v in row]
        if '应还款日期' in cells:
            header_row = row_idx
            for idx, cell in enumerate(cells):
                if cell:
                    column_map[cell] = idx
            break

    if not header_row or '应还款日期' not in column_map:
        raise ValueError('未识别到厂家还款计划表头')

    due_idx = column_map['应还款日期']
    amount_idx = None
    for key in ('客户应还金额合计', '应还金额', '金额合计'):
        if key in column_map:
            amount_idx = column_map[key]
            break
    period_idx = column_map.get('序号', 0)

    rows = []
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        first_cell = row[period_idx] if period_idx < len(row) else None
        if first_cell in ('合计', None, ''):
            continue
        due_date = row[due_idx] if due_idx < len(row) else None
        if not due_date:
            continue
        amount = None
        if amount_idx is not None and amount_idx < len(row):
            amount = row[amount_idx]
        rows.append({
            'period': int(first_cell) if str(first_cell).isdigit() else len(rows) + 1,
            'due_date': normalize_date(due_date),
            'amount': parse_money(amount, 0),
        })

    if not rows:
        raise ValueError('计划表中未解析到有效还款行')

    return rows


def parse_factory_plan_pdf(local_path):
    """解析厂家还款计划 PDF（解放/建行格式）。

    每行形如：序号 应还款日期 利率 本金 利息 贴息 罚金 应还金额合计 还款标记
    例：1 2025-01-20 9.0 3,644.96 1,125.00 0.00 0.00 4,769.96 已归还
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader

    reader = PdfReader(local_path)
    text = '\n'.join((page.extract_text() or '') for page in reader.pages)

    # 行：序号(数字) 日期(YYYY-MM-DD) ... 末尾若干金额，取倒数第一个金额作为应还合计
    row_re = re.compile(
        r'^\s*(\d{1,3})\s+(\d{4}-\d{1,2}-\d{1,2})\s+(.+)$'
    )
    money_re = re.compile(r'-?[\d,]+\.\d{2}')

    rows = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith('合计'):
            continue
        m = row_re.match(line)
        if not m:
            continue
        period = int(m.group(1))
        due_date = normalize_date(m.group(2))
        tail = m.group(3)
        monies = money_re.findall(tail)
        # 应还金额合计是「罚金」之后那一列：金额列依次为 本金 利息 贴息 罚金 应还合计
        # 取倒数第一个金额（应还合计）。利率(如 9.0)不含千分位逗号且无小数两位，已被 money_re 过滤。
        amount = parse_money(monies[-1], 0) if monies else 0
        status_text = ''
        if monies:
            status_text = tail[tail.rfind(monies[-1]) + len(monies[-1]):].strip()
        rows.append({
            'period': period,
            'due_date': due_date,
            'amount': amount,
            'remark': f'PDF状态:{status_text}' if status_text else '',
        })

    if not rows:
        raise ValueError('PDF 计划表中未解析到有效还款行')

    rows.sort(key=lambda r: r['period'])
    return rows


def parse_factory_plan_file(local_path):
    lower_path = local_path.lower()
    if lower_path.endswith('.xlsx'):
        return parse_factory_plan_sheet(local_path), {'source_format': 'xlsx'}
    if lower_path.endswith('.pdf'):
        return parse_factory_plan_pdf(local_path), {'source_format': 'factory_plan_pdf'}
    raise ValueError('当前导入器仅支持 xlsx 或 PDF 格式')


def compare_contract_repayment_plans(conn, contract_id):
    c = conn.cursor()
    c.execute("""
        SELECT id, expected_profit_floor, expected_profit_ceiling, sales_order_id
        FROM contracts
        WHERE id=?
    """, (contract_id,))
    contract = c.fetchone()
    if not contract:
        return {'status': '差异待处理', 'message': '合同不存在'}

    c.execute("""
        SELECT period, due_date, amount, COALESCE(remark, '') AS remark
        FROM repayments
        WHERE contract_id=?
        ORDER BY period ASC, id ASC
    """, (contract_id,))
    customer_rows = [
        dict(row) for row in c.fetchall()
        if (row['remark'] or '') != '押金'
    ]
    c.execute("""
        SELECT period, due_date, amount
        FROM factory_repayments
        WHERE contract_id=?
        ORDER BY period ASC, id ASC
    """, (contract_id,))
    factory_rows = [dict(row) for row in c.fetchall()]

    customer_total = round(sum(parse_money(row['amount']) for row in customer_rows), 2)
    factory_total = round(sum(parse_money(row['amount']) for row in factory_rows), 2)
    customer_periods = len(customer_rows)
    factory_periods = len(factory_rows)
    spread_total = round(customer_total - factory_total, 2)
    floor = parse_money(contract['expected_profit_floor'], 0)
    ceiling = normalize_expected_profit_ceiling(contract['expected_profit_ceiling'])

    customer_due_dates = [row['due_date'] for row in customer_rows if row.get('due_date')]
    factory_due_dates = [row['due_date'] for row in factory_rows if row.get('due_date')]

    # 逐月对比明细：以自然月(YYYY-MM)为键合并客户/厂家两边，计算月流水与累计流水
    def _month_key(d):
        return (d or '')[:7]

    cust_by_month = {}
    for row in customer_rows:
        if (row.get('remark') or '') == '首付款':
            continue  # 首付款不参与逐月利差对比
        mk = _month_key(row.get('due_date'))
        if not mk:
            continue
        slot = cust_by_month.setdefault(mk, {'amount': 0.0, 'due_date': row['due_date'], 'period': row['period']})
        slot['amount'] += parse_money(row['amount'])
    fact_by_month = {}
    for row in factory_rows:
        mk = _month_key(row.get('due_date'))
        if not mk:
            continue
        slot = fact_by_month.setdefault(mk, {'amount': 0.0, 'due_date': row['due_date'], 'period': row['period']})
        slot['amount'] += parse_money(row['amount'])

    all_months = sorted(set(cust_by_month) | set(fact_by_month))
    period_details = []
    cumulative = 0.0
    # 首付款作为期初第一行计入累计流水
    down_payment = round(sum(
        parse_money(row['amount']) for row in customer_rows
        if (row.get('remark') or '') == '首付款'
    ), 2)
    seq = 0
    if down_payment:
        cumulative = down_payment
        dp_row = next((r for r in customer_rows if (r.get('remark') or '') == '首付款'), None)
        period_details.append({
            'period': seq,
            'month': '首付款',
            'customer_due_date': (dp_row.get('due_date') if dp_row else '') or '',
            'customer_amount': down_payment,
            'factory_due_date': '',
            'factory_amount': None,
            'spread': down_payment,
            'cumulative': cumulative,
            'matched': False,
        })
    for idx, mk in enumerate(all_months, start=1):
        cust = cust_by_month.get(mk)
        fact = fact_by_month.get(mk)
        cust_amt = round(cust['amount'], 2) if cust else None
        fact_amt = round(fact['amount'], 2) if fact else None
        monthly = round((cust_amt or 0) - (fact_amt or 0), 2)
        cumulative = round(cumulative + monthly, 2)
        period_details.append({
            'period': idx,
            'month': mk,
            'customer_due_date': cust['due_date'] if cust else '',
            'customer_amount': cust_amt,
            'factory_due_date': fact['due_date'] if fact else '',
            'factory_amount': fact_amt,
            'spread': monthly,
            'cumulative': cumulative,
            'matched': bool(cust and fact),
        })
    spread_range_passed = bool(
        factory_rows
        and spread_total >= floor
        and (ceiling is None or spread_total <= ceiling)
    )
    status = '已通过' if spread_range_passed else '差异待处理'
    if not factory_rows:
        status = '未上传'

    summary = {
        'status': status,
        'customer_total': customer_total,
        'factory_total': factory_total,
        'spread_total': spread_total,
        'expected_profit_floor': floor,
        'expected_profit_ceiling': ceiling,
        'spread_range_passed': spread_range_passed,
        'customer_periods': customer_periods,
        'factory_periods': factory_periods,
        'customer_avg': round(customer_total / customer_periods, 2) if customer_periods else 0,
        'factory_avg': round(factory_total / factory_periods, 2) if factory_periods else 0,
        'avg_spread': round(
            (customer_total / customer_periods if customer_periods else 0)
            - (factory_total / factory_periods if factory_periods else 0),
            2
        ),
        'customer_first_due_date': customer_due_dates[0] if customer_due_dates else '',
        'customer_last_due_date': customer_due_dates[-1] if customer_due_dates else '',
        'factory_first_due_date': factory_due_dates[0] if factory_due_dates else '',
        'factory_last_due_date': factory_due_dates[-1] if factory_due_dates else '',
        'display_only_dimensions': ['period_count', 'first_due_date', 'last_due_date', 'average_period_spread'],
        'period_details': period_details,
    }
    summary_text = json.dumps(summary, ensure_ascii=False)
    c.execute("""
        UPDATE contracts
        SET customer_plan_match_status=?, plan_compare_summary=?
        WHERE id=?
    """, (status, summary_text, contract_id))
    if contract['sales_order_id']:
        c.execute("""
            UPDATE sales_orders
            SET customer_plan_match_status=?,
                factory_plan_match_status=?,
                plan_compare_summary=?
            WHERE id=?
        """, (status, '已上传' if factory_rows else '未上传', summary_text, contract['sales_order_id']))
    return summary


def build_customer_plan_summary(conn, contract_id):
    c = conn.cursor()
    c.execute("""
        SELECT COUNT(*) AS periods,
               COALESCE(SUM(amount), 0) AS total,
               COALESCE(AVG(amount), 0) AS avg_amount,
               MIN(due_date) AS first_due_date,
               MAX(due_date) AS last_due_date
        FROM repayments
        WHERE contract_id=?
          AND period>=1
    """, (contract_id,))
    row = c.fetchone()
    periods = int(row['periods'] or 0) if row else 0
    summary = {
        'status': '已生成' if periods else '未生成',
        'customer_periods': periods,
        'customer_total': round(parse_money(row['total'] if row else 0), 2),
        'customer_avg': round(parse_money(row['avg_amount'] if row else 0), 2),
        'customer_first_due_date': row['first_due_date'] if row else '',
        'customer_last_due_date': row['last_due_date'] if row else '',
        'factory_plan_required': False,
        'message': '客户还款计划已生成，财务可直接确认报单' if periods else '客户还款计划未生成',
    }
    summary_text = json.dumps(summary, ensure_ascii=False)
    c.execute("""
        UPDATE contracts
        SET customer_plan_match_status=?, plan_compare_summary=?
        WHERE id=?
    """, (summary['status'], summary_text, contract_id))
    c.execute("""
        UPDATE sales_orders
        SET customer_plan_match_status=?,
            factory_plan_match_status='无需上传',
            plan_compare_summary=?
        WHERE contract_id=?
    """, (summary['status'], summary_text, contract_id))
    return summary


def sales_order_plan_activation_blocker(conn, order_id):
    """财务确认报单前，分期类报单只需生成客户侧还款计划。"""
    c = conn.cursor()
    c.execute("SELECT id, sales_mode, contract_id FROM sales_orders WHERE id=?", (order_id,))
    order = c.fetchone()
    if not order:
        return '报单不存在'

    linked_contract_id = order['contract_id'] or -1
    c.execute("""
        SELECT id, contract_type, customer_plan_match_status
        FROM contracts
        WHERE sales_order_id=?
        ORDER BY CASE WHEN id=? THEN 0 ELSE 1 END, id DESC
        LIMIT 1
    """, (order_id, linked_contract_id))
    contract = c.fetchone()
    if not contract:
        try:
            ensure_sales_order_planning_contract(conn, order_id, {}, reset_factory=False)
        except ValueError as exc:
            return str(exc)
        c.execute("""
            SELECT id, contract_type, customer_plan_match_status
            FROM contracts
            WHERE sales_order_id=?
            ORDER BY id DESC
            LIMIT 1
        """, (order_id,))
        contract = c.fetchone()
        if not contract:
            return '客户还款计划未生成，不能财务确认报单'

    c.execute("SELECT COUNT(*) AS cnt FROM repayments WHERE contract_id=? AND period>=1", (contract['id'],))
    customer_count = c.fetchone()['cnt']
    if customer_count <= 0:
        return '客户还款计划未生成，不能财务确认报单'
    build_customer_plan_summary(conn, contract['id'])
    return None


def parse_target_periods(value):
    if value in (None, '', 'all', '全部'):
        return None
    if isinstance(value, list):
        return [int(v) for v in value]
    try:
        parsed = json.loads(value)
        if isinstance(parsed, list):
            return [int(v) for v in parsed]
        return [int(parsed)]
    except (TypeError, ValueError):
        return [int(value)]


def parse_period_count(value, default=12):
    text = str(value or '').strip()
    match = re.search(r'\d+', text)
    if match:
        return max(1, int(match.group(0)))
    return max(1, int(default or 12))


def clear_contract_plans(conn, contract_id, include_factory=True):
    c = conn.cursor()
    c.execute("DELETE FROM repayments WHERE contract_id=?", (contract_id,))
    if include_factory:
        c.execute("DELETE FROM factory_repayments WHERE contract_id=?", (contract_id,))


def generate_customer_repayment_plan(conn, contract_id, contract_type, start_date, loan_periods,
                                     rent, repayment_day=1, deposit=0, down_payment=0):
    c = conn.cursor()
    start_dt = datetime.strptime(start_date, '%Y-%m-%d')
    if contract_type == '租赁' and parse_money(deposit) > 0:
        c.execute("""
            INSERT INTO repayments (contract_id, period, due_date, amount, status, remark)
            VALUES (?, 0, NULL, ?, '未激活', '押金')
        """, (contract_id, parse_money(deposit)))
    if contract_type == '以租代售' and parse_money(down_payment) > 0:
        c.execute("""
            INSERT INTO repayments (contract_id, period, due_date, amount, status, remark)
            VALUES (?, 0, NULL, ?, '未激活', '首付款')
        """, (contract_id, parse_money(down_payment)))

    for p in range(1, int(loan_periods) + 1):
        try:
            if contract_type == '租赁':
                due_dt = start_dt + relativedelta(months=p - 1)
            else:
                due_dt = start_dt + relativedelta(months=p)
                due_dt = due_dt.replace(day=min(int(repayment_day or 1), 28))
        except Exception:
            due_dt = start_dt + timedelta(days=30 * (p - 1 if contract_type == '租赁' else p))
        c.execute("""
            INSERT INTO repayments (contract_id, period, due_date, amount, status)
            VALUES (?, ?, ?, ?, '未激活')
        """, (contract_id, p, due_dt.strftime('%Y-%m-%d'), parse_money(rent)))


def ensure_sales_order_planning_contract(conn, order_id, overrides=None, reset_factory=False):
    c = conn.cursor()
    c.execute("SELECT * FROM sales_orders WHERE id=?", (order_id,))
    order = c.fetchone()
    if not order:
        raise ValueError('报单不存在')
    if order['order_status'] in ('草稿', '已作废'):
        return None

    contract_type = contract_type_from_sales_mode(order['sales_mode'])

    data = overrides or {}
    vehicle_id = order['vehicle_id']
    customer_name = data.get('customer_name') or order['customer_name']
    customer_phone = data.get('customer_phone') or order['customer_phone']
    c.execute("SELECT car_type, purchase_price FROM vehicles WHERE id=?", (vehicle_id,))
    vrow = c.fetchone()
    snap_guidance = resolve_guidance_price_for_vehicle(conn, vrow)[0] if vrow else 0
    snap_invoice = vrow['purchase_price'] if vrow else 0

    customer_id = data.get('customer_id')
    if not customer_id and customer_name:
        c.execute("SELECT id FROM customers WHERE name=? AND COALESCE(phone, '')=COALESCE(?, '') ORDER BY id DESC LIMIT 1",
                  (customer_name, customer_phone or ''))
        existing_customer = c.fetchone()
        if existing_customer:
            customer_id = existing_customer['id']
        else:
            c.execute("INSERT INTO customers (name, phone) VALUES (?, ?)", (customer_name, customer_phone or ''))
            customer_id = c.lastrowid

    start_date = data.get('start_date') or order['payment_date'] or datetime.now().strftime('%Y-%m-%d')
    plan_snapshot = {}
    if contract_type == '以租代售' and order['snapshot_finance_plan']:
        try:
            plan_snapshot = json.loads(order['snapshot_finance_plan'])
        except (TypeError, ValueError):
            plan_snapshot = {}
    loan_periods = parse_period_count(
        data.get('loan_periods') or plan_snapshot.get('periods') or order['lease_term'],
        12,
    )
    try:
        repayment_day_default = min(datetime.strptime(normalize_date(start_date), '%Y-%m-%d').day, 28)
    except (TypeError, ValueError):
        repayment_day_default = 1
    repayment_day = int(parse_money(data.get('repayment_day'), repayment_day_default) or repayment_day_default)
    repayment_day = min(max(repayment_day, 1), 28)
    rent = parse_money(
        data.get('rent'),
        parse_money(plan_snapshot.get('period_price'), parse_money(order['vehicle_rent_amount'])),
    )
    order_payment_amount = parse_money(order['deposit_amount'])
    deposit = parse_money(data.get('deposit'), order_payment_amount if contract_type == '租赁' else 0)
    down_payment = parse_money(
        data.get('down_payment'),
        parse_money(
            plan_snapshot.get('down_payment'),
            order_payment_amount or parse_money(order['car_purchase_amount']),
        ) if contract_type == '以租代售' else 0,
    )
    monthly_payment = parse_money(data.get('monthly_payment'), 0)
    factory_periods = int(data.get('factory_periods') or loan_periods)
    factory_repayment_months = int(data.get('factory_repayment_months') or factory_periods)
    customer_loan_amount = parse_money(data.get('customer_loan_amount'), 0)
    loan_amount = parse_money(data.get('loan_amount'), 0)
    total_price = parse_money(data.get('total_price'), 0)  # sale_total_price 已废弃
    company = data.get('company') or order['receiving_company'] or ''
    yard = data.get('yard') or ''
    business_mode = data.get('business_mode') or order['sales_mode'] or contract_type
    rental_method = data.get('rental_method') or ('经营租赁' if contract_type == '租赁' else contract_type)
    expected_profit_floor = parse_money(data.get('expected_profit_floor'), 0)
    expected_profit_ceiling = normalize_expected_profit_ceiling(data.get('expected_profit_ceiling'))
    end_dt = datetime.strptime(start_date, '%Y-%m-%d') + timedelta(days=30 * loan_periods)

    c.execute("""
        SELECT id, customer_plan_match_status
        FROM contracts
        WHERE sales_order_id=?
        ORDER BY id DESC LIMIT 1
    """, (order_id,))
    existing = c.fetchone()
    if existing:
        contract_id = existing['id']
        existing_contract = c.execute("""
            SELECT contract_status, delivery_status, contract_file
            FROM contracts WHERE id=?
        """, (contract_id,)).fetchone()
        has_financial_activity = c.execute("""
            SELECT 1 FROM repayments
            WHERE contract_id=?
              AND (COALESCE(paid_amount,0)>0 OR COALESCE(verified_amount,0)>0)
            LIMIT 1
        """, (contract_id,)).fetchone()
        if (
            not existing_contract
            or existing_contract['contract_status'] != '报单计划中'
            or existing_contract['delivery_status'] != '待报单激活'
            or (existing_contract['contract_file'] or '').strip()
            or has_financial_activity
        ):
            raise ValueError('线下合同已上传、已出库或已有回款记录，不能重新生成客户还款计划')
        c.execute("""
            UPDATE contracts
            SET vehicle_id=?, customer_id=?, contract_type=?, business_mode=?, rental_method=?, repayment_day=?,
                start_date=?, end_date=?, total_price=?, customer_loan_amount=?, loan_amount=?, monthly_payment=?,
                rent=?, loan_periods=?, company=?, yard=?, lease_bank_name=?, lease_bank_card_no=?,
                factory_guarantee_deposit=?, factory_repayment_months=?, factory_periods=?, deposit=?, down_payment=?,
                snapshot_guidance_price=?, snapshot_invoice_price=?,
                expected_profit_floor=?, expected_profit_ceiling=?,
                customer_plan_match_status='已生成',
                plan_compare_summary=NULL
            WHERE id=?
        """, (
            vehicle_id, customer_id, contract_type, business_mode, rental_method, repayment_day,
            start_date, end_dt.strftime('%Y-%m-%d'), total_price, customer_loan_amount, loan_amount, monthly_payment,
            rent, loan_periods, company, yard, data.get('lease_bank_name', ''), data.get('lease_bank_card_no', ''),
            parse_money(data.get('factory_guarantee_deposit'), 0), factory_repayment_months, factory_periods,
            deposit, down_payment, snap_guidance, snap_invoice, expected_profit_floor, expected_profit_ceiling,
            contract_id,
        ))
        clear_contract_plans(conn, contract_id, include_factory=reset_factory)
    else:
        c.execute("""
            INSERT INTO contracts
                (vehicle_id, customer_id, sales_order_id, contract_type, business_mode, rental_method,
                 repayment_day, start_date, end_date, total_price, customer_loan_amount, loan_amount,
                 monthly_payment, rent, loan_periods, company, yard, lease_bank_name, lease_bank_card_no,
                 factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                 down_payment_status, deposit_status, delivery_status, contract_status,
                 snapshot_guidance_price, snapshot_invoice_price, expected_profit_floor, expected_profit_ceiling,
                 created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                    ?, ?, '待报单激活', '报单计划中', ?, ?, ?, ?, '系统')
        """, (
            vehicle_id, customer_id, order_id, contract_type, business_mode, rental_method,
            repayment_day, start_date, end_dt.strftime('%Y-%m-%d'), total_price, customer_loan_amount, loan_amount,
            monthly_payment, rent, loan_periods, company, yard, data.get('lease_bank_name', ''), data.get('lease_bank_card_no', ''),
            parse_money(data.get('factory_guarantee_deposit'), 0), factory_repayment_months, factory_periods,
            deposit, down_payment, '免收' if down_payment == 0 else '待收',
            '免收' if deposit == 0 else '待收', snap_guidance, snap_invoice,
            expected_profit_floor, expected_profit_ceiling,
        ))
        contract_id = c.lastrowid
        c.execute("UPDATE sales_orders SET contract_id=? WHERE id=?", (contract_id, order_id))

    generate_customer_repayment_plan(
        conn, contract_id, contract_type, start_date, loan_periods, rent,
        repayment_day=repayment_day, deposit=deposit, down_payment=down_payment
    )
    c.execute("""
        UPDATE sales_orders
        SET customer_plan_match_status='已生成',
            factory_plan_match_status='无需上传',
            plan_compare_summary=NULL,
            contract_id=?
        WHERE id=?
    """, (contract_id, order_id))
    build_customer_plan_summary(conn, contract_id)
    return contract_id


def ensure_default_sales_order_planning_if_needed(conn, order_id):
    c = conn.cursor()
    c.execute("SELECT sales_mode, order_status FROM sales_orders WHERE id=?", (order_id,))
    order = c.fetchone()
    if not order:
        return None
    if order['order_status'] in ('草稿', '已作废', '待价格特批'):
        return None
    return ensure_sales_order_planning_contract(conn, order_id, {}, reset_factory=False)


def login_required(f):
    """登录校验装饰器"""
    @wraps(f)
    def decorated(*args, **kwargs):
        user = get_current_user()
        if not user:
            return jsonify({'success': False, 'message': '请先登录', 'code': 401}), 401
        request.current_user = user
        return f(*args, **kwargs)
    return decorated


def require_role(*roles):
    """角色权限校验装饰器"""
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            user = get_current_user()
            if not user:
                return jsonify({'success': False, 'message': '请先登录', 'code': 401}), 401
            if user['role'] not in roles and '老板' != user['role']:
                return jsonify({'success': False, 'message': '无权限执行此操作', 'code': 403}), 403
            request.current_user = user
            return f(*args, **kwargs)
        return decorated
    return decorator


# ======================== 登录/登出 ========================
@app.route('/api/auth/login', methods=['POST'])
def login():
    data = request.json
    username = data.get('username', '')
    password = data.get('password', '')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username=? AND password=? AND is_active=1", (username, password))
    user = c.fetchone()
    if not user:
        conn.close()
        return jsonify({'success': False, 'message': '用户名或密码错误'}), 401
    pages = role_pages_for(conn, user['role'])
    actions = role_actions_for(conn, user['role'])
    hidden_fields = role_hidden_fields_for(conn, user['role'])
    conn.close()
    session['user_id'] = user['id']
    return jsonify({
        'success': True,
        'user': {
            'id': user['id'],
            'username': user['username'],
            'display_name': user['display_name'],
            'role': user['role'],
            'pages': pages,
            'actions': actions,
            'hidden_fields': hidden_fields,
        }
    })


@app.route('/api/auth/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({'success': True})


@app.route('/api/auth/me', methods=['GET'])
def get_me():
    user = get_current_user()
    if not user:
        return jsonify({'success': False, 'code': 401}), 401
    conn = get_db()
    pages = role_pages_for(conn, user['role'])
    actions = role_actions_for(conn, user['role'])
    hidden_fields = role_hidden_fields_for(conn, user['role'])
    conn.close()
    return jsonify({
        'success': True,
        'user': {
            **user,
            'pages': pages,
            'actions': actions,
            'hidden_fields': hidden_fields,
        }
    })


@app.route('/api/role-permissions', methods=['GET', 'POST'])
@require_role('老板')
def role_permissions():
    conn = get_db()
    c = conn.cursor()
    if request.method == 'GET':
        roles = sorted(set(list(ROLE_PAGES.keys()) + list(ROLE_ACTIONS.keys())))
        data = {}
        for role in roles:
            data[role] = {
                'pages': role_pages_for(conn, role),
                'actions': role_actions_for(conn, role),
                'hidden_fields': role_hidden_fields_for(conn, role),
            }
        conn.close()
        return jsonify({'success': True, 'roles': data})

    payload = request.json or {}
    role = payload.get('role')
    if role not in ROLE_PAGES:
        conn.close()
        return jsonify({'success': False, 'message': '角色不正确'}), 400
    pages = payload.get('pages')
    actions = payload.get('actions')
    hidden_fields = payload.get('hidden_fields')
    if pages is not None:
        c.execute("DELETE FROM role_pages WHERE role=?", (role,))
        for page in pages:
            c.execute("INSERT OR IGNORE INTO role_pages (role, page_key) VALUES (?, ?)", (role, page))
    if actions is not None:
        c.execute("DELETE FROM role_actions WHERE role=?", (role,))
        for action in actions:
            c.execute("INSERT OR IGNORE INTO role_actions (role, action_key) VALUES (?, ?)", (role, action))
    if hidden_fields is not None:
        c.execute("DELETE FROM role_field_permissions WHERE role=?", (role,))
        for resource, fields in hidden_fields.items():
            for field in fields:
                c.execute("""
                    INSERT OR IGNORE INTO role_field_permissions
                        (role, resource_key, field_key, can_view)
                    VALUES (?, ?, ?, 0)
                """, (role, resource, field))
    log_audit(conn, '字段级权限变更', 'role_permissions', None,
              f'角色:{role}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '权限矩阵已更新'})



# ======================== 审计日志 ========================
def log_audit(conn, action, target_type, target_id, detail, operator='系统'):
    """PRD NFR: 所有财务状态变更必须记录操作人+原始金额+变更后金额"""
    ip = request.remote_addr if request else '127.0.0.1'
    conn.execute(
        "INSERT INTO audit_logs (action, target_type, target_id, detail, operator, ip_address) VALUES (?,?,?,?,?,?)",
        (action, target_type, target_id, detail, operator, ip)
    )


def reconciliation_gate_for_repayment(conn, repayment_id):
    """后续对账门禁：必须先完成线下合同、首次付款、出库。"""
    row = conn.execute("""
        SELECT r.*,
               c.contract_file,
               c.delivery_status,
               c.contract_status,
               c.sales_order_id,
               so.order_status AS sales_order_status
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        LEFT JOIN sales_orders so ON so.id = c.sales_order_id
        WHERE r.id=?
    """, (repayment_id,)).fetchone()
    if not row:
        return None, '记录不存在', 404
    if int(row['period'] or 0) <= 0:
        return row, '首付款/押金必须走“合同首次付款”流程，不能在后续对账单里核销', 400
    if row['status'] == '已还款':
        return row, '该期账单已完成对账，不能重复对账', 400
    if row['status'] == '已取消':
        return row, '该期账单已因退车结算取消，不能再对账', 400
    # “预抵”为部分抵扣（不足一整期），允许继续对账补足差额，不视为已完成
    if row['status'] == '未激活':
        return row, '还款计划尚未激活，请先完成合同首付款审核和出库流程', 400
    if row['sales_order_id'] and row['sales_order_status'] != '已激活':
        return row, '请先完成销售报单的财务确认，再进入合同和对账流程', 400
    if not (row['contract_file'] or '').strip():
        return row, '请先由运营上传线下合同，合同未上传不能对账', 400
    if row['delivery_status'] != '已出库':
        status = row['delivery_status'] or row['contract_status'] or '未知'
        return row, f'当前合同状态为「{status}」，需完成首付款审核并车辆出库后才能对账', 400
    return row, '', 200


# ======================== 自动逾期检测 + T+7锁车 ========================
# ======================== H1 日终批处理（逐日逾期N日状态机 + 滞纳金计提）========================
LATE_FEE_DAILY_RATE = 0.0005  # 万分之5/日，无复利

def run_daily_collect(force=False):
    """H1 统一入口：每日 00:30 / 惰性兜底调用。
    逐日重算 diff，推进 临近还款/还款日/逾期N日；逾期按万5/日写 late_fee_ledger。
    (job='daily-collect', run_date) 当天幂等，漏跑可按 diff 补推到正确档位。"""
    conn = get_db()
    c = conn.cursor()
    today_str = datetime.now().strftime('%Y-%m-%d')
    if not force:
        c.execute("SELECT summary FROM daily_job_runs WHERE job='daily-collect' AND run_date=?", (today_str,))
        cached = c.fetchone()
        if cached:
            conn.close()
            return {'skipped': True, 'run_date': today_str}
    today = datetime.strptime(today_str, '%Y-%m-%d').date()
    expire_sales_order_drafts(conn)
    stats = {'near': 0, 'due': 0, 'overdue': 0, 'late_fee': 0, 'urge_op': 0, 'urge_sales': 0, 't3_reminder': 0}
    contract_late = {}  # contract_id -> 累计滞纳金
    c.execute("""
        SELECT r.id, r.contract_id, r.period, r.due_date, r.amount,
               COALESCE(r.paid_amount,0) AS paid_amount, r.status,
               c.vehicle_id
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        WHERE (r.status IN ('待还款','临近还款','还款日','部分核销','预抵','逾期')
               OR r.status LIKE '逾期%日')
          AND r.period >= 1
          AND COALESCE(c.contract_file, '')!=''
          AND c.delivery_status='已出库'
    """)
    for r in c.fetchall():
        _process_repayment_row(c, r, today, today_str, stats, contract_late)
    # 补生成提醒文案：状态已是「临近还款」但历史漏跑未写提醒的期（幂等）。
    c.execute("""
        SELECT r.id, r.contract_id, r.due_date
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        WHERE r.status='临近还款'
          AND r.period >= 1
          AND c.delivery_status='已出库'
    """)
    for r in c.fetchall():
        if ensure_t3_reminder(c, r, today_str):
            stats['t3_reminder'] += 1
    c.execute("""
        SELECT id, contract_id, repayment_id, initial_payment_id, receivable_type,
               source_period, amount, paid_amount, due_date, promised_repay_date, status
        FROM receivables
        WHERE status NOT IN ('已结清', '已取消')
    """)
    for rec in c.fetchall():
        _process_receivable_row(c, rec, today, today_str, stats, contract_late)
    # 同步合同级 late_fee 应收
    for cid in contract_late:
        _sync_late_fee_item(c, cid)
    # 厂家还款仍按简单逾期标记
    c.execute("UPDATE factory_repayments SET status='逾期' WHERE status='待还款' AND due_date < ?", (today_str,))
    summary = (f"near={stats['near']} due={stats['due']} overdue={stats['overdue']} "
               f"late_fee_rows={stats['late_fee']} urge_op={stats['urge_op']} urge_sales={stats['urge_sales']} "
               f"t3_reminder={stats['t3_reminder']}")
    c.execute("INSERT OR REPLACE INTO daily_job_runs (job, run_date, ran_at, summary) "
              "VALUES ('daily-collect', ?, datetime('now','localtime'), ?)", (today_str, summary))
    conn.commit()
    conn.close()
    return {'skipped': False, 'run_date': today_str, 'stats': stats}


def check_overdue():
    """惰性兜底：看板/账单/风控等查询入口调用，总是重算当前状态。
    重算幂等（状态按 diff 重推；late_fee_ledger 按 (repayment_id,accrued_date) 唯一去重），
    故用 force=True 保证视图实时反映最新到期/逾期，不受当天调度幂等门控影响。"""
    run_daily_collect(force=True)


def ensure_urge_task(c, repayment, urge_day):
    """按还款期和节点幂等创建待执行催收任务。"""
    urge_type = '运营催款' if urge_day == 3 else '销售催款'
    c.execute("""
        INSERT OR IGNORE INTO urge_records
            (repayment_id, contract_id, vehicle_id, urge_type, urge_day, status, result, remark)
        VALUES (?, ?, ?, ?, ?, '待执行', '待跟进', ?)
    """, (
        repayment['id'],
        repayment['contract_id'],
        repayment['vehicle_id'],
        urge_type,
        urge_day,
        f'系统于逾期 T+{urge_day} 创建催收任务',
    ))


def _t3_reminder_text(customer_name, plate_number, vin, due_date):
    """流程图 4.1 T-3 提醒文案：尊敬的xx，您的车架号为xx的车辆将于xx日进行还款，请您知悉。"""
    return f"尊敬的{customer_name or '客户'}，您的车架号为{vin or '-'}的车辆将于{due_date}进行还款，请您知悉"


def ensure_t3_reminder(c, repayment, today_str):
    """T-3 临近还款提醒：按还款期幂等生成提醒文案（写审计日志，界面从还款单状态呈现黄色预警）。"""
    rid = repayment['id']
    exists = c.execute(
        "SELECT 1 FROM audit_logs WHERE action='T-3还款提醒' AND target_type='repayment' AND target_id=? LIMIT 1",
        (rid,),
    ).fetchone()
    if exists:
        return False
    row = c.execute("""
        SELECT v.plate_number, v.vin, cu.name AS customer_name
        FROM contracts ct
        JOIN vehicles v ON v.id = ct.vehicle_id
        LEFT JOIN customers cu ON cu.id = ct.customer_id
        WHERE ct.id=?
    """, (repayment['contract_id'],)).fetchone()
    text = _t3_reminder_text(
        row['customer_name'] if row else '',
        row['plate_number'] if row else '',
        row['vin'] if row else '',
        repayment['due_date'],
    )
    c.execute(
        "INSERT INTO audit_logs (action, target_type, target_id, detail, operator) VALUES (?, ?, ?, ?, ?)",
        ('T-3还款提醒', 'repayment', rid, text, '系统'),
    )
    return True


def _process_repayment_row(c, r, today, today_str, stats, contract_late):
    """对单笔 repayment 按 diff 推进状态 + 计提滞纳金。c 为共享游标。"""
    rid = r['id']
    cid = r['contract_id']
    status = r['status']
    amount = parse_money(r['amount'])
    paid = parse_money(r['paid_amount'])
    settled = paid >= amount and amount > 0
    if not r['due_date']:
        return
    try:
        due = datetime.strptime(r['due_date'], '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return
    diff = (today - due).days

    # 已结清：仅复核，不回退、不计滞纳金
    if status in ('已还款', '预抵') and settled:
        return

    new_status = status
    if diff < -3:
        new_status = '待还款' if status in ('临近还款',) else status
    elif -3 <= diff <= -1:
        if status in ('待还款',):
            new_status = '临近还款'
            stats['near'] += 1
            if ensure_t3_reminder(c, r, today_str):
                stats['t3_reminder'] += 1
    elif diff == 0:
        if status in ('待还款', '临近还款', '部分核销'):
            new_status = '还款日'
            stats['due'] += 1
    else:  # diff >= 1 且未结清 → 逾期N日
        if not settled:
            new_status = f'逾期{diff}日'
            stats['overdue'] += 1
            cum = _accrue_late_fee(c, rid, cid, r['period'], amount, paid, due, today, today_str)
            if cum is not None:
                contract_late[cid] = contract_late.get(cid, 0) + cum
                stats['late_fee'] += 1
            # 日终任务可能首次就在 T+3/T+7 之后执行，需补齐所有已到期节点。
            if diff >= 3:
                ensure_urge_task(c, r, 3)
                stats['urge_op'] += 1
            if diff >= 7:
                ensure_urge_task(c, r, 7)
                stats['urge_sales'] += 1

    if new_status != status:
        c.execute("UPDATE repayments SET status=? WHERE id=?", (new_status, rid))


def _process_receivable_row(c, rec, today, today_str, stats, contract_late):
    amount = parse_money(rec['amount'])
    paid = parse_money(rec['paid_amount'])
    outstanding = round(amount - paid, 2)
    if outstanding <= 0:
        c.execute("UPDATE receivables SET status='已结清', settled_at=COALESCE(settled_at, datetime('now','localtime')) WHERE id=?", (rec['id'],))
        return
    due_text = rec['promised_repay_date'] or rec['due_date']
    if not due_text:
        return
    try:
        due = datetime.strptime(due_text, '%Y-%m-%d').date()
    except (TypeError, ValueError):
        return
    diff = (today - due).days
    if rec['receivable_type'] == 'initial_payment_shortfall' and diff <= 0:
        c.execute("UPDATE receivables SET status='待归还' WHERE id=?", (rec['id'],))
        return
    if diff >= 1:
        daily = round(outstanding * LATE_FEE_DAILY_RATE, 2)
        accrued = round(daily * diff, 2)
        c.execute("""
            UPDATE receivables
            SET status='逾期应收', late_fee_accrued=?
            WHERE id=?
        """, (accrued, rec['id']))
        contract_late[rec['contract_id']] = contract_late.get(rec['contract_id'], 0) + accrued
        stats['overdue'] += 1
        stats['late_fee'] += 1


def _accrue_late_fee(c, rid, cid, period, amount, paid, due, today, today_str):
    """逐日计提滞纳金写 late_fee_ledger，(repayment_id, accrued_date) 幂等。
    基数 = amount - paid（原始未还本金，押金/首付款不抵扣），万5/日无复利。
    支持漏跑补提：从逾期首日(due+1)到今天，缺哪天补哪天。返回最新累计额。"""
    outstanding = amount - paid
    if outstanding <= 0:
        return None
    # 当前已计提到哪天 + 累计额
    c.execute("SELECT accrued_date, cumulative_amount FROM late_fee_ledger "
              "WHERE repayment_id=? ORDER BY accrued_date DESC LIMIT 1", (rid,))
    last = c.fetchone()
    cumulative = parse_money(last['cumulative_amount']) if last else 0.0
    last_date = None
    if last and last['accrued_date']:
        try:
            last_date = datetime.strptime(last['accrued_date'], '%Y-%m-%d').date()
        except (TypeError, ValueError):
            last_date = None
    first_overdue = due + timedelta(days=1)
    cursor_day = first_overdue if last_date is None else last_date + timedelta(days=1)
    daily = round(outstanding * LATE_FEE_DAILY_RATE, 2)
    while cursor_day <= today:
        cumulative = round(cumulative + daily, 2)
        c.execute("INSERT OR IGNORE INTO late_fee_ledger "
                  "(repayment_id, contract_id, period, accrued_date, outstanding, daily_amount, cumulative_amount) "
                  "VALUES (?,?,?,?,?,?,?)",
                  (rid, cid, period, cursor_day.strftime('%Y-%m-%d'), outstanding, daily, cumulative))
        cursor_day += timedelta(days=1)
    return cumulative


def _sync_late_fee_item(c, cid):
    """按全量台账重算合同滞纳金应收，避免跨期累计被当前期覆盖。"""
    c.execute("""
        SELECT COALESCE(SUM(cumulative_amount), 0) AS gross
        FROM (
            SELECT repayment_id, MAX(cumulative_amount) AS cumulative_amount
            FROM late_fee_ledger
            WHERE contract_id=?
            GROUP BY repayment_id
        )
    """, (cid,))
    gross = parse_money(c.fetchone()['gross'])
    c.execute("SELECT COALESCE(SUM(waived_amount),0) AS w FROM late_fee_ledger WHERE contract_id=? AND waived=1", (cid,))
    waived = parse_money(c.fetchone()['w'])
    net = round(max(0.0, gross - waived), 2)
    c.execute("SELECT id FROM contract_fee_items WHERE contract_id=? AND fee_type='late_fee' LIMIT 1", (cid,))
    row = c.fetchone()
    if row:
        c.execute("UPDATE contract_fee_items SET amount_due=? WHERE id=?", (net, row['id']))
        sync_fee_item_status(c.connection, row['id'])
    else:
        c.execute("INSERT INTO contract_fee_items (contract_id, fee_type, description, amount_due, amount_paid, status, created_by) "
                  "VALUES (?, 'late_fee', '逾期滞纳金(系统计提)', ?, 0, '待支付', '系统')", (cid, net))


def reverse_late_fee_accruals_from_payment_date(conn, repayment_id, payment_date, reason):
    """冲回收款日及之后、因延迟同步付款状态产生的错误滞纳金。"""
    if not payment_date:
        return 0
    try:
        payment_day = datetime.strptime(payment_date[:10], '%Y-%m-%d').strftime('%Y-%m-%d')
    except (TypeError, ValueError):
        return 0

    c = conn.cursor()
    row = c.execute("""
        SELECT contract_id
        FROM late_fee_ledger
        WHERE repayment_id=?
        LIMIT 1
    """, (repayment_id,)).fetchone()
    if not row:
        return 0
    c.execute("""
        UPDATE late_fee_ledger
        SET waived=1,
            waived_amount=daily_amount
        WHERE repayment_id=?
          AND accrued_date>=?
          AND COALESCE(waived, 0)=0
    """, (repayment_id, payment_day))
    reversed_count = c.rowcount
    if reversed_count:
        _sync_late_fee_item(c, row['contract_id'])
        log_audit(
            conn,
            '冲回错误滞纳金',
            'repayment',
            repayment_id,
            f'{reason}；自{payment_day}起冲回{reversed_count}日系统误计滞纳金',
            '系统',
        )
    return reversed_count


# ======================== 页面路由 ========================
@app.route('/')
def index():
    return send_from_directory('templates', 'index.html')


# ======================== 仪表盘统计 ========================
def _fetch_scalar(c, sql, params=()):
    c.execute(sql, params)
    row = c.fetchone()
    if row is None:
        return 0
    # sqlite3: row[0] 可用；MySQL DictCursor: row 是 dict
    if isinstance(row, dict):
        values = list(row.values())
        return values[0] if values else 0
    return row[0]


def _money_expr(expr="amount", paid_expr="paid_amount"):
    return f"CASE WHEN COALESCE({expr},0) > COALESCE({paid_expr},0) THEN COALESCE({expr},0) - COALESCE({paid_expr},0) ELSE 0 END"


def build_dashboard_metrics(conn):
    """统一仪表盘统计口径，普通仪表盘与老板穿透看板共用。"""
    c = conn.cursor()
    today = datetime.now().date()
    today_str = today.strftime('%Y-%m-%d')
    month_start = today.replace(day=1)
    next_month = month_start + relativedelta(months=1)
    month_start_str = month_start.strftime('%Y-%m-%d')
    next_month_str = next_month.strftime('%Y-%m-%d')

    open_order_where = """
        order_status IN ('待价格特批','待财务确认')
        OR (order_status='已激活' AND contract_id IS NULL)
    """
    active_bill_filter = """
        r.period >= 1
        AND COALESCE(c.contract_file, '')!=''
        AND c.delivery_status='已出库'
    """

    total_vehicles = _fetch_scalar(c, "SELECT COUNT(*) FROM vehicles WHERE (is_deleted IS NULL OR is_deleted = 0)")
    active_vehicles = _fetch_scalar(
        c,
        "SELECT COUNT(*) FROM vehicles WHERE COALESCE(status,'') NOT IN ('已过户') AND (is_deleted IS NULL OR is_deleted = 0)"
    )
    active_contract_count = _fetch_scalar(
        c,
        "SELECT COUNT(*) FROM contracts WHERE contract_status='执行中'"
    )
    open_order_count = _fetch_scalar(c, f"SELECT COUNT(*) FROM sales_orders WHERE {open_order_where}")

    total_invoice = parse_money(_fetch_scalar(c, "SELECT COALESCE(SUM(purchase_price),0) FROM vehicles WHERE (is_deleted IS NULL OR is_deleted = 0)"))
    total_residual = parse_money(_fetch_scalar(c, "SELECT COALESCE(SUM(estimated_residual_value),0) FROM vehicles WHERE (is_deleted IS NULL OR is_deleted = 0)"))
    c.execute("SELECT COALESCE(SUM(loan_amount),0) AS v1, COALESCE(SUM(paid_principal),0) AS v2 FROM contracts")
    r = c.fetchone()
    total_loan = parse_money(r['v1'] if isinstance(r, dict) else r[0])
    total_paid_principal = parse_money(r['v2'] if isinstance(r, dict) else r[1])
    c.execute("SELECT COALESCE(SUM(collected_rent),0) AS v1, COALESCE(SUM(collected_deposit),0) AS v2 FROM contracts")
    r = c.fetchone()
    total_rent = parse_money(r['v1'] if isinstance(r, dict) else r[0])
    total_deposit = parse_money(r['v1'] if isinstance(r, dict) else r[1])

    overdue_count = _fetch_scalar(c, f"""
        SELECT COUNT(*)
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        WHERE ({active_bill_filter})
          AND (r.status LIKE '逾期%' OR r.status='部分核销')
    """)
    factory_overdue_count = _fetch_scalar(c, """
        SELECT COUNT(*)
        FROM factory_repayments fr
        JOIN contracts c ON c.id = fr.contract_id
        WHERE fr.status='逾期'
          AND COALESCE(c.contract_file, '')!=''
    """)

    monthly_due = parse_money(_fetch_scalar(c, f"""
        SELECT COALESCE(SUM({_money_expr('r.amount', 'r.paid_amount')}),0)
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        WHERE ({active_bill_filter})
          AND r.status NOT IN ('已还款','预抵','未激活')
          AND r.due_date >= ?
          AND r.due_date < ?
    """, (month_start_str, next_month_str)))
    monthly_factory_due = parse_money(_fetch_scalar(c, """
        SELECT COALESCE(SUM(fr.amount),0)
        FROM factory_repayments fr
        JOIN contracts c ON c.id = fr.contract_id
        WHERE fr.status!='已还款'
          AND COALESCE(c.contract_file, '')!=''
          AND fr.due_date >= ?
          AND fr.due_date < ?
    """, (month_start_str, next_month_str)))

    total_customer_received = parse_money(_fetch_scalar(c, """
        SELECT COALESCE(SUM(
            CASE
                WHEN COALESCE(paid_amount,0)>0 THEN COALESCE(paid_amount,0)
                WHEN status='已还款' THEN COALESCE(amount,0)
                ELSE 0
            END
        ),0)
        FROM repayments
        WHERE period>=1
    """))
    total_factory_paid = parse_money(_fetch_scalar(
        c,
        "SELECT COALESCE(SUM(amount),0) FROM factory_repayments WHERE status='已还款'"
    ))
    rebate_total = parse_money(_fetch_scalar(c, "SELECT COALESCE(SUM(rebate_amount),0) FROM vehicle_rebates"))
    gross_profit = round(total_customer_received - total_factory_paid, 2)

    expiring_insurance_count = _fetch_scalar(c, """
        SELECT COUNT(*)
        FROM vehicles
        WHERE insurance_expiry_date IS NOT NULL
          AND insurance_expiry_date != ''
          AND date(insurance_expiry_date) <= date(?, '+30 day')
          AND COALESCE(status,'') NOT IN ('已过户')
          AND (is_deleted IS NULL OR is_deleted = 0)
    """, (today_str,))

    c.execute("""
        SELECT COALESCE(NULLIF(status, ''), '未知') AS label, COUNT(*) AS count
        FROM vehicles
        WHERE (is_deleted IS NULL OR is_deleted = 0)
        GROUP BY COALESCE(NULLIF(status, ''), '未知')
        ORDER BY count DESC, label ASC
    """)
    status_distribution = [dict(row) for row in c.fetchall()]

    c.execute("SELECT created_at, purchase_price FROM vehicles WHERE (is_deleted IS NULL OR is_deleted = 0)")
    vehicle_rows = [dict(row) for row in c.fetchall()]
    c.execute("""
        SELECT paid_at,
               verified_at,
               CASE
                   WHEN COALESCE(paid_amount,0)>0 THEN COALESCE(paid_amount,0)
                   WHEN status='已还款' THEN COALESCE(amount,0)
                   ELSE 0
               END AS received_amount
        FROM repayments
        WHERE period>=1
          AND (COALESCE(paid_amount,0)>0 OR status='已还款')
    """)
    customer_receipts = [dict(row) for row in c.fetchall()]
    c.execute("""
        SELECT approved_at, received_amount, amount
        FROM contract_initial_payments
        WHERE status='已通过'
    """)
    initial_receipts = [dict(row) for row in c.fetchall()]

    months = []
    asset_values = []
    income_values = []
    def _chart_date(value, fallback=today_str):
        clean = normalize_date(value)
        return clean if clean else fallback

    for month in range(1, 13):
        start = datetime(today.year, month, 1).date()
        end = start + relativedelta(months=1) - timedelta(days=1)
        months.append(f"{month}月")
        asset_values.append(round(sum(
            parse_money(v.get('purchase_price'))
            for v in vehicle_rows
            if _chart_date(v.get('created_at')) <= end.strftime('%Y-%m-%d')
        ), 2))
        customer_income = sum(
            parse_money(r.get('received_amount'))
            for r in customer_receipts
            if _chart_date(r.get('paid_at') or r.get('verified_at')) <= end.strftime('%Y-%m-%d')
        )
        initial_income = sum(
            parse_money(r.get('received_amount'), parse_money(r.get('amount')))
            for r in initial_receipts
            if _chart_date(r.get('approved_at')) <= end.strftime('%Y-%m-%d')
        )
        income_values.append(round(customer_income + initial_income, 2))

    pending_invoice_count = _fetch_scalar(c, "SELECT COUNT(*) FROM invoice_requests WHERE status='待开票'")
    pending_waiver_count = _fetch_scalar(c, "SELECT COUNT(*) FROM waivers WHERE status IN ('待审批','已通过')")
    pending_return_count = _fetch_scalar(c, "SELECT COUNT(*) FROM return_inspections WHERE status NOT IN ('已完成','已入库')")
    pending_lock_count = _fetch_scalar(c, "SELECT COUNT(*) FROM lock_requests WHERE status IN ('待运营审核','待老板审批','已批准待执行')")

    return {
        'total_vehicles': total_vehicles,
        'vehicle_count': total_vehicles,
        'active_vehicles': active_vehicles,
        'active_contract_count': active_contract_count,
        'open_order_count': open_order_count,
        'active_order_contract_count': active_contract_count + open_order_count,
        'total_invoice': round(total_invoice, 2),
        'total_residual': round(total_residual, 2),
        'total_loan': round(parse_money(total_loan), 2),
        'total_paid_principal': round(parse_money(total_paid_principal), 2),
        'total_rent': round(parse_money(total_rent), 2),
        'total_deposit': round(parse_money(total_deposit), 2),
        'overdue_count': overdue_count,
        'overdue_repayment_count': overdue_count,
        'factory_overdue_count': factory_overdue_count,
        'monthly_due': round(monthly_due, 2),
        'monthly_factory_due': round(monthly_factory_due, 2),
        'total_customer_received': round(total_customer_received, 2),
        'customer_received': round(total_customer_received, 2),
        'total_factory_paid': round(total_factory_paid, 2),
        'factory_paid': round(total_factory_paid, 2),
        'rebate_total': round(rebate_total, 2),
        'gross_profit': gross_profit,
        'realized_cash_profit': round(gross_profit + rebate_total, 2),
        'expiring_insurance_count': expiring_insurance_count,
        'vehicle_status_distribution': status_distribution,
        'asset_chart': {
            'months': months,
            'asset_values': asset_values,
            'income_values': income_values,
        },
        'pending_invoice_count': pending_invoice_count,
        'pending_waiver_count': pending_waiver_count,
        'pending_return_count': pending_return_count,
        'pending_lock_count': pending_lock_count,
        'data_as_of': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }


@app.route('/api/dashboard/stats', methods=['GET'])
@login_required
def get_stats():
    # 每次查看仪表盘时触发逾期检测
    check_overdue()

    conn = get_db()
    metrics = build_dashboard_metrics(conn)
    conn.close()
    return jsonify(metrics)


# ======================== 车辆资产 CRUD ========================
@app.route('/api/vehicles', methods=['GET'])
@login_required
def get_vehicles():
    user = request.current_user
    page = request.args.get('page', 1, type=int)
    page_size = request.args.get('page_size', 20, type=int)
    page_size = max(1, min(page_size, 100))
    show_deleted = request.args.get('show_deleted', 0, type=int)

    # 仅老板可以查看已删除车辆
    if show_deleted and user['role'] != '老板':
        show_deleted = 0

    conn = get_db()
    c = conn.cursor()

    where_clause = "" if show_deleted else "WHERE COALESCE(is_deleted,0)=0"
    c.execute(f"SELECT COUNT(*) AS cnt FROM vehicles {where_clause}")
    total = c.fetchone()['cnt']

    offset = (page - 1) * page_size
    c.execute(f"""
        SELECT v.*, c.rental_method, c.business_mode, c.loan_amount, c.monthly_payment, c.rent,
               c.loan_periods, c.deposit, c.paid_principal, c.loan_balance,
               c.collected_deposit, c.collected_rent, c.contract_status,
               cu.name as customer_name
        FROM vehicles v
        LEFT JOIN contracts c ON c.vehicle_id = v.id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        {where_clause.replace('COALESCE(is_deleted,0)=0', 'COALESCE(v.is_deleted, 0) = 0') if not show_deleted else ''}
        ORDER BY v.id ASC
        LIMIT ? OFFSET ?
    """, (page_size, offset))
    vehicles = [dict(row) for row in c.fetchall()]
    vehicles = redact_for_role(conn, user['role'], 'vehicles', vehicles)
    conn.close()
    return jsonify({'total': total, 'page': page, 'page_size': page_size, 'data': vehicles})


@app.route('/api/vehicles/list', methods=['GET'])
@login_required
def get_vehicles_list():
    """精简车辆列表（供筛选/车型统计用；含成色/能源/规格维度，筛选生效时前端基于此全量过滤）"""
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT v.id, v.vin, v.car_type, v.plate_number, v.status,
               v.is_deleted, v.box_type, v.condition, v.fuel_form,
               v.battery_brand, v.battery_capacity, v.battery_model,
               v.horsepower, v.gear_position
        FROM vehicles v
        ORDER BY v.id ASC
    """)
    rows = [dict(row) for row in c.fetchall()]
    # 历史导入记录可能没有燃料形式；统一返回后端识别的能源分类，
    # 让销售端筛选与后端车型口径保持一致。
    for row in rows:
        row['energy_type'] = energy_type_for_vehicle(row)
    rows = redact_for_role(conn, user['role'], 'vehicles', rows)
    conn.close()
    return jsonify(rows)


# 旧硬编码车型预设（已废弃：/api/car-types 改为库存聚合，保留仅为种子数据参考）
CAR_TYPE_PRESETS = [
    # 纯电
    {'label': '解放轻卡4米2-虎6G140度纯电-宁德电池', 'category': '纯电'},
    {'label': '解放轻卡4米2-虎6G120度纯电-宁德电池', 'category': '纯电'},
    {'label': '解放轻卡-虎VR纯电-轻盈版',           'category': '纯电'},
    # 混动
    {'label': '解放轻卡4米2-虎6G 180混动-盟固利电池', 'category': '混动'},
    # 油车
    {'label': '解放轻卡4米2-领途190马力',            'category': '油车'},
    {'label': '解放轻卡4米2-领途150马力',            'category': '油车'},
    {'label': '解放轻卡3米8-云内150排半',            'category': '油车'},
]

# 报单端静态配置（20260810 SKU 改造：已入库 data_dictionaries，以下保留为种子数据来源与兜底）
SALES_ORDER_CAR_OPTION_CONFIG = [
    {
        'category': '纯电',
        'cab_options': ['虎六G宽体', '虎六G中体', '虎VR地库'],
        'engine_battery_options': ['宁德时代', '弗迪', '力神'],
        'power_battery_options': ['81度电', '100度电', '120度电', '132度电', '140度电'],
        'gearbox_options': ['两档', '四挡', '五档', '六档', '8档', '10档', '65KW', '95KW'],
        'color_options': ['白色', '红色', '蓝色', '金色', '灰色', '黄色'],
        'box_options': ['厢货', '高栏', '平板', '冷藏'],
        'box_remark_placeholder': '①厢尺寸 X米*X米\n②侧门 双开/三开/侧全开/后开\n③带尾板/不带尾板\n④特殊要求',
    },
    {
        'category': '混动',
        'cab_options': ['虎VR双排', '虎VR单排'],
        'engine_battery_options': ['锡柴'],
        'power_battery_options': ['140马力', '150马力', '180马力'],
        'gearbox_options': [],
        'color_options': [],
        'box_options': [],
        'box_remark_placeholder': '',
    },
    {
        'category': '燃油车',
        'cab_options': ['虎六G宽体', '虎六G中体', '领途宽体', '领途中体', '招财虎'],
        'engine_battery_options': ['潍柴', '锡柴', '全柴', '云内'],
        'power_battery_options': ['150马力', '150排半', '160马力', '170马力', '180马力', '190马力', 'YN30'],
        'gearbox_options': [],
        'color_options': [],
        'box_options': [],
        'box_remark_placeholder': '',
    },
]

SALES_ORDER_COMMON_OPTIONS = {
    'gearbox_options': ['两档', '四挡', '五档', '六档', '8档', '10档'],
    'color_options': ['白色', '红色', '蓝色', '金色', '灰色', '黄色', '黑色', '银色'],
    'box_options': ['厢货', '高栏', '平板', '冷藏'],
    'box_remark_placeholder': '填写厢尺寸、侧门、尾板等要求',
}

# 静态配置 → 字典 category 映射（种子数据灌入用）
DICT_SEED_MAPPING = {
    # 保留旧报单结构所需的映射；实际落库分类以当前基准车型维度为准。
    'box_options': ('box_type', 'categories'),
    'brand': ('brand', 'common'),
    'product_series': ('product_series', 'common'),
    'battery_capacity': ('battery_capacity', 'common'),
    'horsepower': ('horsepower', 'common'),
    'gear_position': ('gear_position', 'common'),
    'battery_brand': ('battery_brand', 'common'),
}

# 所有字典 category 及其中文标签（设置页字典管理 Tab 用）
DICT_CATEGORIES = [
    ('brand', '品牌'), ('product_series', '品系'), ('box_type', '厢型'),
    ('battery_capacity', '电池度数'), ('horsepower', '马力'),
    ('gear_position', '档位'),
    ('battery_brand', '电池品牌'),
]


def seed_data_dictionaries(conn):
    """把静态配置 + 存量车辆维度值灌入 data_dictionaries（幂等）。"""
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    inserted = 0

    def _add(category, value, energy_type='', sort_order=0):
        nonlocal inserted
        if not value:
            return
        cur = conn.execute("""
            INSERT OR IGNORE INTO data_dictionaries (category, value, energy_type, sort_order, status, created_by, updated_by, updated_at)
            VALUES (?, ?, ?, ?, '启用', ?, ?, ?)
        """, (category, str(value).strip(), energy_type, sort_order, '系统', '系统', now))
        if cur.rowcount and cur.rowcount > 0:
            inserted += 1

    # 1) 静态配置 → 字典（仅保留基准车型与 SKU 主数据维度）
    for cfg in SALES_ORDER_CAR_OPTION_CONFIG:
        energy = cfg['category']
        for v in (cfg.get('box_options') or []):
            _add('box_type', v)
        if energy == '纯电':
            for v in (cfg.get('engine_battery_options') or []):
                _add('battery_brand', v, energy_type='纯电')
            for v in (cfg.get('power_battery_options') or []):
                _add('battery_capacity', v, energy_type='纯电')
        elif energy in ('混动', '燃油车'):
            for v in (cfg.get('power_battery_options') or []):
                _add('horsepower', v, energy_type=energy)
            for v in (cfg.get('gearbox_options') or []):
                _add('gear_position', v, energy_type=energy)
    for key, values in SALES_ORDER_COMMON_OPTIONS.items():
        if not isinstance(values, list):
            continue
        category = DICT_SEED_MAPPING.get(key, ('', ''))[0]
        if category not in dict(DICT_CATEGORIES):
            continue
        for v in values:
            _add(category, v)

    # 2) 存量车辆维度值补录（防下拉无值）
    vehicle_dim_cols = {
        'brand': 'brand', 'product_series': 'product_series', 'box_type': 'box_type',
        'battery_capacity': 'battery_capacity', 'horsepower': 'horsepower',
        'gear_position': 'gear_position',
        'battery_brand': 'battery_brand',
    }
    for category, col in vehicle_dim_cols.items():
        try:
            rows = conn.execute(
                f"""SELECT DISTINCT {col} AS v, fuel_form, battery_model,
                           battery_brand, battery_capacity, horsepower, gear_position
                    FROM vehicles
                    WHERE {col} IS NOT NULL AND TRIM({col}) != ''"""
            ).fetchall()
        except Exception:
            continue
        for r in rows:
            allowed_energies = DICT_CATEGORY_ENERGY_LIMITS.get(category, ())
            energy_type = energy_type_for_vehicle(dict(r))
            if allowed_energies and energy_type not in allowed_energies:
                continue
            _add(category, r['v'], energy_type=energy_type if allowed_energies else '')

    conn.commit()
    if inserted:
        print(f"Data dictionaries seeded: {inserted} entries")


# 车型列表（库存真实基准车型，按能源类型分组；替代旧硬编码 CAR_TYPE_PRESETS）
@app.route('/api/car-types', methods=['GET'])
@login_required
def get_car_types():
    """从库存车辆聚合最新基准车型：normalize_base_car_type 归一化（去成色/厢型、统一尾板），
    按能源类型分组返回，供车辆管理筛选等下拉使用。"""
    conn = get_db()
    rows = conn.execute("""
        SELECT car_type, fuel_form, battery_brand, battery_capacity, horsepower, gear_position
        FROM vehicles
        WHERE COALESCE(is_deleted,0)=0 AND car_type IS NOT NULL AND car_type != ''
          AND COALESCE(box_type, '') != '底盘'
    """).fetchall()
    conn.close()
    seen = {}
    order = []
    for row in rows:
        base = normalize_base_car_type(row['car_type'])
        if not base or base in seen:
            continue
        cat = energy_type_for_vehicle(dict(row)) or '其他'
        if cat == '燃油车':
            cat = '油车'
        seen[base] = cat
        order.append({'label': base, 'category': cat})
    group_order = {'纯电': 0, '混动': 1, '油车': 2, '其他': 3}
    order.sort(key=lambda x: (group_order.get(x['category'], 3), x['label']))
    return jsonify(order)


@app.route('/api/sales-order-car-options', methods=['GET'])
@login_required
def get_sales_order_car_options():
    """报单端规格配置。20260810 SKU 改造：从 data_dictionaries 组装，
    保持原 {categories, common_options} 结构（前端 fetchSalesOrderCarOptions 无需改动）。
    categories：能源专属值 + 通用值（去重）；common_options：全部启用值去重（兜底池）。"""
    conn = get_db()
    rows = conn.execute(
        "SELECT category, value, energy_type, sort_order, status FROM data_dictionaries WHERE status='启用' ORDER BY sort_order ASC, id ASC"
    ).fetchall()
    conn.close()

    def _dedup(vals):
        seen = set()
        out = []
        for v in vals:
            if v and v not in seen:
                seen.add(v)
                out.append(v)
        return out

    def _values(category, energy_type=''):
        return _dedup([
            r['value'] for r in rows
            if r['category'] == category
            and (
                (not DICT_CATEGORY_ENERGY_LIMITS.get(category) and not r['energy_type'])
                or r['energy_type'] == energy_type
            )
        ])

    categories = []
    for energy in ('纯电', '混动', '燃油车'):
        cat = {'category': energy, 'box_remark_placeholder': ''}
        static_cfg = next(
            (item for item in SALES_ORDER_CAR_OPTION_CONFIG if item['category'] == energy),
            {},
        )
        # 颜色、驾驶室属于车辆实例/报单快照，不属于可维护的数据字典。
        cat['cab_options'] = list(static_cfg.get('cab_options') or [])
        cat['color_options'] = list(static_cfg.get('color_options') or [])
        cat['box_options'] = _values('box_type')
        if energy == '纯电':
            cat['engine_battery_options'] = _values('battery_brand', energy)
            cat['power_battery_options'] = _values('battery_capacity', energy)
            cat['gearbox_options'] = []
        else:
            cat['engine_battery_options'] = []
            cat['power_battery_options'] = _values('horsepower', energy)
            cat['gearbox_options'] = _values('gear_position', energy)
        categories.append(cat)

    common = {
        'color_options': list(SALES_ORDER_COMMON_OPTIONS['color_options']),
    }
    box_values = _values('box_type')
    if box_values:
        common['box_options'] = box_values
    common['box_remark_placeholder'] = SALES_ORDER_COMMON_OPTIONS['box_remark_placeholder']

    return jsonify({
        'categories': categories,
        'common_options': common,
    })


# ======================== 数据字典 CRUD（SKU 改造）====================
@app.route('/api/data-dictionaries', methods=['GET'])
@login_required
def get_data_dictionaries():
    """字典查询；?category=xx & energy_type=xx 可选过滤。"""
    conn = get_db()
    category = (request.args.get('category') or '').strip()
    energy_type = (request.args.get('energy_type') or '').strip()
    sql = "SELECT * FROM data_dictionaries WHERE category NOT IN ({})".format(
        ','.join('?' for _ in RETIRED_DICT_CATEGORIES)
    )
    params = list(RETIRED_DICT_CATEGORIES)
    if category:
        sql += " AND category=?"
        params.append(category)
    if energy_type:
        sql += " AND energy_type=?"
        params.append(energy_type)
    sql += " ORDER BY category ASC, sort_order ASC, id ASC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    visible_rows = []
    for row in rows:
        item = dict(row)
        allowed_energies = DICT_CATEGORY_ENERGY_LIMITS.get(item['category'], ())
        # 旧版本写入的错误能耗范围只保留在库中，不再展示或参与操作。
        if allowed_energies and item.get('energy_type') not in allowed_energies:
            continue
        if not allowed_energies and item.get('energy_type'):
            continue
        visible_rows.append(item)
    return jsonify(visible_rows)


@app.route('/api/data-dictionaries', methods=['POST'])
@require_role('老板')
def upsert_data_dictionary():
    """老板新增/修改字典项。"""
    data = request.json or {}
    category = (data.get('category') or '').strip()
    value = (data.get('value') or '').strip()
    energy_type = (data.get('energy_type') or '').strip()
    sort_order = int(data.get('sort_order') or 0)
    remark = (data.get('remark') or '').strip()
    if not category:
        return jsonify({'success': False, 'message': '请选择维度分类'}), 400
    if not value:
        return jsonify({'success': False, 'message': '请填写选项值'}), 400
    if category in RETIRED_DICT_CATEGORIES:
        return jsonify({'success': False, 'message': f'维度“{category}”已停用，不再维护'}), 400
    if category not in dict(DICT_CATEGORIES):
        return jsonify({'success': False, 'message': f'未知维度分类：{category}'}), 400
    allowed_energies = DICT_CATEGORY_ENERGY_LIMITS.get(category, ())
    if allowed_energies and energy_type not in allowed_energies:
        return jsonify({
            'success': False,
            'message': f'{dict(DICT_CATEGORIES)[category]}仅适用于“{" / ".join(allowed_energies)}”',
        }), 400
    if not allowed_energies and energy_type:
        return jsonify({
            'success': False,
            'message': f'{dict(DICT_CATEGORIES)[category]}为通用维度，不应指定能源类型',
        }), 400

    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    try:
        dict_id = data.get('id')
        if dict_id:
            conn.execute("""
                UPDATE data_dictionaries
                SET value=?, energy_type=?, sort_order=?, remark=?, status=?, updated_by=?, updated_at=?
                WHERE id=?
            """, (value, energy_type, sort_order, remark, data.get('status') or '启用',
                  user['display_name'], now, dict_id))
        else:
            # 重复值则直接置回启用
            exists = conn.execute(
                """SELECT id FROM data_dictionaries
                   WHERE category=? AND value=? AND energy_type=?
                   LIMIT 1""",
                (category, value, energy_type)
            ).fetchone()
            if exists:
                conn.execute("""
                    UPDATE data_dictionaries SET status='启用', energy_type=?, sort_order=?, remark=?,
                        updated_by=?, updated_at=? WHERE id=?
                """, (energy_type, sort_order, remark, user['display_name'], now, exists['id']))
                dict_id = exists['id']
            else:
                cur = conn.execute("""
                    INSERT INTO data_dictionaries (category, value, energy_type, sort_order, status, created_by, updated_by, updated_at)
                    VALUES (?, ?, ?, ?, '启用', ?, ?, ?)
                """, (category, value, energy_type, sort_order, user['display_name'], user['display_name'], now))
                dict_id = cur.lastrowid
        log_audit(conn, '维护数据字典', 'data_dictionary', dict_id,
                  f'{category}「{value}」', user['display_name'])
        conn.commit()
        return jsonify({'success': True, 'id': dict_id, 'message': '字典项已保存'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


@app.route('/api/data-dictionaries/<int:did>', methods=['DELETE'])
@require_role('老板')
def delete_data_dictionary(did):
    """软删字典项（status='停用'）。"""
    conn = get_db()
    row = conn.execute("SELECT category, value FROM data_dictionaries WHERE id=?", (did,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '字典项不存在'}), 404
    conn.execute("UPDATE data_dictionaries SET status='停用', updated_by=?, updated_at=? WHERE id=?",
                 (request.current_user['display_name'], datetime.now().strftime('%Y-%m-%d %H:%M:%S'), did))
    log_audit(conn, '停用数据字典', 'data_dictionary', did,
              f"{row['category']}「{row['value']}」", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '字典项已停用'})


@app.route('/api/model-guidance-prices', methods=['GET'])
@login_required
def get_model_guidance_prices():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT car_type, is_new, guidance_price, lease_installment_price,
               lease_deposit_guidance,
               box_standard_price, box_wide_price, box_high_rail_price,
               box_refrigerated_price, box_flatbed_price, tail_plate_price,
               remark, updated_by, updated_at
        FROM model_guidance_prices
        ORDER BY car_type ASC, is_new ASC
    """)
    rows_all = [dict(row) for row in c.fetchall()]
    # 20260807 方案一：configured 键归一化为基准车型（去成色+去厢型），同 base+is_new 多行保留非0金额列最多的那行
    _MONEY_COLS = ['guidance_price', 'lease_installment_price',
                   'lease_deposit_guidance', 'box_standard_price', 'box_wide_price',
                   'box_high_rail_price', 'box_refrigerated_price', 'box_flatbed_price',
                   'tail_plate_price']

    def _money_count(item):
        return sum(1 for k in _MONEY_COLS if parse_money(item.get(k)) > 0)

    configured = {}
    for row in rows_all:
        base = normalize_base_car_type(row['car_type'])
        key = (base, row['is_new'])
        if key not in configured or _money_count(row) > _money_count(configured[key]):
            item = dict(row)
            item['car_type'] = base
            configured[key] = item

    if request.args.get('dict_only'):
        conn.close()
        return jsonify(list(configured.values()))

    c.execute("""
        SELECT car_type,
               COALESCE(NULLIF(condition, ''), '新车') AS is_new,
               COUNT(*) AS vehicle_count
        FROM vehicles
        WHERE COALESCE(car_type, '') != ''
          AND (is_deleted IS NULL OR is_deleted = 0)
        GROUP BY car_type, COALESCE(NULLIF(condition, ''), '新车')
        ORDER BY car_type ASC
    """)
    rows = []
    seen = set()
    # 多个带厢型 car_type 归一化到同一基准，并按真实成色累计车辆数。
    agg = {}
    for row in c.fetchall():
        base = normalize_base_car_type(row['car_type'])
        is_new = row['is_new'] if row['is_new'] in ('新车', '二手车') else '新车'
        key = (base, is_new)
        g = agg.setdefault(key, {'vehicle_count': 0})
        g['vehicle_count'] += row['vehicle_count']
    for key, g in sorted(agg.items()):
        base, is_new = key
        item = configured.get(key, {
            'car_type': base, 'is_new': is_new,
            'guidance_price': 0,
            'lease_installment_price': 0,
            'lease_deposit_guidance': 0,
            'box_standard_price': 0, 'box_wide_price': 0, 'box_high_rail_price': 0,
            'box_refrigerated_price': 0, 'box_flatbed_price': 0, 'tail_plate_price': 0,
            'remark': '',
            'updated_by': '',
            'updated_at': '',
        })
        item = dict(item)
        item['guidance_price'] = item.get('guidance_price') or 0
        item['lease_installment_price'] = item.get('lease_installment_price') or 0
        item['vehicle_count'] = g['vehicle_count']
        rows.append(item)
        seen.add(key)

    for key, item in configured.items():
        if key not in seen:
            item = dict(item)
            item['guidance_price'] = item.get('guidance_price') or 0
            item['lease_installment_price'] = item.get('lease_installment_price') or 0
            item['vehicle_count'] = 0
            rows.append(item)

    conn.close()
    return jsonify(rows)


@app.route('/api/model-guidance-workbench', methods=['GET'])
@require_role('老板')
def get_model_guidance_workbench():
    """返回一个基准车型的租赁覆盖矩阵与以租代售新车厢型方案矩阵。"""
    requested_car_type = (request.args.get('car_type') or '').strip()
    base_type = normalize_base_car_type(requested_car_type)
    if not base_type:
        return jsonify({'success': False, 'message': '请选择基准车型'}), 400

    conn = get_db()
    try:
        guidance_rows = conn.execute("""
            SELECT car_type, is_new, lease_deposit_guidance,
                   box_standard_price, box_wide_price, box_high_rail_price,
                   box_refrigerated_price, box_flatbed_price, tail_plate_price,
                   remark, updated_by, updated_at
            FROM model_guidance_prices
        """).fetchall()
        guidance_by_condition = {}
        for db_row in guidance_rows:
            row = dict(db_row)
            if normalize_base_car_type(row['car_type']) != base_type:
                continue
            condition = row['is_new'] if row['is_new'] in ('新车', '二手车') else '新车'
            previous = guidance_by_condition.get(condition)
            money_fields = ('lease_deposit_guidance',) + tuple(BOX_TYPE_MONTHLY_PRICE_COLUMNS.values())
            if not previous or sum(parse_money(row.get(k)) > 0 for k in money_fields) > sum(
                parse_money(previous.get(k)) > 0 for k in money_fields
            ):
                guidance_by_condition[condition] = row

        inventory_counts = {}
        vehicles = conn.execute("""
            SELECT car_type, condition, box_type
            FROM vehicles
            WHERE (is_deleted IS NULL OR is_deleted=0)
              AND COALESCE(status, '') IN ('在库', '报单锁定中')
        """).fetchall()
        for db_vehicle in vehicles:
            vehicle = dict(db_vehicle)
            if normalize_base_car_type(vehicle['car_type']) != base_type:
                continue
            condition = vehicle['condition'] if vehicle['condition'] in ('新车', '二手车') else '新车'
            box_type = (vehicle['box_type'] or '').strip()
            if box_type in GUIDANCE_BOX_TYPES:
                key = (condition, box_type)
                inventory_counts[key] = inventory_counts.get(key, 0) + 1

        lease_rows = []
        for condition in ('新车', '二手车'):
            price_row = guidance_by_condition.get(condition, {})
            for box_type, price_field in BOX_TYPE_MONTHLY_PRICE_COLUMNS.items():
                deposit = parse_money(price_row.get('lease_deposit_guidance'))
                monthly = parse_money(price_row.get(price_field))
                lease_rows.append({
                    'condition': condition,
                    'box_type': box_type,
                    'price_field': price_field,
                    'inventory_count': inventory_counts.get((condition, box_type), 0),
                    'lease_deposit_guidance': deposit,
                    'monthly_price': monthly,
                    'tail_plate_price': parse_money(price_row.get('tail_plate_price')) or TAIL_PLATE_SURCHARGE,
                    'remark': price_row.get('remark') or '',
                    'updated_by': price_row.get('updated_by') or '',
                    'updated_at': price_row.get('updated_at') or '',
                    'status': '已设置' if deposit > 0 and monthly > 0 else '待设置',
                })

        plans = []
        for db_plan in conn.execute(
            "SELECT * FROM finance_plans ORDER BY sort_order ASC, id DESC"
        ).fetchall():
            plan = dict(db_plan)
            if normalize_base_car_type(plan['car_type']) == base_type:
                plan['car_type'] = base_type
                plans.append(plan)

        plans_by_box_type = {box_type: [] for box_type in GUIDANCE_BOX_TYPES}
        legacy_plans = []
        for plan in plans:
            plan_condition = normalize_vehicle_condition(plan.get('condition'))
            plan_box_type = (plan.get('box_type') or '').strip()
            if plan_condition == '新车' and plan_box_type in plans_by_box_type:
                plan['condition'] = '新车'
                plans_by_box_type[plan_box_type].append(plan)
            else:
                # 历史方案原先只有基准车型，无法可靠推断厢型，不能参与报单匹配。
                legacy_plans.append(plan)

        rent_to_buy_rows = []
        for box_type in GUIDANCE_BOX_TYPES:
            box_plans = plans_by_box_type[box_type]
            active_count = sum(1 for plan in box_plans if plan.get('status') != '停用')
            rent_to_buy_rows.append({
                'condition': '新车',
                'box_type': box_type,
                'inventory_count': inventory_counts.get(('新车', box_type), 0),
                'plan_count': len(box_plans),
                'active_plan_count': active_count,
                'status': '已设置' if active_count else '待设置',
                'plans': box_plans,
            })

        configured_count = sum(1 for row in lease_rows if row['status'] == '已设置')
        return jsonify({
            'success': True,
            'car_type': base_type,
            'lease_rows': lease_rows,
            'finance_plans': plans,
            'rent_to_buy_rows': rent_to_buy_rows,
            'legacy_finance_plans': legacy_plans,
            'summary': {
                'total_count': len(lease_rows),
                'configured_count': configured_count,
                'missing_count': len(lease_rows) - configured_count,
                'rent_to_buy_total_count': len(rent_to_buy_rows),
                'rent_to_buy_configured_count': sum(
                    1 for row in rent_to_buy_rows if row['status'] == '已设置'
                ),
                'rent_to_buy_plan_count': sum(row['plan_count'] for row in rent_to_buy_rows),
                'legacy_finance_plan_count': len(legacy_plans),
            },
        })
    finally:
        conn.close()


@app.route('/api/guidance-price-alerts', methods=['GET'])
@require_role('老板')
def get_guidance_price_alerts():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT car_type,
               COALESCE(NULLIF(condition, ''), '新车') as condition,
               COUNT(*) as vehicle_count,
               MAX(created_at) as last_created_at
        FROM vehicles
        WHERE (is_deleted IS NULL OR is_deleted = 0)
          AND COALESCE(box_type, '') != '底盘'
          AND COALESCE(status, '') IN ('在库', '报单锁定中')
        GROUP BY car_type, condition
    """)
    vehicle_groups = [dict(r) for r in c.fetchall()]
    c.execute("""
        SELECT car_type, is_new,
               guidance_price, lease_deposit_guidance,
               box_standard_price, box_wide_price, box_high_rail_price,
               box_refrigerated_price, box_flatbed_price
        FROM model_guidance_prices
    """)
    guidance_rows = [dict(r) for r in c.fetchall()]
    conn.close()

    def _configured(r):
        return any(parse_money(r.get(k)) > 0 for k in (
            'guidance_price', 'lease_deposit_guidance',
            'box_standard_price', 'box_wide_price', 'box_high_rail_price',
            'box_refrigerated_price', 'box_flatbed_price'))

    configured_keys = set()
    for r in guidance_rows:
        base = normalize_base_car_type(r['car_type'])
        if _configured(r):
            configured_keys.add((base, r['is_new']))

    # 按基准车型+成色归组，未配置的汇总为提醒
    groups = {}
    for v in vehicle_groups:
        base = normalize_base_car_type(v['car_type'])
        key = (base, v['condition'])
        if key in configured_keys:
            continue
        g = groups.setdefault(key, {
            'car_type': base, 'condition': v['condition'],
            'vehicle_count': 0, 'last_created_at': v['last_created_at'],
        })
        g['vehicle_count'] += v['vehicle_count']
        if (v['last_created_at'] or '') > (g['last_created_at'] or ''):
            g['last_created_at'] = v['last_created_at']

    rows = sorted(groups.values(), key=lambda r: r['last_created_at'] or '', reverse=True)
    total_vehicles = sum(r.get('vehicle_count') or 0 for r in rows)
    return jsonify({
        'count': len(rows),
        'vehicle_count': total_vehicles,
        'items': rows,
        'message': f'有 {len(rows)} 个车型未设置指导价（涉及 {total_vehicles} 台车）' if rows else '',
    })



@app.route('/api/model-guidance-prices/history', methods=['GET'])
@require_role('老板')
def get_guidance_price_history():
    car_type = request.args.get('car_type', '').strip()
    if not car_type:
        return jsonify([])
    conn = get_db()
    rows = conn.execute("""
        SELECT price_kind, old_price, new_price, changed_by, effective_at, remark
        FROM model_guidance_price_history
        WHERE car_type=?
        ORDER BY effective_at DESC
        LIMIT 30
    """, (car_type,)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/model-guidance-prices', methods=['POST'])
@require_role('老板')
def upsert_model_guidance_price():
    data = request.json or {}
    car_type = (data.get('car_type') or '').strip()
    is_new = (data.get('is_new') or '新车').strip()
    if is_new not in ('新车', '二手车'):
        is_new = '新车'
    base_type = normalize_base_car_type(car_type)  # 去成色前缀+去厢型后缀，指导价按基准车型存（20260807）
    legacy_input = data.get('guidance_price', data.get('price'))
    lease_price = parse_money(data.get('lease_installment_price'))
    new_price = parse_money(legacy_input) or lease_price
    remark = (data.get('remark') or '').strip()
    # 新增Excel字段
    chassis_base_price = parse_money(data.get('chassis_base_price')) or 0
    landing_price = parse_money(data.get('landing_price')) or 0
    interest_free_plan = (data.get('interest_free_plan') or '').strip() or None
    rent_to_buy_plan = (data.get('rent_to_buy_plan') or '').strip() or None
    min_loan_plan = (data.get('min_loan_plan') or '').strip() or None
    lease_plan = (data.get('lease_plan') or '').strip() or None
    product_code = (data.get('product_code') or '').strip() or None
    fuel_type = (data.get('fuel_type') or '').strip() or None
    if landing_price: new_price = new_price or landing_price
    # 20260804 租赁指导价：押金 + 箱型月供 + 尾板加价
    lease_deposit_guidance = parse_money(data.get('lease_deposit_guidance')) or 0
    box_standard_price = parse_money(data.get('box_standard_price')) or 0
    box_wide_price = parse_money(data.get('box_wide_price')) or 0
    box_high_rail_price = parse_money(data.get('box_high_rail_price')) or 0
    box_refrigerated_price = parse_money(data.get('box_refrigerated_price')) or 0
    box_flatbed_price = parse_money(data.get('box_flatbed_price')) or 0
    tail_plate_price = parse_money(data.get('tail_plate_price')) or 0
    if not car_type:
        return jsonify({'success': False, 'message': '请选择或填写车型'}), 400

    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    try:
        # 将旧版未显式记录“无尾板”的指导价在首次保存时迁移到新基准名，
        # 避免同一无尾板车型在新旧名称下形成两条维护记录。
        if base_type.endswith('无尾板'):
            legacy_base_type = base_type[:-len('无尾板')]
            legacy_row = c.execute(
                "SELECT 1 FROM model_guidance_prices WHERE car_type=? AND is_new=?",
                (legacy_base_type, is_new)
            ).fetchone()
            canonical_row = c.execute(
                "SELECT 1 FROM model_guidance_prices WHERE car_type=? AND is_new=?",
                (base_type, is_new)
            ).fetchone()
            if legacy_row and not canonical_row:
                c.execute(
                    "UPDATE model_guidance_prices SET car_type=? WHERE car_type=? AND is_new=?",
                    (base_type, legacy_base_type, is_new)
                )
        c.execute("""
            SELECT guidance_price, lease_installment_price
            FROM model_guidance_prices
            WHERE car_type=? AND is_new=?
        """, (base_type, is_new))
        existing = c.fetchone()
        old_price = existing['guidance_price'] if existing else 0
        old_lease_price = existing['lease_installment_price'] if existing else 0

        # 20260807 方案一：按基准车型同步车辆（LIKE 后缀无法命中带厢型车名，改 Python 归一化过滤）
        all_vehicles = conn.execute(
            "SELECT id, car_type FROM vehicles WHERE (is_deleted IS NULL OR is_deleted = 0)"
        ).fetchall()
        affected_vehicles = [v for v in all_vehicles if normalize_base_car_type(v['car_type']) == base_type]
        affected_count = len(affected_vehicles)

        c.execute("""
            INSERT INTO model_guidance_prices
                (car_type, is_new, guidance_price, lease_installment_price,
                 chassis_base_price, landing_price, interest_free_plan, rent_to_buy_plan,
                 min_loan_plan, lease_plan, product_code, fuel_type,
                 lease_deposit_guidance, box_standard_price, box_wide_price, box_high_rail_price,
                 box_refrigerated_price, box_flatbed_price, tail_plate_price,
                 remark, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(car_type, is_new) DO UPDATE SET
                guidance_price=excluded.guidance_price,
                lease_installment_price=excluded.lease_installment_price,
                chassis_base_price=excluded.chassis_base_price,
                landing_price=excluded.landing_price,
                interest_free_plan=excluded.interest_free_plan,
                rent_to_buy_plan=excluded.rent_to_buy_plan,
                min_loan_plan=excluded.min_loan_plan,
                lease_plan=excluded.lease_plan,
                product_code=COALESCE(excluded.product_code, product_code),
                fuel_type=COALESCE(excluded.fuel_type, fuel_type),
                lease_deposit_guidance=excluded.lease_deposit_guidance,
                box_standard_price=excluded.box_standard_price,
                box_wide_price=excluded.box_wide_price,
                box_high_rail_price=excluded.box_high_rail_price,
                box_refrigerated_price=excluded.box_refrigerated_price,
                box_flatbed_price=excluded.box_flatbed_price,
                tail_plate_price=excluded.tail_plate_price,
                remark=excluded.remark,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at
        """, (
            base_type, is_new, new_price, lease_price,
            chassis_base_price, landing_price, interest_free_plan, rent_to_buy_plan,
            min_loan_plan, lease_plan, product_code, fuel_type,
            lease_deposit_guidance, box_standard_price, box_wide_price, box_high_rail_price,
            box_refrigerated_price, box_flatbed_price, tail_plate_price,
            remark, user['display_name'], now,
        ))
        history_rows = [
            ('lease_installment', old_lease_price or old_price or 0, lease_price),
        ]
        for price_kind, old_value, new_value in history_rows:
            c.execute("""
                INSERT INTO model_guidance_price_history
                    (car_type, price_kind, old_price, new_price, changed_by, effective_at, affected_vehicle_count, remark)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (car_type, price_kind, old_value, new_value, user['display_name'], now, affected_count, remark))

        for vehicle in affected_vehicles:
            old_hist = c.execute("""
                SELECT new_price FROM vehicle_guidance_price_history
                WHERE vehicle_id=? ORDER BY effective_at DESC, id DESC LIMIT 1
            """, (vehicle['id'],)).fetchone()
            old_price = old_hist['new_price'] if old_hist else 0
            c.execute("""
                INSERT INTO vehicle_guidance_price_history
                    (vehicle_id, old_price, new_price, changed_by, effective_at)
                VALUES (?, ?, ?, ?, ?)
            """, (vehicle['id'], old_price, new_price, user['display_name'], now))
        remaining_missing_count = unresolved_guidance_vehicle_count(conn)

        log_audit(conn, '更新车型指导价', 'model_guidance_price', None,
                  f'{user["display_name"]} 将车型「{car_type}」租赁指导价调为月租{lease_price or new_price}，同步车辆 {affected_count} 台',
                  user['display_name'])
        conn.commit()
        return jsonify({
            'success': True,
            'message': f'车型指导价已更新，同步 {affected_count} 台库存车辆',
            'affected_vehicle_count': affected_count,
            'remaining_missing_guidance_count': remaining_missing_count,
            'lease_installment_price': lease_price,
        })
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()



# 指导价 Excel 导入
@app.route('/api/model-guidance-prices/import', methods=['POST'])
@require_role('老板')
def import_guidance_prices():
    import re
    from openpyxl import load_workbook
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '请上传文件'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'message': '请上传 .xlsx/.xls 文件'}), 400

    def parse_monthly(text):
        """从 '首付3万，3600三年' / '首付3万含1年保险，月还4500，3年' 提取月供金额"""
        if not text or str(text).strip() in ('-', '无', ''):
            return None
        s = str(text)
        m = re.search(r'月还(\d+)', s)
        if m:
            return float(m.group(1))
        nums = re.findall(r'\d+', s)
        return float(nums[1]) if len(nums) >= 2 else None

    def to_yuan(v):
        try:
            return float(v) * 10000 if v else 0
        except:
            return 0

    wb = load_workbook(f, data_only=True)
    conn = get_db()
    c = conn.cursor()
    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    created, updated, skipped = 0, 0, 0

    # sheet名→fuel_type映射；优先匹配含"油"或"新能"的sheet
    sheet_map = {}
    for name in wb.sheetnames:
        if '新能' in name or '电' in name:
            sheet_map[name] = '新能源'
        elif '油' in name or '燃' in name:
            sheet_map[name] = '油车'
    if not sheet_map:  # 兜底：取前两个sheet
        for i, name in enumerate(wb.sheetnames[:2]):
            sheet_map[name] = ['油车', '新能源'][i]

    for sheet_name, fuel_type in sheet_map.items():
        ws = wb[sheet_name]
        headers = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]
        col = {h: i for i, h in enumerate(headers)}

        def get(row, key, default=None):
            idx = col.get(key)
            return row[idx] if idx is not None and idx < len(row) else default

        for row_vals in ws.iter_rows(min_row=2, values_only=True):
            car_type = get(row_vals, '车型')
            if not car_type or not isinstance(car_type, str) or car_type.startswith('1、'):
                continue
            car_type = car_type.strip()

            sale_total = to_yuan(get(row_vals, '整车'))
            landing = to_yuan(get(row_vals, '落地不含商业险'))
            chassis = to_yuan(get(row_vals, '底盘底价'))
            interest_free = str(get(row_vals, '免息金额') or get(row_vals, '0息方案') or '').strip() or None
            rent_buy = str(get(row_vals, '以租代购方案') or '').strip() or None
            min_loan = str(get(row_vals, '最低贷款方案优惠') or '').strip() or None
            lease_plan = str(get(row_vals, '租赁方案') or '').strip() or None
            product_code = str(get(row_vals, '产品码') or '').strip() or None
            monthly = parse_monthly(rent_buy)
            # sale_total_price（以租代售整车价）已废弃：Excel「整车」列仅作为历史单指导价展示来源
            guidance = landing or sale_total

            c.execute("SELECT id FROM model_guidance_prices WHERE car_type=?", (car_type,))
            exists = c.fetchone()
            if exists:
                c.execute("""
                    UPDATE model_guidance_prices SET
                        product_code=COALESCE(?, product_code),
                        fuel_type=?, chassis_base_price=?,
                        landing_price=?, guidance_price=?,
                        interest_free_plan=?, rent_to_buy_plan=?, min_loan_plan=?,
                        lease_plan=?,
                        lease_installment_price=CASE WHEN ? IS NOT NULL THEN ? ELSE lease_installment_price END,
                        updated_by=?, updated_at=?
                    WHERE car_type=?
                """, (product_code, fuel_type, chassis, landing, guidance,
                      interest_free, rent_buy, min_loan, lease_plan,
                      monthly, monthly, user['display_name'], now, car_type))
                updated += 1
            else:
                c.execute("""
                    INSERT INTO model_guidance_prices
                        (car_type, product_code, fuel_type, chassis_base_price,
                         landing_price, guidance_price, interest_free_plan, rent_to_buy_plan,
                         min_loan_plan, lease_plan, lease_installment_price,
                         updated_by, updated_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (car_type, product_code, fuel_type, chassis,
                      landing, guidance, interest_free, rent_buy, min_loan,
                      lease_plan, monthly or 0, user['display_name'], now))
                created += 1

    conn.commit()
    conn.close()
    return jsonify({'success': True, 'created': created, 'updated': updated,
                    'message': f'导入完成：新建 {created} 条，更新 {updated} 条'})


# ======================== 以租代售金融方案 CRUD ========================
@app.route('/api/finance-plans', methods=['GET'])
@login_required
def get_finance_plans():
    """金融方案列表；可按基准车型、新车厢型过滤。"""
    conn = get_db()
    c = conn.cursor()
    car_type = (request.args.get('car_type') or '').strip()
    condition = (request.args.get('condition') or '').strip()
    box_type = (request.args.get('box_type') or '').strip()
    if car_type:
        base = normalize_base_car_type(car_type)
        rows = [
            dict(row) for row in c.execute(
                "SELECT * FROM finance_plans ORDER BY sort_order ASC, id DESC"
            ).fetchall()
            if normalize_base_car_type(row['car_type']) == base
        ]
    else:
        rows = [dict(r) for r in c.execute(
            "SELECT * FROM finance_plans ORDER BY car_type ASC, sort_order ASC, id DESC"
        ).fetchall()]
    if condition:
        rows = [
            row for row in rows
            if normalize_vehicle_condition(row.get('condition')) == normalize_vehicle_condition(condition)
        ]
    if box_type:
        rows = [row for row in rows if (row.get('box_type') or '').strip() == box_type]
    conn.close()
    return jsonify(rows)


@app.route('/api/finance-plans', methods=['POST'])
@require_role('老板')
def upsert_finance_plan():
    """老板新增/修改金融方案（基准车型 + 新车 + 厢型范围）。"""
    data = request.json or {}
    raw_car_type = (data.get('car_type') or '').strip()
    car_type = normalize_base_car_type(raw_car_type)
    condition = normalize_vehicle_condition(
        (data.get('condition') or ('二手车' if raw_car_type.startswith('二手车') else '新车')).strip()
    )
    box_type = (data.get('box_type') or extract_guidance_box_type(raw_car_type) or '').strip()
    plan_name = (data.get('plan_name') or '').strip()
    down_payment = parse_money(data.get('down_payment'))
    period_price = parse_money(data.get('period_price'))
    periods = int(parse_money(data.get('periods')) or 0)
    sort_order = int(parse_money(data.get('sort_order')) or 0)
    remark = (data.get('remark') or '').strip()
    # 20260810 SKU 改造：生效/失效日期
    effective_date = (data.get('effective_date') or '').strip() or None
    expiry_date = (data.get('expiry_date') or '').strip() or None
    if effective_date and expiry_date and effective_date > expiry_date:
        return jsonify({'success': False, 'message': '生效日期不能晚于失效日期'}), 400
    if not car_type:
        return jsonify({'success': False, 'message': '请填写车型'}), 400
    if condition != '新车':
        return jsonify({'success': False, 'message': '以租代售方案仅支持新车子型号'}), 400
    if box_type not in GUIDANCE_BOX_TYPES:
        return jsonify({'success': False, 'message': '请选择厢货、宽体、高栏、冷藏或平板中的一个厢型'}), 400
    if down_payment <= 0 or period_price <= 0 or periods <= 0:
        return jsonify({'success': False, 'message': '首付款、每期价格、期数均须大于0'}), 400
    if not plan_name:
        plan_name = f"首付{round(down_payment)}/{round(period_price)}×{periods}期"

    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    try:
        plan_id = data.get('id')
        if plan_id:
            c.execute("""
                UPDATE finance_plans
                SET car_type=?, condition=?, box_type=?, plan_name=?, down_payment=?, period_price=?, periods=?,
                    sort_order=?, remark=?, status=?, effective_date=?, expiry_date=?,
                    updated_by=?, updated_at=?
                WHERE id=?
            """, (car_type, condition, box_type, plan_name, down_payment, period_price, periods,
                  sort_order, remark, data.get('status') or '启用',
                  effective_date, expiry_date,
                  user['display_name'], now, plan_id))
            log_audit(conn, '修改金融方案', 'finance_plan', plan_id,
                      f'{car_type} 新车{box_type} {plan_name} 首付{down_payment}/每期{period_price}×{periods}期', user['display_name'])
        else:
            c.execute("""
                INSERT INTO finance_plans
                    (car_type, condition, box_type, plan_name, down_payment, period_price, periods,
                     sort_order, remark, status, effective_date, expiry_date,
                     created_by, updated_by, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '启用', ?, ?, ?, ?, ?)
            """, (car_type, condition, box_type, plan_name, down_payment, period_price, periods,
                  sort_order, remark, effective_date, expiry_date,
                  user['display_name'], user['display_name'], now))
            plan_id = c.lastrowid
            log_audit(conn, '新增金融方案', 'finance_plan', plan_id,
                      f'{car_type} 新车{box_type} {plan_name} 首付{down_payment}/每期{period_price}×{periods}期', user['display_name'])
        conn.commit()
        return jsonify({'success': True, 'id': plan_id, 'message': '金融方案已保存'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


@app.route('/api/finance-plans/<int:pid>', methods=['DELETE'])
@require_role('老板')
def delete_finance_plan(pid):
    """软删金融方案（status='停用'）。已出库报单走快照不受影响。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT car_type, condition, box_type, plan_name FROM finance_plans WHERE id=?", (pid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '方案不存在'}), 404
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("UPDATE finance_plans SET status='停用', updated_by=?, updated_at=? WHERE id=?",
              (request.current_user['display_name'], now, pid))
    log_audit(conn, '停用金融方案', 'finance_plan', pid,
              f"{row['car_type']} {row['condition'] or '新车'}{row['box_type'] or '未归属'} {row['plan_name'] or ''}",
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '金融方案已停用'})


# ======================== SKU 主表 + 库存聚合（SKU 改造）====================
def _sku_match(v, sku):
    """车辆是否匹配 SKU（归一化在 Python 端做，SQL 无法调用 Python 归一化函数）。"""
    if not v:
        return False
    if normalize_base_car_type((v.get('car_type') or '')) != sku['car_type']:
        return False
    if (v.get('condition') or '新车') != sku['condition']:
        return False
    if (v.get('box_type') or '') != sku['box_type']:
        return False
    tail = normalize_tailgate(v.get('tailgate') or v.get('tail_plate'))
    if tail != sku['tailgate']:
        return False
    return True


def _load_sku_inventory(conn, skus, only_in_stock=True):
    """为 SKU 列表聚合在库/总车辆数（Python 端归一化匹配）。"""
    vehicles = conn.execute("""
        SELECT id, car_type, condition, box_type, tailgate, status
        FROM vehicles
        WHERE (is_deleted IS NULL OR is_deleted=0)
    """).fetchall()
    vdicts = [dict(v) for v in vehicles]
    result = []
    for s in skus:
        sdict = dict(s)
        in_stock = sum(1 for v in vdicts if v['status'] == '在库' and _sku_match(v, sdict))
        total = sum(1 for v in vdicts if _sku_match(v, sdict))
        item = dict(sdict)
        item['in_stock_count'] = in_stock
        item['total_count'] = total
        item['sku_id'] = item.get('sku_id') or item.get('id')
        result.append(item)
    return result


@app.route('/api/skus', methods=['GET'])
@login_required
def get_skus():
    """SKU 列表 + 按 SKU 聚合在库数量。?car_type=xx & condition=xx & box_type=xx & tailgate=xx 可选过滤。"""
    conn = get_db()
    car_type = (request.args.get('car_type') or '').strip()
    condition = (request.args.get('condition') or '').strip()
    box_type = (request.args.get('box_type') or '').strip()
    tailgate = (request.args.get('tailgate') or '').strip()

    where = "WHERE 1=1"
    params = []
    if car_type:
        where += " AND s.car_type=?"
        params.append(normalize_base_car_type(car_type))
    if condition:
        where += " AND s.condition=?"
        params.append(condition)
    if box_type:
        where += " AND s.box_type=?"
        params.append(box_type)
    if tailgate:
        where += " AND s.tailgate=?"
        params.append(tailgate)

    rows = conn.execute(f"""
        SELECT s.*
        FROM skus s
        {where}
        ORDER BY s.car_type ASC, s.condition ASC, s.box_type ASC, s.tailgate ASC
    """, params).fetchall()
    result = _load_sku_inventory(conn, rows, only_in_stock=False)
    conn.close()
    return jsonify(result)


@app.route('/api/skus/inventory', methods=['GET'])
@login_required
def get_sku_inventory():
    """SKU 实时库存视图：生效 SKU 按在库车辆聚合数量。
    ?car_type=xx & condition=xx & box_type=xx & tailgate=xx 可选过滤（供销售端筛选联动）。"""
    conn = get_db()
    car_type = (request.args.get('car_type') or '').strip()
    condition = (request.args.get('condition') or '').strip()
    box_type = (request.args.get('box_type') or '').strip()
    tailgate = (request.args.get('tailgate') or '').strip()

    where = "WHERE s.status='生效'"
    params = []
    if car_type:
        where += " AND s.car_type=?"
        params.append(normalize_base_car_type(car_type))
    if condition:
        where += " AND s.condition=?"
        params.append(condition)
    if box_type:
        where += " AND s.box_type=?"
        params.append(box_type)
    if tailgate:
        where += " AND s.tailgate=?"
        params.append(tailgate)

    rows = conn.execute(f"""
        SELECT s.id AS sku_id, s.car_type, s.condition, s.box_type, s.tailgate, s.status
        FROM skus s
        {where}
        ORDER BY s.car_type ASC, s.condition ASC, s.box_type ASC, s.tailgate ASC
    """, params).fetchall()
    result = _load_sku_inventory(conn, rows, only_in_stock=True)
    conn.close()
    return jsonify(result)


@app.route('/api/skus/<int:sku_id>/vehicles', methods=['GET'])
@login_required
def get_sku_vehicles(sku_id):
    """SKU 下在库车辆列表（销售选 VIN 用）。"""
    conn = get_db()
    sku = conn.execute("SELECT * FROM skus WHERE id=?", (sku_id,)).fetchone()
    if not sku:
        conn.close()
        return jsonify({'success': False, 'message': 'SKU 不存在'}), 404
    sku_dict = dict(sku)
    rows = conn.execute("""
        SELECT id, vin, car_type, plate_number, condition, box_type, tailgate,
               purchase_price, created_at, status,
               vehicle_category, vehicle_engine_battery, vehicle_power_battery,
               vehicle_color, box_type_remark
        FROM vehicles
        WHERE (is_deleted IS NULL OR is_deleted=0)
        ORDER BY id ASC
    """).fetchall()
    matched = [dict(r) for r in rows if r['status'] == '在库' and _sku_match(dict(r), sku_dict)]
    conn.close()
    return jsonify(matched)


@app.route('/api/skus/<int:sku_id>/toggle', methods=['POST'])
@require_role('老板')
def toggle_sku(sku_id):
    """老板启用/停用 SKU（停用后销售端不可见）。"""
    data = request.json or {}
    conn = get_db()
    row = conn.execute("SELECT car_type, condition, box_type, tailgate FROM skus WHERE id=?", (sku_id,)).fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': 'SKU 不存在'}), 404
    status = '生效' if data.get('status') != '失效' else '失效'
    conn.execute("UPDATE skus SET status=?, updated_by=?, updated_at=? WHERE id=?",
                 (status, request.current_user['display_name'],
                  datetime.now().strftime('%Y-%m-%d %H:%M:%S'), sku_id))
    log_audit(conn, '变更SKU状态', 'sku', sku_id,
              f"{row['car_type']} {row['condition']} {row['box_type']} {row['tailgate']} → {status}",
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'SKU 已{status}'})


def backfill_skus(conn, operator='系统'):
    """启动时回填：为存量在库车辆自动创建/关联 SKU（幂等）。返回新增数量。"""
    created = 0
    try:
        vehicles = conn.execute("""
            SELECT id, car_type, condition, box_type, tailgate, status
            FROM vehicles
            WHERE (is_deleted IS NULL OR is_deleted=0)
        """).fetchall()
        for v in vehicles:
            vd = dict(v)
            key = sku_key_for_vehicle(vd)
            if not key:
                continue
            cur = conn.execute("""
                INSERT OR IGNORE INTO skus (car_type, condition, box_type, tailgate, status, created_by, updated_by, updated_at)
                VALUES (?, ?, ?, ?, '生效', ?, ?, ?)
            """, (key['car_type'], key['condition'], key['box_type'], key['tailgate'],
                  operator, operator, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            if cur.rowcount and cur.rowcount > 0:
                created += 1
        conn.commit()
    except Exception as e:
        print(f'[backfill_skus] error: {e}')
    return created


# ======================== 报单驳回退款闭环 ========================
@app.route('/api/order-refunds', methods=['GET'])
@login_required
def get_order_refunds():
    """退款列表；?status=待退款 可选过滤。"""
    conn = get_db()
    c = conn.cursor()
    status = (request.args.get('status') or '').strip()
    if status:
        c.execute("SELECT * FROM order_refunds WHERE status=? ORDER BY id DESC", (status,))
    else:
        c.execute("SELECT * FROM order_refunds ORDER BY id DESC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/order-refunds', methods=['POST'])
@require_role('老板')
def create_order_refund():
    """老板发起退款：报单须已作废且未发起过退款。"""
    data = request.json or {}
    sales_order_id = data.get('sales_order_id')
    if not sales_order_id:
        return jsonify({'success': False, 'message': '请指定报单'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT so.*, v.plate_number FROM sales_orders so
        LEFT JOIN vehicles v ON v.id = so.vehicle_id
        WHERE so.id=?
    """, (sales_order_id,))
    order = c.fetchone()
    if not order:
        conn.close()
        return jsonify({'success': False, 'message': '报单不存在'}), 404
    if order['order_status'] != '已作废':
        conn.close()
        return jsonify({'success': False, 'message': f'仅已作废的报单可发起退款（当前 {order["order_status"]}）'}), 400
    if order['refund_id']:
        conn.close()
        return jsonify({'success': False, 'message': '该报单已发起退款，请勿重复操作'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = request.current_user['display_name']
    c.execute("""
        INSERT INTO order_refunds
            (sales_order_id, vehicle_id, contract_id, refund_amount, status,
             initiated_by, initiated_at, customer_name, customer_phone, remark)
        VALUES (?, ?, ?, ?, '待退款', ?, ?, ?, ?, ?)
    """, (
        sales_order_id,
        order['vehicle_id'],
        order['contract_id'],
        parse_money(order['first_payment_received_amount']),
        user,
        now,
        order['customer_name'] or '',
        order['customer_phone'] or '',
        (data.get('remark') or '').strip(),
    ))
    refund_id = c.lastrowid
    c.execute("UPDATE sales_orders SET refund_id=? WHERE id=?", (refund_id, sales_order_id))
    log_audit(conn, '发起报单退款', 'sales_order', sales_order_id,
              f'应退 ¥{parse_money(order["first_payment_received_amount"])} 车辆 {order["plate_number"] or ""}',
              user)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': refund_id, 'message': '退款已发起，等待财务执行'})


@app.route('/api/order-refunds/<int:rid>/pay', methods=['POST'])
@require_role('财务')
def pay_order_refund(rid):
    """财务执行退款：填流水号+金额 → 已退款 → 车辆释放回在库。"""
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM order_refunds WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退款单不存在'}), 404
    if row['status'] != '待退款':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能执行退款'}), 400
    refund_serial = (data.get('refund_serial') or '').strip()
    refund_paid_amount = data.get('refund_paid_amount')
    if len(refund_serial) < 4:
        conn.close()
        return jsonify({'success': False, 'message': '银行流水号至少填写4位'}), 400
    try:
        refund_paid_amount = float(refund_paid_amount)
        if refund_paid_amount <= 0:
            raise ValueError
    except (TypeError, ValueError):
        conn.close()
        return jsonify({'success': False, 'message': '请填写有效的退款金额'}), 400
    refund_due = round(parse_money(row['refund_amount']), 2)
    if round(refund_paid_amount, 2) != refund_due:
        conn.close()
        return jsonify({'success': False, 'message': f'退款金额必须与应退金额一致（应退 ¥{refund_due}）'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = request.current_user['display_name']
    c.execute("""
        UPDATE order_refunds
        SET status='已退款', refund_serial=?, refund_paid_amount=?,
            bank_name=?, bank_card_no=?, executed_by=?, executed_at=?
        WHERE id=?
    """, (
        refund_serial, refund_paid_amount,
        (data.get('bank_name') or '').strip(),
        (data.get('bank_card_no') or '').strip(),
        user, now, rid,
    ))
    # 车辆释放回在库（SKU 改造：多车报单批量释放）
    release_vids = []
    if row['vehicle_id']:
        release_vids = [row['vehicle_id']]
    if row['sales_order_id']:
        so = c.execute("SELECT vehicle_id, vehicle_ids FROM sales_orders WHERE id=?", (row['sales_order_id'],)).fetchone()
        if so:
            release_vids = order_vehicle_ids(conn, so) or release_vids
    for vid in release_vids:
        c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (vid,))
    # 若有合同壳则终止
    if row['contract_id']:
        c.execute("""
            UPDATE contracts SET contract_status='已终止', delivery_status='已终止'
            WHERE id=? AND contract_status='报单计划中'
        """, (row['contract_id'],))
    log_audit(conn, '执行报单退款', 'order_refund', rid,
              f'报单{row["sales_order_id"]} 退款 ¥{refund_paid_amount} 流水号{refund_serial}', user)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '退款已执行，车辆已释放回在库'})


@app.route('/api/order-refunds/<int:rid>/cancel', methods=['POST'])
@require_role('老板')
def cancel_order_refund(rid):
    """老板兜底取消退款：不退款，直接释放车辆回在库。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM order_refunds WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退款单不存在'}), 404
    if row['status'] != '待退款':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能取消'}), 400
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("UPDATE order_refunds SET status='已取消', executed_by=?, executed_at=? WHERE id=?",
              (request.current_user['display_name'], now, rid))
    release_vids = []
    if row['vehicle_id']:
        release_vids = [row['vehicle_id']]
    if row['sales_order_id']:
        so = c.execute("SELECT vehicle_id, vehicle_ids FROM sales_orders WHERE id=?", (row['sales_order_id'],)).fetchone()
        if so:
            release_vids = order_vehicle_ids(conn, so) or release_vids
    for vid in release_vids:
        c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (vid,))
    if row['contract_id']:
        c.execute("""
            UPDATE contracts SET contract_status='已终止', delivery_status='已终止'
            WHERE id=? AND contract_status='报单计划中'
        """, (row['contract_id'],))
    log_audit(conn, '取消报单退款', 'order_refund', rid,
              f'报单{row["sales_order_id"]} 取消退款，车辆释放回在库', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '退款已取消，车辆已释放回在库'})


# 文件上传
@app.route('/api/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有文件'}), 400
    f = request.files['file']
    if not f.filename:
        return jsonify({'success': False, 'message': '文件名为空'}), 400
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ['.jpg', '.jpeg', '.png', '.pdf', '.doc', '.docx', '.xls', '.xlsx']:
        return jsonify({'success': False, 'message': '不支持的文件格式'}), 400
    filename = f"{uuid.uuid4().hex}{ext}"
    filepath = os.path.join(UPLOAD_DIR, filename)
    f.save(filepath)
    return jsonify({'success': True, 'url': f'/uploads/{filename}', 'filename': f.filename})

@app.route('/uploads/<filename>')
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename)


@app.route('/api/vehicles', methods=['POST'])
@require_role('车管')
def add_vehicle():
    data = request.json or {}
    vin = (data.get('vin') or '').strip().upper()
    if len(vin) != 17:
        return jsonify({'success': False, 'message': '请填写17位VIN'}), 400

    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("SELECT id, status FROM vehicles WHERE vin=?", (vin,))
        existing = c.fetchone()
        if existing:
            return jsonify({'success': False, 'message': '车辆已在库'}), 400

        # 手工录入按能源类型校验规格；尾板缺省即“无”，不会再生成未区分尾板的基准车型。
        normalize_base_model_tailgate(data)
        base_errors = validate_manual_vehicle_base_model(data)
        if base_errors:
            return jsonify({'success': False, 'message': '；'.join(base_errors)}), 400

        # 车型是基准车型的计算字段；旧数据未提供任何规格时兼容保留传入车型。
        if has_base_model_input(data):
            data['car_type'] = compute_car_type(data)
        car_type = (data.get('car_type') or '').strip()

        # 判断该车型是否已维护指导价（model_guidance_prices 字典）
        guidance_price = 0
        if car_type:
            model_price = c.execute(
                "SELECT guidance_price FROM model_guidance_prices WHERE car_type IN ({}) LIMIT 1".format(
                    ','.join('?' for _ in guidance_base_candidates(car_type))
                ),
                guidance_base_candidates(car_type)
            ).fetchone()
            if model_price and parse_money(model_price['guidance_price']) > 0:
                guidance_price = parse_money(model_price['guidance_price'])

        # 入库字段：新模板 25 列全部维度
        base_cols = ['vin', 'plate_number', 'company', 'car_type',
                     'purchase_price', 'tax_rate', 'estimated_residual_value',
                     'insurance_expiry_date', 'annual_review_date',
                     'vehicle_category', 'vehicle_engine_battery', 'vehicle_power_battery',
                     'vehicle_color', 'vehicle_box_type',
                     'box_type_remark', 'invoice_contract_file', 'status',
                     # 新模板维度
                     'condition', 'brand', 'product_series', 'battery_capacity',
                     'horsepower', 'box_type', 'box_dimension', 'gear_position',
                     'tailgate', 'battery_brand', 'product_code', 'cab_type',
                     'other_config', 'fuel_form', 'cab_style', 'engine_spec',
                     'gearbox_spec', 'drive_motor_model', 'battery_model', 'suspension_model']
        all_cols = base_cols
        placeholders = ', '.join('?' for _ in all_cols)
        base_vals = [
            vin, data.get('plate_number'), data.get('company', '陕西金聚源汽车服务有限公司'),
            car_type,
            parse_money(data.get('purchase_price') or data.get('网员价'), 0),
            data.get('tax_rate', 0.13),
            data.get('estimated_residual_value', 0),
            data.get('insurance_expiry_date', ''), data.get('annual_review_date', ''),
            data.get('vehicle_category', ''),
            data.get('vehicle_engine_battery', ''),
            data.get('vehicle_power_battery', ''),
            data.get('vehicle_color', ''),
            data.get('vehicle_box_type', ''),
            data.get('box_type_remark', ''),
            data.get('invoice_contract_file', ''),
            data.get('status', '在库'),
            # 新模板维度
            data.get('condition', ''), data.get('brand', ''),
            data.get('product_series', ''), data.get('battery_capacity', ''),
            data.get('horsepower', ''), data.get('box_type', ''),
            data.get('box_dimension', ''), data.get('gear_position', ''),
            data.get('tailgate', ''), data.get('battery_brand', ''),
            data.get('product_code', ''), data.get('cab_type', ''),
            data.get('other_config', ''), data.get('fuel_form', ''),
            data.get('cab_style', ''), data.get('engine_spec', ''),
            data.get('gearbox_spec', ''), data.get('drive_motor_model', ''),
            data.get('battery_model', ''), data.get('suspension_model', ''),
        ]
        # 字典校验（SKU 改造：传 vehicle 做维度值字典检查）
        val_status, val_msg = validate_vehicle_dict(conn, car_type, data)
        all_cols = all_cols + ['validation_status', 'validation_message']
        base_vals = base_vals + [val_status, val_msg]
        placeholders = ', '.join('?' for _ in all_cols)
        c.execute(f'''
        INSERT INTO vehicles ({', '.join(all_cols)})
        VALUES ({placeholders})
        ''', base_vals)
        vehicle_id = c.lastrowid
        # SKU 改造：入库后自动关联/创建 SKU
        try:
            ensure_sku_for_vehicle(conn, data, request.current_user['display_name'] if getattr(request, 'current_user', None) else '系统')
        except Exception:
            pass
        if guidance_price <= 0:
            log_audit(conn, '缺少指导价提醒', 'vehicle', vehicle_id,
                      f'新入库车辆 VIN:{vin} 车型:{car_type or "-"} 未设置指导价，请老板维护指导价')
        conn.commit()
        message = '车辆入库成功'
        if guidance_price <= 0:
            message = '车辆入库成功，该车未设置指导价，已提醒老板维护'
        if val_status == 'invalid':
            message += f'；{val_msg}'
        return jsonify({'success': True, 'id': vehicle_id, 'message': message, 'missing_guidance_price': guidance_price <= 0})
    except Exception as e:
        conn.rollback()
        message = '车辆已在库' if 'UNIQUE constraint failed: vehicles.vin' in str(e) else str(e)
        return jsonify({'success': False, 'message': message}), 400
    finally:
        conn.close()

# 新车 Excel 批量上传入库（车管入库信息表 — 79列模板）
# 表头位于第2行（openpyxl 1-based），数据自第3行起
# 按表头中文名匹配，列顺序变化也不影响
VEHICLE_HEADER_MAP = {
    # === 20260804 最新模板（最终定稿）：车辆入库导入-月份发票 25列 ===
    '成色': 'condition', '品牌': 'brand', '品系': 'product_series',
    '电池度数': 'battery_capacity', '马力': 'horsepower', '厢型': 'box_type',
    '厢尺寸': 'box_dimension', '档位': 'gear_position', '尾板': 'tailgate',
    '颜色': 'vehicle_color', '电池品牌': 'battery_brand',
    '车型': 'car_type', '车牌号': 'plate_number',
    '产品代码': 'product_code', 'VIN': 'vin', 'VIN码': 'vin',
    '网员价': 'purchase_price',
    '驾驶室': 'cab_type', '其他': 'other_config', '燃料形式': 'fuel_form',
    '驾驶室类型': 'cab_style',
    '发动机厂家及功率': 'engine_spec', '变速箱厂家及型号': 'gearbox_spec',
    '驱动电机型号': 'drive_motor_model', '新能源动力电池型号': 'battery_model',
    '悬架型号': 'suspension_model',
    # === 旧模板兼容 ===
    '产品名称': 'car_type',
    '结算价格': 'purchase_price', '结算价': 'purchase_price', '购买价': 'purchase_price',
    '技术路线': 'tech_route',
    '经销商代码': 'dealer_code', '经销商编号': 'dealer_code',
}

# 基准车型计算源字段。成色、厢型、厢尺寸均属于子型号或附加描述，不参与基准车型。
CAR_TYPE_SOURCE_FIELDS = [
    'fuel_form', 'brand', 'product_series', 'battery_brand',
    'battery_capacity', 'horsepower', 'gear_position', 'tailgate',
]


def compute_car_type(v):
    """按能源类型生成基准车型。

    纯电：品牌 + 品系 + 电池品牌 + 电池度数 + 尾板
    混动/燃油车：品牌 + 品系 + 马力 + 档位 + 尾板
    """
    energy_type = energy_type_for_vehicle(v)
    fields = BASE_MODEL_FIELDS_BY_ENERGY.get(energy_type, ())
    if not fields:
        return ''
    parts = [str(v.get(field, '') or '').strip() for field in fields]
    if not all(parts):
        return ''
    return ''.join(parts) + base_model_tailgate_label(v.get('tailgate'))


def derive_battery_capacity(battery_model):
    """电池度数：从电池型号 kWh 值向下取整 + 度（如 61.1kWh -> 61度）"""
    if not battery_model:
        return ''
    m = re.search(r'(\d+(?:\.\d+)?)\s*kWh', str(battery_model), re.IGNORECASE)
    if not m:
        return ''
    try:
        return str(int(float(m.group(1)))) + '度'
    except (TypeError, ValueError):
        return ''


def derive_battery_brand(battery_model):
    """电池品牌：电池型号前 2 个字符（力神/宁德/弗迪）"""
    if not battery_model:
        return ''
    return str(battery_model).strip()[:2]


def derive_vehicle_computed(row_vals):
    """导入时补充电池规格；规格字段齐全时按能源口径重算基准车型，
    规格不齐时保留 Excel 的「车型」列原值（不再覆盖为空）。"""
    vals = dict(row_vals)
    battery_model = vals.get('battery_model') or ''
    if not vals.get('battery_capacity') and battery_model:
        vals['battery_capacity'] = derive_battery_capacity(battery_model)
    if not vals.get('battery_brand') and battery_model:
        vals['battery_brand'] = derive_battery_brand(battery_model)
    normalize_base_model_tailgate(vals)
    if has_base_model_input(vals):
        computed = compute_car_type(vals)
        if computed:
            vals['car_type'] = computed
    return vals


# ======================== SKU 关联（SKU 改造）====================
def sku_key_for_vehicle(v):
    """计算车辆对应 SKU 组合键；底盘车（纯车头）不建 SKU，返回 None。

    组合 = 基准车型（去成色前缀+去厢型后缀）+ 成色 + 厢型 + 尾板（空→无）。
    """
    if not v:
        return None
    box_type = (v.get('box_type') or v.get('vehicle_box_type') or '').strip()
    if box_type == '底盘':
        return None
    car_type = (v.get('car_type') or '').strip()
    if not car_type:
        return None
    base = normalize_base_car_type(car_type)
    if not base:
        return None
    condition = (v.get('condition') or '新车').strip() or '新车'
    if condition not in ('新车', '二手车'):
        condition = '新车'
    tailgate = normalize_tailgate(v.get('tailgate') or v.get('tail_plate'))
    return {'car_type': base, 'condition': condition, 'box_type': box_type, 'tailgate': tailgate}


def ensure_sku_for_vehicle(conn, vehicle, operator='系统'):
    """车辆入库/改维度后，自动创建/复用对应 SKU（INSERT OR IGNORE）。
    返回 sku_id 或 None（底盘车/缺车型不建 SKU）。"""
    key = sku_key_for_vehicle(vehicle)
    if not key:
        return None
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute("""
        INSERT OR IGNORE INTO skus (car_type, condition, box_type, tailgate, status, created_by, updated_by, updated_at)
        VALUES (?, ?, ?, ?, '生效', ?, ?, ?)
    """, (key['car_type'], key['condition'], key['box_type'], key['tailgate'],
          operator, operator, now))
    row = conn.execute(
        "SELECT id FROM skus WHERE car_type=? AND condition=? AND box_type=? AND tailgate=? LIMIT 1",
        (key['car_type'], key['condition'], key['box_type'], key['tailgate'])
    ).fetchone()
    return row['id'] if row else None


def _excel_cell_text(value):
    """把单元格值规整为去空白的字符串；日期取 YYYY-MM-DD。"""
    if value is None:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if hasattr(value, 'strftime'):
        try:
            return value.strftime('%Y-%m-%d')
        except Exception:
            pass
    return str(value).strip()


@app.route('/api/vehicles/import', methods=['POST'])
def import_vehicles():
    """上传经销商买断库存 Excel，批量入库车辆。
    注意：所有错误都返回 200（success=false），避免 el-upload on-error 无法拿到错误消息。"""
    # 手动鉴权，不依赖 require_role（否则 401/403 会触发 el-upload on-error）
    user = get_current_user()
    if not user:
        return jsonify({'success': False, 'message': '登录已失效，请刷新页面'}), 200
    if user['role'] not in ('车管', '老板'):
        return jsonify({'success': False, 'message': '仅车管角色可以批量导入'}), 200
    request.current_user = user
    file_url = ''
    if request.is_json:
        file_url = ((request.json or {}).get('file_url') or '').strip()
    if not file_url:
        file_url = (request.form.get('file_url') or '').strip()

    if file_url:
        local_path = upload_url_to_path(file_url)
    elif 'file' in request.files:
        f = request.files['file']
        if not f.filename:
            return jsonify({'success': False, 'message': '文件名为空'}), 200
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ('.xls', '.xlsx'):
            return jsonify({'success': False, 'message': '仅支持 xls / xlsx 格式'}), 200
        local_path = os.path.join(UPLOAD_DIR, uuid.uuid4().hex + ext)
        f.save(local_path)
    else:
        return jsonify({'success': False, 'message': '请上传车辆信息 Excel'}), 200

    if not local_path or not os.path.exists(local_path):
        return jsonify({'success': False, 'message': '导入文件不存在，请重新上传'}), 200
    if not local_path.lower().endswith(('.xls', '.xlsx')):
        return jsonify({'success': False, 'message': '仅支持 xls / xlsx 格式'}), 200

    try:
        wb = load_workbook(local_path, data_only=True)
        ws = wb.active
    except Exception as e:
        return jsonify({'success': False, 'message': '解析 Excel 失败：' + str(e)}), 200

    # 自动定位表头行：在前若干行中查找包含 "VIN" 的行
    # （20260804 新模板：第1行分组说明、第2行表头、第3行起数据）
    header_row = None
    header = []
    for r in range(1, min(ws.max_row, 10) + 1):
        values = [_excel_cell_text(cell.value) for cell in ws[r]]
        joined = ''.join(values).upper()
        if 'VIN' in joined and ('成色' in values or '品牌' in values or '产品名称' in values or '经销商' in values):
            header_row = r
            header = values
            break
    if header_row is None:
        header_row = 4
        header = [_excel_cell_text(cell.value) for cell in ws[header_row]]

    # 按表头中文名建立列映射（列顺序变化也不影响）
    col_map = {}
    for col_idx, h in enumerate(header):
        h_clean = h.replace(' ', '').replace('\u3000', '').replace('\uff1a', ':')
        if h_clean in VEHICLE_HEADER_MAP:
            col_map[VEHICLE_HEADER_MAP[h_clean]] = col_idx
        for cn_name, field in VEHICLE_HEADER_MAP.items():
            if field not in col_map and cn_name in h_clean:
                col_map[field] = col_idx

    if 'vin' not in col_map:
        for i, h in enumerate(header):
            if 'VIN' in h.upper():
                col_map['vin'] = i
                break

    conn = get_db()
    c = conn.cursor()

    imported, skipped, failed = 0, 0, 0
    details = []
    missing_guidance_count = 0
    dictionary_invalid_count = 0

    try:
        for r in range(header_row + 1, ws.max_row + 1):
            cells = ws[r]
            row_vals = {}
            for field, col_idx in col_map.items():
                cell = cells[col_idx] if col_idx < len(cells) else None
                if cell is not None and cell.value is not None:
                    if field == 'vin' and isinstance(cell.value, datetime):
                        row_vals[field] = ''
                    else:
                        row_vals[field] = _excel_cell_text(cell.value)
                else:
                    row_vals[field] = ''

            vin = (row_vals.get('vin') or '').strip().upper()
            if not vin:
                continue
            if len(vin) != 17:
                failed += 1
                details.append({'row': r, 'vin': vin, 'status': 'failed', 'reason': 'VIN 非17位'})
                continue

            c.execute("SELECT id FROM vehicles WHERE vin=?", (vin,))
            if c.fetchone():
                skipped += 1
                details.append({'row': r, 'vin': vin, 'status': 'skipped', 'reason': '车辆已在库'})
                continue

            raw = {}
            for idx in range(len(header)):
                key = header[idx] if (idx < len(header) and header[idx]) else ('col' + str(idx))
                cell = cells[idx] if idx < len(cells) else None
                raw[key] = _excel_cell_text(cell.value) if cell is not None else ''

            # 补充计算字段：电池度数/电池品牌（Excel 公式模拟）+ 车型拼接
            row_vals = derive_vehicle_computed(row_vals)

            purchase_price = parse_money(row_vals.get('purchase_price') or row_vals.get('settlement_price') or row_vals.get('invoice_price'), 0)
            car_type = row_vals.get('car_type') or row_vals.get('product_name') or ''

            # 判断该车型是否已维护指导价（model_guidance_prices 字典）
            guidance_price = 0
            if car_type:
                mp = c.execute(
                    "SELECT guidance_price FROM model_guidance_prices WHERE car_type IN ({}) LIMIT 1".format(
                        ','.join('?' for _ in guidance_base_candidates(car_type))
                    ),
                    guidance_base_candidates(car_type)
                ).fetchone()
                if mp and parse_money(mp['guidance_price']) > 0:
                    guidance_price = parse_money(mp['guidance_price'])

            # === 字典校验（SKU 改造：传 row_vals 做维度值字典检查）===
            imp_val_status, imp_val_msg = validate_vehicle_dict(conn, car_type, row_vals)
            if imp_val_status == 'invalid':
                dictionary_invalid_count += 1

            # 入库字段：新模板 25 列全部维度
            VINSERT = [
                ('vin', vin), ('company', '陕西金聚源汽车服务有限公司'),
                ('car_type', car_type), ('plate_number', row_vals.get('plate_number') or ''),
                # 可编辑维度（A-M）
                ('condition', row_vals.get('condition') or ''),
                ('brand', row_vals.get('brand') or ''),
                ('product_series', row_vals.get('product_series') or ''),
                ('battery_capacity', row_vals.get('battery_capacity') or ''),
                ('horsepower', row_vals.get('horsepower') or ''),
                ('box_type', row_vals.get('box_type') or ''),
                ('box_dimension', row_vals.get('box_dimension') or ''),
                ('gear_position', row_vals.get('gear_position') or ''),
                ('tailgate', row_vals.get('tailgate') or ''),
                ('vehicle_color', row_vals.get('vehicle_color') or ''),
                ('battery_brand', row_vals.get('battery_brand') or ''),
                # 不可编辑维度（N-Y）
                ('product_code', row_vals.get('product_code') or ''),
                ('purchase_price', purchase_price),
                ('cab_type', row_vals.get('cab_type') or ''),
                ('other_config', row_vals.get('other_config') or ''),
                ('fuel_form', row_vals.get('fuel_form') or ''),
                ('cab_style', row_vals.get('cab_style') or ''),
                ('engine_spec', row_vals.get('engine_spec') or ''),
                ('gearbox_spec', row_vals.get('gearbox_spec') or ''),
                ('drive_motor_model', row_vals.get('drive_motor_model') or ''),
                ('battery_model', row_vals.get('battery_model') or ''),
                ('suspension_model', row_vals.get('suspension_model') or ''),
                # 旧模板兼容
                ('tech_route', row_vals.get('tech_route') or ''),
                ('dealer_code', row_vals.get('dealer_code') or ''),
                ('status', '在库'),
                ('import_raw', json.dumps(raw, ensure_ascii=False)),
                ('validation_status', imp_val_status),
                ('validation_message', imp_val_msg),
            ]
            cols = ', '.join(v[0] for v in VINSERT)
            placeholders = ', '.join('?' for _ in VINSERT)
            vals = tuple(v[1] for v in VINSERT)

            try:
                c.execute(
                    f"INSERT INTO vehicles ({cols}) VALUES ({placeholders})",
                    vals
                )
                vehicle_id = c.lastrowid
                # SKU 改造：导入入库后自动关联/创建 SKU
                try:
                    ensure_sku_for_vehicle(conn, row_vals, request.current_user['display_name'])
                except Exception:
                    pass
                if guidance_price <= 0:
                    missing_guidance_count += 1
                    log_audit(conn, '缺少指导价提醒', 'vehicle', vehicle_id,
                              '上传入库车辆 VIN:' + vin + ' 车型:' + (car_type or '-') + ' 未设置指导价，请老板维护指导价')
                imported += 1
                details.append({'row': r, 'vin': vin, 'status': 'imported', 'car_type': car_type,
                                'validation_status': imp_val_status, 'validation_message': imp_val_msg})
            except Exception as e:
                failed += 1
                msg = '车辆已在库' if 'UNIQUE constraint failed' in str(e) else str(e)
                details.append({'row': r, 'vin': vin, 'status': 'failed', 'reason': msg})

        log_audit(conn, '上传入库', 'vehicle', None,
                  'Excel批量入库：成功' + str(imported) + ' 重复' + str(skipped) + ' 失败' + str(failed),
                  request.current_user['display_name'])
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': '导入失败：' + str(e)}), 200
    finally:
        try:
            conn.close()
        except Exception:
            pass

    message = '上传入库完成：成功 ' + str(imported) + ' 台，重复跳过 ' + str(skipped) + ' 台，失败 ' + str(failed) + ' 台'
    if missing_guidance_count:
        message += '；其中 ' + str(missing_guidance_count) + ' 台未匹配指导价，已提醒老板维护'
    if dictionary_invalid_count:
        message += '；其中 ' + str(dictionary_invalid_count) + ' 台车型不在指导价字典中，请及时维护'
    return jsonify({
        'success': True,
        'message': message,
        'imported': imported,
        'skipped': skipped,
        'failed': failed,
        'missing_guidance': missing_guidance_count,
        'dictionary_invalid': dictionary_invalid_count,
        'details': details,
    })


# 车辆维度编辑权限：老板可编辑所有字段；运营仅可编辑 12 个可编辑维度；其他角色无编辑权限
BOSS_VEHICLE_EDITABLE = [
    'vin', 'plate_number', 'car_type', 'condition', 'brand', 'product_series', 'battery_capacity',
    'horsepower', 'box_type', 'gear_position', 'tailgate',
    'vehicle_color', 'battery_brand',
    'product_code', 'purchase_price', 'cab_type', 'other_config',
    'cab_style', 'engine_spec', 'gearbox_spec', 'drive_motor_model',
    'battery_model', 'suspension_model',
    'company', 'tax_rate', 'invoice_contract_file', 'status',
    'insurance_expiry_date', 'annual_review_date', 'vehicle_category',
    'vehicle_engine_battery', 'vehicle_power_battery', 'vehicle_box_type',
    'box_type_remark', 'tech_route', 'dealer_code',
]
OPS_VEHICLE_EDITABLE = [
    'condition', 'brand', 'product_series', 'battery_capacity', 'horsepower',
    'box_type', 'gear_position', 'tailgate',
    'vehicle_color', 'battery_brand', 'plate_number',
]
# 运营不可编辑车型/VIN；老板可手动维护（如二手车无车型时）


@app.route('/api/vehicles/<int:vid>', methods=['PUT'])
@require_role('运营', '老板')
def update_vehicle(vid):
    data = request.json or {}
    role = request.current_user['role']
    conn = get_db()
    c = conn.cursor()

    if role == '老板':
        editable = BOSS_VEHICLE_EDITABLE
    else:
        editable = OPS_VEHICLE_EDITABLE

    # 前端「网员价」提交为 dealer_price → 写入 purchase_price（网员价=购买价）
    if 'dealer_price' in data and role == '老板':
        data['purchase_price'] = data['dealer_price']
    # 只接受当前角色被授权维护的字段。燃料形式来自采购导入，是能源规则
    # 的只读依据，不能通过构造 PUT 请求把车辆切换到另一种能源类型。
    data = {key: data[key] for key in editable if key in data}

    existing = c.execute("SELECT * FROM vehicles WHERE id=?", (vid,)).fetchone()
    if not existing:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    candidate = dict(existing)
    for key in editable:
        if key in data:
            candidate[key] = data[key]

    # 成色/厢型是子型号字段，不触发基准车型重算；燃料和基准规格字段才会。
    src_changed = any(k in data for k in CAR_TYPE_SOURCE_FIELDS)
    if src_changed:
        normalize_base_model_tailgate(candidate)
        base_errors = validate_manual_vehicle_base_model(candidate)
        if base_errors:
            conn.close()
            return jsonify({'success': False, 'message': '；'.join(base_errors)}), 400
        # 规格变更时把尾板规范值一并落库，避免继续产生无尾板标识缺失的基准车型。
        data['tailgate'] = candidate['tailgate']

    fields = []
    values = []
    for key in editable:
        if key in data:
            fields.append(f"{key} = ?")
            values.append(data[key])
    if fields:
        values.append(vid)
        c.execute(f"UPDATE vehicles SET {', '.join(fields)} WHERE id = ?", values)

    # 手动维护的车型（老板直改 car_type）同样要走字典校验，保证报单不被漏放
    if 'car_type' in data and not src_changed:
        candidate['car_type'] = data['car_type']
        val_status, val_msg = validate_vehicle_dict(conn, candidate['car_type'], candidate)
        c.execute("UPDATE vehicles SET validation_status=?, validation_message=? WHERE id=?",
                  (val_status, val_msg, vid))

    # 若修改了基准车型源字段 → 自动重算车型 + 重新字典校验。
    if src_changed:
        new_car_type = compute_car_type(candidate)
        candidate['car_type'] = new_car_type
        c.execute("UPDATE vehicles SET car_type=? WHERE id=?", (new_car_type, vid))
        val_status, val_msg = validate_vehicle_dict(conn, new_car_type, candidate)
        c.execute("UPDATE vehicles SET validation_status=?, validation_message=? WHERE id=?",
                  (val_status, val_msg, vid))
    elif any(k in data for k in ('condition', 'box_type')):
        # 子型号变化不改变基准车型，但仍需重建 SKU 归属。
        candidate['car_type'] = existing['car_type']

    if src_changed or any(k in data for k in ('condition', 'box_type')):
        try:
            ensure_sku_for_vehicle(conn, candidate, request.current_user['display_name'])
        except Exception:
            pass
    conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/vehicles/revalidate-all', methods=['POST'])
@require_role('车管', '老板')
def revalidate_all_vehicles():
    """对所有库存车辆重新执行字典校验（新增字段时用于回填）"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM vehicles WHERE COALESCE(is_deleted,0)=0")
    rows = c.fetchall()
    updated = 0
    for row in rows:
        vehicle = dict(row)
        status, msg = validate_vehicle_dict(conn, vehicle['car_type'], vehicle)
        c.execute("UPDATE vehicles SET validation_status=?, validation_message=? WHERE id=?",
                  (status, msg, row['id']))
        updated += 1
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'已完成 {updated} 台车辆的字典校验'})


@app.route('/api/vehicles/<int:vid>', methods=['DELETE'])
@require_role('老板')
def delete_vehicle(vid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vin, plate_number FROM vehicles WHERE id=?", (vid,))
    v = c.fetchone()
    if not v:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    # 软删除：标记为已删除，数据保留
    c.execute("UPDATE vehicles SET is_deleted = 1 WHERE id = ?", (vid,))
    log_audit(conn, '软删除车辆', 'vehicle', vid,
              f'{request.current_user["display_name"]} 软删除了车辆 VIN:{v["vin"]} {v["plate_number"] or ""}',
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已隐藏，所有关联数据保留'})

@app.route('/api/vehicles/<int:vid>/restore', methods=['POST'])
@require_role('老板')
def restore_vehicle(vid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vin, plate_number FROM vehicles WHERE id=?", (vid,))
    v = c.fetchone()
    if not v:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    c.execute("UPDATE vehicles SET is_deleted = 0 WHERE id = ?", (vid,))
    log_audit(conn, '恢复车辆', 'vehicle', vid,
              f'{request.current_user["display_name"]} 恢复了车辆 VIN:{v["vin"]} {v["plate_number"] or ""}',
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已恢复'})

@app.route('/api/vehicles/<int:vid>/guidance_price', methods=['POST'])
@require_role('老板')
def update_guidance_price(vid):
    data = request.json or {}
    new_price = parse_money(data.get('price'))
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    vehicle = c.execute("SELECT id, car_type FROM vehicles WHERE id=?", (vid,)).fetchone()
    if not vehicle:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    old_hist = c.execute("""
        SELECT new_price FROM vehicle_guidance_price_history
        WHERE vehicle_id=? ORDER BY effective_at DESC, id DESC LIMIT 1
    """, (vid,)).fetchone()
    old_price = old_hist['new_price'] if old_hist else 0
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if new_price <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '指导价必须大于0'}), 400
    c.execute("""
        INSERT INTO vehicle_guidance_price_history
            (vehicle_id, old_price, new_price, changed_by, effective_at)
        VALUES (?, ?, ?, ?, ?)
    """, (vid, old_price, new_price, user['display_name'], now))
    log_audit(conn, '更新指导价', 'vehicle', vid,
              f'{user["display_name"]} 将指导价由 ¥{old_price} 调整为 ¥{new_price}，自 {now} 后生效',
              user['display_name'])
    remaining_missing_count = unresolved_guidance_vehicle_count(conn)
    conn.commit()
    conn.close()
    return jsonify({
        'success': True,
        'message': '指导价更新成功，后续成交将按新指导价判断',
        'remaining_missing_guidance_count': remaining_missing_count,
    })

@app.route('/api/customer-blacklist', methods=['GET'])
@require_role('财务', '老板')
def list_customer_blacklist():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM customer_blacklist ORDER BY id DESC")
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/customer-blacklist', methods=['POST'])
@require_role('财务', '老板')
def upsert_customer_blacklist():
    data = request.json or {}
    customer_name = (data.get('customer_name') or '').strip()
    customer_phone = (data.get('customer_phone') or '').strip()
    if not customer_name and not customer_phone:
        return jsonify({'success': False, 'message': '客户姓名或电话至少填写一项'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO customer_blacklist
            (customer_name, customer_phone, id_card, level, reason, status, created_by)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        customer_name,
        customer_phone,
        (data.get('id_card') or '').strip(),
        data.get('level') or '禁止报单',
        (data.get('reason') or '').strip(),
        data.get('status') or '生效',
        request.current_user['display_name'],
    ))
    blacklist_id = c.lastrowid
    log_audit(conn, '维护客户黑名单', 'customer_blacklist', blacklist_id,
              f"{customer_name or customer_phone} {data.get('reason') or ''}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': blacklist_id, 'message': '黑名单已保存'})


@app.route('/api/receiving-companies', methods=['GET'])
@login_required
def list_receiving_companies():
    include_disabled = request.args.get('include_disabled') == '1'
    conn = get_db()
    c = conn.cursor()
    if include_disabled and request.current_user['role'] in ('财务', '老板'):
        c.execute("SELECT * FROM receiving_companies ORDER BY status ASC, company_name ASC")
    else:
        c.execute("SELECT * FROM receiving_companies WHERE status='启用' ORDER BY company_name ASC")
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/receiving-companies', methods=['POST'])
@require_role('财务', '老板')
def upsert_receiving_company():
    data = request.json or {}
    company_name = (data.get('company_name') or data.get('name') or '').strip()
    if not company_name:
        return jsonify({'success': False, 'message': '收款公司名称不能为空'}), 400
    status = data.get('status') or '启用'
    if status not in ('启用', '停用'):
        return jsonify({'success': False, 'message': '状态只能为启用或停用'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("""
            INSERT INTO receiving_companies
                (company_name, bank_name, bank_account_no, tax_no, status, remark, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(company_name) DO UPDATE SET
                bank_name=excluded.bank_name,
                bank_account_no=excluded.bank_account_no,
                tax_no=excluded.tax_no,
                status=excluded.status,
                remark=excluded.remark,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at
        """, (
            company_name,
            (data.get('bank_name') or '').strip(),
            (data.get('bank_account_no') or '').strip(),
            (data.get('tax_no') or '').strip(),
            status,
            (data.get('remark') or '').strip(),
            request.current_user['display_name'],
            now,
        ))
        c.execute("SELECT id FROM receiving_companies WHERE company_name=?", (company_name,))
        row = c.fetchone()
        company_id = row['id'] if row else None
        log_audit(conn, '维护收款公司', 'receiving_company', company_id,
                  f'{company_name} 状态:{status}', request.current_user['display_name'])
        conn.commit()
        return jsonify({'success': True, 'id': company_id, 'message': '收款公司已保存'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


@app.route('/api/vehicle-rebates', methods=['GET'])
@require_role('财务', '老板')
def list_vehicle_rebates():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT vr.*, v.vin, v.plate_number, v.car_type
        FROM vehicle_rebates vr
        JOIN vehicles v ON v.id = vr.vehicle_id
        WHERE (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY vr.id DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/vehicle-rebates', methods=['POST'])
@require_role('财务')
def create_vehicle_rebate():
    data = request.json or {}
    vehicle_id = data.get('vehicle_id')
    amount = parse_money(data.get('rebate_amount'))
    if not vehicle_id or amount <= 0:
        return jsonify({'success': False, 'message': '请选择车辆并填写大于0的返利金额'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO vehicle_rebates
            (vehicle_id, contract_id, rebate_amount, rebate_date, remark, created_by)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        vehicle_id,
        data.get('contract_id'),
        amount,
        data.get('rebate_date') or datetime.now().strftime('%Y-%m-%d'),
        (data.get('remark') or '').strip(),
        request.current_user['display_name'],
    ))
    rebate_id = c.lastrowid
    log_audit(conn, '录入车辆返利', 'vehicle_rebate', rebate_id,
              f"车辆{vehicle_id} 返利{amount}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': rebate_id, 'message': '返利已录入'})


@app.route('/api/invoice-requests', methods=['GET'])
@require_role('运营', '财务', '老板')
def list_invoice_requests():
    conn = get_db()
    c = conn.cursor()
    status = request.args.get('status')
    if status:
        c.execute("""
            SELECT ir.*, c.contract_type, cu.name AS customer_name
            FROM invoice_requests ir
            JOIN contracts c ON c.id = ir.contract_id
            LEFT JOIN customers cu ON cu.id = c.customer_id
            WHERE ir.status=?
            ORDER BY ir.id DESC
        """, (status,))
    else:
        c.execute("""
            SELECT ir.*, c.contract_type, cu.name AS customer_name
            FROM invoice_requests ir
            JOIN contracts c ON c.id = ir.contract_id
            LEFT JOIN customers cu ON cu.id = c.customer_id
            ORDER BY ir.id DESC
        """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/contracts/<int:cid>/invoices', methods=['POST'])
@require_role('运营')
def create_invoice_request(cid):
    data = request.json or {}
    amount = parse_money(data.get('amount'))
    if amount <= 0:
        return jsonify({'success': False, 'message': '开票金额必须大于0'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, company FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404

    receiving_company = (data.get('receiving_company') or contract['company'] or '').strip()
    invoice_entity_name = (data.get('invoice_entity_name') or '').strip()
    if not invoice_entity_name:
        conn.close()
        return jsonify({'success': False, 'message': '开票抬头不能为空'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        INSERT INTO invoice_requests
            (contract_id, period, amount, receiving_company, title_type,
             invoice_entity_name, invoice_entity_tax_no, status, applied_by, applied_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, '待审批', ?, ?)
    """, (
        cid,
        data.get('period'),
        amount,
        receiving_company,
        data.get('title_type') or '企业',
        invoice_entity_name,
        (data.get('invoice_entity_tax_no') or '').strip(),
        request.current_user['display_name'],
        now,
    ))
    invoice_id = c.lastrowid
    # 老板审批环节：创建审批流（老板通过后进入待开票，财务开票）
    create_approval_flow(conn, 'invoice', invoice_id)
    log_audit(conn, '发起发票申请', 'invoice_request', invoice_id,
              f'合同{cid} 金额{amount} 抬头:{invoice_entity_name}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': invoice_id, 'message': '发票申请已提交，等待老板审批'})


@app.route('/api/invoice-requests/<int:iid>/approve', methods=['POST'])
@require_role('老板')
def boss_approve_invoice_request(iid):
    """老板通过发票申请，进入待开票，财务可开票。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] != '待审批':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能审批'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = request.current_user['display_name']
    # 审批流置通过
    c.execute("""
        UPDATE approval_flows SET status='已通过', operator_name=?, comment='老板通过发票审批', acted_at=?
        WHERE ref_type='invoice' AND ref_id=? AND status='待审批'
    """, (user, now, iid))
    # 发票进入待开票
    c.execute("""
        UPDATE invoice_requests
        SET status='待开票',
            boss_approved_by=?,
            boss_approved_at=?
        WHERE id=?
    """, (user, now, iid))
    log_audit(conn, '发票审批通过', 'invoice_request', iid,
              f'老板通过，进入待开票', user)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票审批通过，等待财务开票'})


@app.route('/api/invoice-requests/<int:iid>/reject', methods=['POST'])
@require_role('老板')
def boss_reject_invoice_request(iid):
    """老板驳回发票申请（与各端取消同一终止态：已驳回）。"""
    data = request.json or {}
    reason = (data.get('reason') or '').strip()
    if not reason:
        conn.close()
        return jsonify({'success': False, 'message': '驳回原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] not in ('待审批', '待开票'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能驳回'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = request.current_user['display_name']
    # 审批流置取消（含待审批/待开票阶段的老板驳回）
    c.execute("""
        UPDATE approval_flows SET status='已取消', operator_name=?, comment=?, acted_at=?
        WHERE ref_type='invoice' AND ref_id=? AND status='待审批'
    """, (user, f'老板驳回：{reason}', now, iid))
    # 发票置已驳回（终止态）
    c.execute("""
        UPDATE invoice_requests
        SET status='已驳回',
            voided_by=?,
            voided_at=?,
            void_reason=?
        WHERE id=?
    """, (user, now, reason, iid))
    log_audit(conn, '发票审批驳回', 'invoice_request', iid, reason, user)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票申请已驳回'})


@app.route('/api/invoice-requests/<int:iid>/cancel', methods=['POST'])
@require_role('运营', '老板', '财务')
def cancel_invoice_request(iid):
    """各端（运营/老板/财务）取消发票申请，统一终止态：已驳回。"""
    data = request.json or {}
    reason = (data.get('reason') or '').strip()
    if not reason:
        conn.close()
        return jsonify({'success': False, 'message': '取消原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] not in ('待审批', '待开票'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能取消（已开票请走红冲）'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = request.current_user['display_name']
    # 审批流置取消
    c.execute("""
        UPDATE approval_flows SET status='已取消', operator_name=?, comment=?, acted_at=?
        WHERE ref_type='invoice' AND ref_id=? AND status='待审批'
    """, (user, f'取消：{reason}', now, iid))
    # 发票置已驳回（统一终止态）
    c.execute("""
        UPDATE invoice_requests
        SET status='已驳回',
            voided_by=?,
            voided_at=?,
            void_reason=?
        WHERE id=?
    """, (user, now, reason, iid))
    log_audit(conn, '取消发票申请', 'invoice_request', iid, f'{user} 取消：{reason}', user)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票申请已取消'})


@app.route('/api/invoice-requests/<int:iid>/issue', methods=['POST'])
@require_role('财务')
def issue_invoice_request(iid):
    data = request.json or {}
    invoice_no = (data.get('invoice_no') or '').strip()
    if not invoice_no:
        return jsonify({'success': False, 'message': '发票号不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] != '待开票':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能开票'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE invoice_requests
        SET status='已开票',
            invoice_no=?,
            invoiced_at=?,
            invoice_file_path=?,
            processed_by=?
        WHERE id=?
    """, (invoice_no, now, (data.get('invoice_file_path') or '').strip(), request.current_user['display_name'], iid))
    log_audit(conn, '发票开具', 'invoice_request', iid,
              f'发票号:{invoice_no}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票已开具'})


@app.route('/api/invoice-requests/<int:iid>/void', methods=['POST'])
@require_role('财务')
def void_invoice_request(iid):
    data = request.json or {}
    reason = (data.get('reason') or '').strip()
    if not reason:
        return jsonify({'success': False, 'message': '作废原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] != '已开票':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能作废（待审批/待开票请用取消）'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE invoice_requests
        SET status='已作废',
            voided_by=?,
            voided_at=?,
            void_reason=?
        WHERE id=?
    """, (request.current_user['display_name'], now, reason, iid))
    log_audit(conn, '发票作废', 'invoice_request', iid, reason, request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票已作废'})


@app.route('/api/invoice-requests/<int:iid>/red', methods=['POST'])
@require_role('财务')
def red_invoice_request(iid):
    data = request.json or {}
    red_invoice_no = (data.get('red_invoice_no') or '').strip()
    reason = (data.get('reason') or data.get('red_reason') or '').strip()
    if not red_invoice_no or not reason:
        return jsonify({'success': False, 'message': '红字发票号和红冲原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM invoice_requests WHERE id=?", (iid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '发票申请不存在'}), 404
    if row['status'] != '已开票':
        conn.close()
        return jsonify({'success': False, 'message': '仅已开票记录可红冲'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE invoice_requests
        SET status='已红冲',
            red_invoice_no=?,
            red_invoiced_at=?,
            red_reason=?,
            red_certificate_path=?
        WHERE id=?
    """, (red_invoice_no, now, reason, (data.get('red_certificate_path') or '').strip(), iid))
    log_audit(conn, '发票红冲', 'invoice_request', iid,
              f'红字发票号:{red_invoice_no} 原因:{reason}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '发票已红冲'})


# ======================== 销售报单 ========================
@app.route('/api/sales-orders', methods=['GET'])
@login_required
def get_sales_orders():
    conn = get_db()
    expire_sales_order_drafts(conn)
    conn.commit()
    c = conn.cursor()
    c.execute("""
        SELECT so.*, v.plate_number as vehicle_plate_number, v.car_type as vehicle_car_type, v.company, v.status as vehicle_status,
               c.contract_type, c.contract_status, c.delivery_status, c.contract_file,
               c.rent AS contract_rent,
               c.loan_periods AS contract_loan_periods,
               c.deposit AS contract_deposit,
               c.down_payment AS contract_down_payment,
               c.start_date AS contract_start_date,
               c.total_price AS contract_total_price,
               c.company AS contract_company,
               customer_plan.customer_plan_periods,
               customer_plan.customer_plan_total,
               customer_plan.customer_plan_avg,
               customer_plan.customer_plan_first_due_date,
               customer_plan.customer_plan_last_due_date,
               factory_plan.factory_plan_periods,
               factory_plan.factory_plan_total,
               factory_plan.factory_plan_avg,
               factory_plan.factory_plan_first_due_date,
               factory_plan.factory_plan_last_due_date
        FROM sales_orders so
        LEFT JOIN vehicles v ON v.id = so.vehicle_id
        LEFT JOIN contracts c ON c.id = so.contract_id
        LEFT JOIN (
            SELECT contract_id,
                   COUNT(CASE WHEN COALESCE(remark, '') NOT IN ('押金', '首付款') AND COALESCE(due_date, '')!='' THEN 1 END) AS customer_plan_periods,
                   COALESCE(SUM(CASE WHEN COALESCE(remark, '') NOT IN ('押金', '首付款') AND COALESCE(due_date, '')!='' THEN amount ELSE 0 END), 0) AS customer_plan_total,
                   COALESCE(AVG(CASE WHEN COALESCE(remark, '') NOT IN ('押金', '首付款') AND COALESCE(due_date, '')!='' THEN amount END), 0) AS customer_plan_avg,
                   MIN(CASE WHEN COALESCE(remark, '') NOT IN ('押金', '首付款') AND COALESCE(due_date, '')!='' THEN due_date END) AS customer_plan_first_due_date,
                   MAX(CASE WHEN COALESCE(remark, '') NOT IN ('押金', '首付款') AND COALESCE(due_date, '')!='' THEN due_date END) AS customer_plan_last_due_date
            FROM repayments
            GROUP BY contract_id
        ) customer_plan ON customer_plan.contract_id = c.id
        LEFT JOIN (
            SELECT contract_id,
                   COUNT(*) AS factory_plan_periods,
                   COALESCE(SUM(amount), 0) AS factory_plan_total,
                   COALESCE(AVG(amount), 0) AS factory_plan_avg,
                   MIN(due_date) AS factory_plan_first_due_date,
                   MAX(due_date) AS factory_plan_last_due_date
            FROM factory_repayments
            GROUP BY contract_id
        ) factory_plan ON factory_plan.contract_id = c.id
        ORDER BY so.id DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
    rows = redact_for_role(conn, request.current_user['role'], 'sales_orders', rows)
    conn.close()
    return jsonify(rows)


@app.route('/api/sales-orders', methods=['POST'])
@require_role('销售')
def create_sales_order():
    data = request.json or {}
    user = request.current_user
    vin = (data.get('vin') or '').strip().upper()
    is_draft = bool(data.get('save_as_draft') or data.get('draft') or data.get('action') == 'draft')

    conn = get_db()
    c = conn.cursor()
    vehicle = None
    vehicles_multi = []

    # 多车支持（SKU 改造）：vehicle_ids 为 VIN 数组（批量租赁/以租代售）；单车主车仍走 vin
    raw_vehicle_ids = data.get('vehicle_ids')
    vin_list = []
    if raw_vehicle_ids:
        if isinstance(raw_vehicle_ids, list):
            vin_list = [str(v).strip().upper() for v in raw_vehicle_ids if str(v).strip()]
        elif isinstance(raw_vehicle_ids, str):
            vin_list = [v.strip().upper() for v in raw_vehicle_ids.replace('，', ',').split(',') if v.strip()]
    elif vin:
        vin_list = [vin]

    if vin_list:
        # 校验所有 VIN
        for i, v_vin in enumerate(vin_list):
            if len(v_vin) != 17:
                conn.close()
                return jsonify({'success': False, 'message': f'第{i+1}个车架号需为17位(VIN)'}), 400
            c.execute("""
                SELECT id, plate_number, car_type, condition, status, box_type, vehicle_box_type,
                       tailgate, brand, fuel_form, cab_type, cab_style, battery_brand,
                       engine_spec, battery_capacity, horsepower, gear_position, gearbox_spec,
                       vehicle_color, purchase_price, validation_status, validation_message, vin
                FROM vehicles
                WHERE vin=?
            """, (v_vin,))
            v_row = c.fetchone()
            if not v_row:
                conn.close()
                return jsonify({'success': False, 'message': f'VIN {v_vin} 未找到对应库存车辆'}), 404
            if not is_draft and v_row['status'] not in ('在库', '报单锁定中'):
                conn.close()
                return jsonify({'success': False, 'message': f'车辆 {v_vin} 仅在库或报单锁定中可发起报单（当前 {v_row["status"]}）'}), 400
            if not is_draft and (v_row['box_type'] or v_row['vehicle_box_type'] or '').strip() == '底盘':
                conn.close()
                return jsonify({
                    'success': False,
                    'message': f'车辆 {v_vin} 为底盘车，请先配置实际厢体后再发起租赁或以租代售报单'
                }), 400
            # 字典校验拦截：无效车辆不能发起报单
            if not is_draft and v_row['validation_status'] == 'invalid':
                v_msg = v_row['validation_message'] or '车型未在指导价字典中维护'
                conn.close()
                return jsonify({'success': False, 'message': f'车辆 {v_vin} 字典校验不通过：{v_msg}，请先修正车辆信息'}), 400
            vehicles_multi.append(dict(v_row))
        vehicle = vehicles_multi[0]
        vin = vehicle['vin']

    # 多车唯一性校验（未结清合同/他人锁单）
    if not is_draft and vehicles_multi:
        vehicle_ids = [v['id'] for v in vehicles_multi]
        placeholders = ', '.join('?' for _ in vehicle_ids)
        c.execute(f"""
            SELECT vehicle_id FROM contracts
            WHERE vehicle_id IN ({placeholders}) AND contract_status NOT IN ('已结清', '已终止')
            ORDER BY id DESC LIMIT 1
        """, vehicle_ids)
        conflict = c.fetchone()
        if conflict:
            conn.close()
            return jsonify({'success': False, 'message': f'存在车辆（id={conflict["vehicle_id"]}）有未结清合同，不能再次报单'}), 400

        for v_row in vehicles_multi:
            c.execute("""
                SELECT id, created_by FROM sales_orders
                WHERE vehicle_id=? AND order_status IN ('待价格特批', '待财务确认', '已激活')
                ORDER BY id DESC LIMIT 1
            """, (v_row['id'],))
            existing_order = c.fetchone()
            if existing_order and existing_order['created_by'] != user['display_name'] and v_row['status'] == '报单锁定中':
                conn.close()
                return jsonify({'success': False, 'message': f'车辆 {v_row["vin"]} 已被 {existing_order["created_by"]} 报单锁定，其他销售不能进行二次报单'}), 400

    if not is_draft:
        customer_name = (data.get('customer_name') or '').strip()
        customer_phone = (data.get('customer_phone') or '').strip()
        c.execute("""
            SELECT id, reason FROM customer_blacklist
            WHERE status='生效'
              AND (
                    (customer_phone!='' AND customer_phone=?)
                    OR (customer_name!='' AND customer_name=?)
              )
            ORDER BY id DESC LIMIT 1
        """, (customer_phone, customer_name))
        blacklist = c.fetchone()
        if blacklist:
            conn.close()
            return jsonify({'success': False, 'message': '客户命中黑名单，请联系老板审核解禁'}), 400

    sale_total_price = 0  # sale_total_price 已废弃（整车销售下线，以租代售仅走金融方案）
    sales_mode = normalize_sales_mode(data.get('sales_mode', '经营租赁'))
    lease_quote = parse_money(data.get('vehicle_rent_amount', data.get('rent')))
    upfront_amount = parse_money(data.get('deposit_amount'))
    order_vehicle = vehicle_order_snapshot(vehicle)
    vehicle_box_type = (
        order_vehicle.get('vehicle_box_type') or data.get('vehicle_box_type') or ''
    ).strip()
    tail_plate = order_vehicle.get('tail_plate') or normalize_tailgate(data.get('tail_plate'))
    finance_plan_id = data.get('finance_plan_id')

    # Use dummy vehicle dict if vehicle is None so guidance check can still use car_type from data
    effective_vehicle = vehicle or {'car_type': data.get('car_type', '').strip(), 'guidance_price': 0}
    guidance_check = calculate_guidance_check(conn, effective_vehicle, sales_mode, sale_total_price, lease_quote, upfront_amount,
                                              vehicle_box_type=vehicle_box_type, tail_plate=tail_plate, finance_plan_id=finance_plan_id)

    quote_price = guidance_check['quote_price']
    guidance_price = guidance_check['guidance_price']
    guidance_label = '指导价'
    # ===== 20260804 首付前移：报单时上传首付截图 + 判断首付足额 =====
    customer_screenshot_path = (data.get('customer_screenshot_path') or data.get('screenshot_path') or '').strip()
    first_payment_received = parse_money(data.get('first_payment_received_amount'))
    if guidance_check['mode'] == '租赁':
        expected_first_payment = round(upfront_amount + lease_quote, 2)  # 押金 + 首月月供
    elif guidance_check['mode'] == '以租代售':
        plan = guidance_check['guidance'].get('plan')
        expected_first_payment = round(parse_money(plan['down_payment'] if plan else 0), 2)
    else:
        expected_first_payment = round(upfront_amount, 2)
    shortage = round(max(0, expected_first_payment - first_payment_received), 2)
    first_payment_shortage_reason = (data.get('first_payment_shortage_reason') or '').strip()
    first_payment_promised_date = normalize_date(data.get('first_payment_promised_date'))

    # 合并审批异常项：租赁价格异常 + 首付不足
    abnormalities = []
    if guidance_check['mode'] == '租赁':
        abnormalities.extend(guidance_check['below'])
    abnormalities.extend(guidance_check['missing'])
    if shortage > 0:
        abnormalities.append(f'首付不足 ¥{shortage}（实收 {first_payment_received} / 应付 {expected_first_payment}）')

    needs_order_approval = not is_draft and bool(abnormalities)
    order_status = '草稿' if is_draft else ('待老板审批' if needs_order_approval else '待财务确认')
    price_check_status = '未提交' if is_draft else ('待老板审批' if needs_order_approval else '无需审批')
    first_payment_check_status = '未校验' if is_draft else ('不足' if shortage > 0 else '足额')
    order_exception_reason = '；'.join(abnormalities) if abnormalities else ''

    # 快照（版本控制落点）：报单提交时锁定方案/指导价
    snapshot_finance_plan = None
    snapshot_lease_deposit_guidance = 0
    snapshot_box_monthly_guidance = 0
    if not is_draft:
        if guidance_check['mode'] == '以租代售' and guidance_check['guidance'].get('plan'):
            p = guidance_check['guidance']['plan']
            snapshot_finance_plan = json.dumps({
                'id': p['id'], 'plan_name': p['plan_name'],
                'down_payment': parse_money(p['down_payment']),
                'period_price': parse_money(p['period_price']),
                'periods': p['periods'],
            }, ensure_ascii=False)
        elif guidance_check['mode'] == '租赁':
            snapshot_lease_deposit_guidance = guidance_check['guidance'].get('deposit_guidance') or 0
            snapshot_box_monthly_guidance = guidance_check['guidance'].get('monthly_guidance') or 0

    if not is_draft:
        if sales_mode in ('租赁', '以租代售'):
            if not customer_screenshot_path:
                conn.close()
                return jsonify({'success': False, 'message': '请先上传客户首次付款截图'}), 400
        if guidance_check['mode'] == '以租代售' and not guidance_check['guidance'].get('plan'):
            conn.close()
            return jsonify({'success': False, 'message': '请选择与该新车厢型一致的以租代售金融方案'}), 400
        if shortage > 0:
            if not first_payment_shortage_reason:
                conn.close()
                return jsonify({'success': False, 'message': '首付不足，请填写不足原因'}), 400
            if not first_payment_promised_date:
                conn.close()
                return jsonify({'success': False, 'message': '首付不足，请填写承诺归还时间'}), 400

    now = datetime.now()
    saved_at = now.strftime('%Y-%m-%d %H:%M:%S') if is_draft else None
    expires_at = (now + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S') if is_draft else None

    c.execute("""
        INSERT INTO sales_orders
            (payment_date, customer_name, customer_phone, customer_id_card, sales_mode, vehicle_id, vin,
             is_new, vehicle_brand, lease_start_date,
             vehicle_category, vehicle_cab, vehicle_engine_battery, vehicle_power_battery,
             vehicle_gearbox, vehicle_box_type,
             car_type, vehicle_color, plate_number, tail_plate, lease_term, cargo_length, sale_total_price,
             payment_category, car_purchase_amount, vehicle_rent_amount, receiving_company,
             wechat_interest, wechat_registration_fee, wechat_purchase_tax,
             full_package, wechat_private_fee, gifted_items, deposit_amount, order_status,
             snapshot_guidance_price, snapshot_lease_installment_price, snapshot_sale_total_price,
             price_check_status, price_exception_reason, customer_plan_match_status, factory_plan_match_status,
             saved_at, expires_at, sales_advisor, remark, created_by,
             customer_screenshot_path, first_payment_received_amount, first_payment_shortage_amount,
             first_payment_shortage_reason, first_payment_promised_date, first_payment_check_status,
             order_exception_reason, finance_plan_id, snapshot_finance_plan,
             snapshot_lease_deposit_guidance, snapshot_box_monthly_guidance, vehicle_ids)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data.get('payment_date') or datetime.now().strftime('%Y-%m-%d'),
        data.get('customer_name', '').strip() or ('草稿客户' if is_draft else ''),
        data.get('customer_phone', '').strip(),
        (data.get('customer_id_card') or '').strip(),
        sales_mode,
        vehicle['id'] if vehicle else None,
        vin,
        order_vehicle.get('is_new') or (data.get('is_new') or '新车').strip(),
        order_vehicle.get('vehicle_brand') or (data.get('vehicle_brand') or '').strip(),
        (data.get('lease_start_date') or '').strip(),
        (data.get('vehicle_category') or '').strip(),
        (data.get('vehicle_cab') or '').strip(),
        (data.get('vehicle_engine_battery') or '').strip(),
        (data.get('vehicle_power_battery') or '').strip(),
        (data.get('vehicle_gearbox') or '').strip(),
        vehicle_box_type,
        order_vehicle.get('car_type') or data.get('car_type', '').strip(),
        order_vehicle.get('vehicle_color') or (data.get('vehicle_color') or '').strip(),
        order_vehicle.get('plate_number') or (data.get('plate_number') or '').strip(),
        tail_plate,
        data.get('lease_term', '').strip(),
        data.get('cargo_length', '').strip(),
        sale_total_price,
        data.get('payment_category', '').strip(),
        order_vehicle.get('car_purchase_amount', parse_money(data.get('car_purchase_amount'))),
        parse_money(data.get('vehicle_rent_amount')),
        data.get('receiving_company', '').strip(),
        parse_money(data.get('wechat_interest')),
        parse_money(data.get('wechat_registration_fee')),
        parse_money(data.get('wechat_purchase_tax')),
        1 if data.get('full_package') else 0,
        parse_money(data.get('wechat_private_fee')),
        data.get('gifted_items', '').strip(),
        parse_money(data.get('deposit_amount')),
        order_status,
        0 if is_draft else guidance_price,
        0 if is_draft else guidance_check['guidance'].get('lease_installment_price', 0) if isinstance(guidance_check['guidance'], dict) else 0,
        0,  # snapshot_sale_total_price 已废弃
        price_check_status,
        order_exception_reason,
        '未生成',
        '未上传',
        saved_at,
        expires_at,
        data.get('sales_advisor', '').strip() or user['display_name'],
        data.get('remark', '').strip(),
        user['display_name'],
        customer_screenshot_path,
        first_payment_received,
        shortage,
        first_payment_shortage_reason,
        first_payment_promised_date,
        first_payment_check_status,
        order_exception_reason,
        finance_plan_id,
        snapshot_finance_plan,
        snapshot_lease_deposit_guidance,
        snapshot_box_monthly_guidance,
        json.dumps([v['id'] for v in vehicles_multi], ensure_ascii=False) if vehicles_multi else None,
    ))
    order_id = c.lastrowid
    if needs_order_approval:
        create_approval_flow(conn, 'order_exception', order_id)
    elif not is_draft:
        ensure_default_sales_order_planning_if_needed(conn, order_id)
        create_approval_flow(conn, 'sale_payment', order_id)
    # 批量锁车（多车报单锁定所有车辆）
    if not is_draft and vehicles_multi:
        vids = tuple(v['id'] for v in vehicles_multi)
        placeholders = ', '.join('?' for _ in vids)
        c.execute(f"UPDATE vehicles SET status='报单锁定中' WHERE id IN ({placeholders})", vids)
    log_audit(conn, '创建销售报单', 'sales_order', order_id,
              f"{user['display_name']} 报单 VIN:{vin} 模式:{sales_mode} 首付实收:{first_payment_received} 状态:{order_status}"
              + (f' 异常:{order_exception_reason}' if order_exception_reason else ''),
              user['display_name'])
    conn.commit()
    conn.close()
    if is_draft:
        return jsonify({'success': True, 'id': order_id, 'message': '草稿已保存，7天内有效'})
    message = '销售报单已提交，车辆已锁定'
    if needs_order_approval:
        message = '销售报单已提交，等待老板报单审批'
    return jsonify({'success': True, 'id': order_id, 'message': message})


@app.route('/api/sales-orders/<int:order_id>', methods=['PUT'])
@require_role('销售')
def update_sales_order_draft(order_id):
    data = request.json or {}
    submit_now = data.get('action') == 'submit' or bool(data.get('submit'))
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM sales_orders WHERE id=?", (order_id,))
    order = c.fetchone()
    if not order:
        conn.close()
        return jsonify({'success': False, 'message': '草稿不存在'}), 404
    if order['order_status'] != '草稿':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{order["order_status"]}，只能编辑草稿'}), 400

    vin = (data.get('vin') or order['vin'] or '').strip().upper()
    vehicle = None
    vehicles_multi = []

    # 多车支持（SKU 改造）：草稿提交同样支持 vehicle_ids
    raw_vehicle_ids = data.get('vehicle_ids')
    vin_list = []
    if raw_vehicle_ids:
        if isinstance(raw_vehicle_ids, list):
            vin_list = [str(v).strip().upper() for v in raw_vehicle_ids if str(v).strip()]
        elif isinstance(raw_vehicle_ids, str):
            vin_list = [v.strip().upper() for v in raw_vehicle_ids.replace('，', ',').split(',') if v.strip()]
    elif vin:
        vin_list = [vin]
    if vin_list:
        for v_vin in vin_list:
            if len(v_vin) != 17:
                conn.close()
                return jsonify({'success': False, 'message': '如填写车架号，请填写17位(VIN)'}), 400
            c.execute("""
                SELECT id, plate_number, car_type, condition, status, box_type, vehicle_box_type,
                       tailgate, brand, fuel_form, cab_type, cab_style, battery_brand,
                       engine_spec, battery_capacity, horsepower, gear_position, gearbox_spec,
                       vehicle_color, purchase_price, validation_status, validation_message, vin
                FROM vehicles WHERE vin=?
            """, (v_vin,))
            v_row = c.fetchone()
            if not v_row:
                conn.close()
                return jsonify({'success': False, 'message': '未找到对应库存车辆'}), 404
            if submit_now and v_row['validation_status'] == 'invalid':
                v_msg = v_row['validation_message'] or '车型未在指导价字典中维护'
                conn.close()
                return jsonify({'success': False, 'message': f'车辆 {v_vin} 字典校验不通过：{v_msg}，请先修正车辆信息'}), 400
            vehicles_multi.append(dict(v_row))
        vehicle = vehicles_multi[0]
        vin = vehicle['vin']

    sales_mode = normalize_sales_mode(data.get('sales_mode', order['sales_mode']))
    sale_total_price = 0  # sale_total_price 已废弃（整车销售下线，以租代售仅走金融方案）
    lease_quote = parse_money(data.get('vehicle_rent_amount', data.get('rent')), order['vehicle_rent_amount'])
    upfront_amount = parse_money(data.get('deposit_amount'), order['deposit_amount'])
    order_vehicle = vehicle_order_snapshot(vehicle)
    vehicle_box_type = (
        order_vehicle.get('vehicle_box_type')
        or data.get('vehicle_box_type')
        or order['vehicle_box_type']
        or ''
    ).strip()
    tail_plate = order_vehicle.get('tail_plate') or normalize_tailgate(
        data.get('tail_plate') or order['tail_plate']
    )
    finance_plan_id = data.get('finance_plan_id') if data.get('finance_plan_id') is not None else order['finance_plan_id']
    effective_vehicle = vehicle or {'car_type': data.get('car_type', order['car_type']).strip(), 'guidance_price': 0}
    guidance_check = calculate_guidance_check(conn, effective_vehicle, sales_mode, sale_total_price, lease_quote, upfront_amount,
                                              vehicle_box_type=vehicle_box_type, tail_plate=tail_plate, finance_plan_id=finance_plan_id)
    quote_price = guidance_check['quote_price']
    guidance_price = guidance_check['guidance_price']

    order_status = '草稿'
    price_check_status = '未提交'
    price_exception_reason = ''
    saved_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    expires_at = (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
    snapshot_guidance_price = 0
    snapshot_lease_price = 0
    snapshot_sale_price = 0

    customer_screenshot_path = (data.get('customer_screenshot_path') or order['customer_screenshot_path'] or '').strip()
    first_payment_received = parse_money(data.get('first_payment_received_amount'), order['first_payment_received_amount'])
    first_payment_shortage_reason = (data.get('first_payment_shortage_reason') or order['first_payment_shortage_reason'] or '').strip()
    first_payment_promised_date = normalize_date(data.get('first_payment_promised_date')) or order['first_payment_promised_date']
    first_payment_check_status = '未校验'
    order_exception_reason = ''
    snapshot_finance_plan = None
    snapshot_lease_deposit_guidance = 0
    snapshot_box_monthly_guidance = 0

    if submit_now:
        if vehicle and vehicle['status'] not in ('在库', '报单锁定中'):
            conn.close()
            return jsonify({'success': False, 'message': '该车辆已被其他合同占用，请调整 VIN'}), 400
        if vehicle and (vehicle.get('box_type') or vehicle.get('vehicle_box_type') or '').strip() == '底盘':
            conn.close()
            return jsonify({'success': False, 'message': '底盘车请先配置实际厢体后再发起租赁或以租代售报单'}), 400
        customer_name = (data.get('customer_name') or order['customer_name'] or '').strip()
        customer_phone = (data.get('customer_phone') or order['customer_phone'] or '').strip()
        c.execute("""
            SELECT id FROM customer_blacklist
            WHERE status='生效'
              AND ((customer_phone!='' AND customer_phone=?) OR (customer_name!='' AND customer_name=?))
            LIMIT 1
        """, (customer_phone, customer_name))
        if c.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '客户命中黑名单，请联系老板审核解禁'}), 400

        if vehicle:
            c.execute("""
                SELECT id, created_by FROM sales_orders
                WHERE vehicle_id=? AND id!=? AND order_status IN ('待价格特批', '待财务确认', '已激活')
                ORDER BY id DESC LIMIT 1
            """, (vehicle['id'], order_id))
            existing_order = c.fetchone()
            if existing_order and existing_order['created_by'] != user['display_name']:
                conn.close()
                return jsonify({'success': False, 'message': f'该车辆已被 {existing_order["created_by"]} 报单锁定，其他销售不能进行二次报单'}), 400

        # 幂等检查：已存在待审批流则不允许重复提交
        c.execute("""
            SELECT id FROM approval_flows
            WHERE ref_type IN ('order_exception', 'sale_payment') AND ref_id=? AND status='待审批'
            LIMIT 1
        """, (order_id,))
        if c.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '该报单已在审批中，请勿重复提交'}), 400

        # ===== 首付判断（同 create）=====
        if guidance_check['mode'] == '租赁':
            expected_first_payment = round(upfront_amount + lease_quote, 2)
        elif guidance_check['mode'] == '以租代售':
            plan = guidance_check['guidance'].get('plan')
            expected_first_payment = round(parse_money(plan['down_payment'] if plan else 0), 2)
        else:
            expected_first_payment = round(upfront_amount, 2)
        shortage = round(max(0, expected_first_payment - first_payment_received), 2)

        abnormalities = []
        if guidance_check['mode'] == '租赁':
            abnormalities.extend(guidance_check['below'])
        abnormalities.extend(guidance_check['missing'])
        if shortage > 0:
            abnormalities.append(f'首付不足 ¥{shortage}（实收 {first_payment_received} / 应付 {expected_first_payment}）')

        needs_order_approval = bool(abnormalities)
        order_status = '待老板审批' if needs_order_approval else '待财务确认'
        price_check_status = '待老板审批' if needs_order_approval else '无需审批'
        first_payment_check_status = '不足' if shortage > 0 else '足额'
        order_exception_reason = '；'.join(abnormalities) if abnormalities else ''
        saved_at = None
        expires_at = None
        snapshot_guidance_price = guidance_price
        snapshot_lease_price = guidance_check['guidance'].get('lease_installment_price', 0) if isinstance(guidance_check['guidance'], dict) else 0
        snapshot_sale_price = 0  # sale_total_price 已废弃（以租代售仅走金融方案）

        if guidance_check['mode'] == '以租代售' and guidance_check['guidance'].get('plan'):
            p = guidance_check['guidance']['plan']
            snapshot_finance_plan = json.dumps({
                'id': p['id'], 'plan_name': p['plan_name'],
                'down_payment': parse_money(p['down_payment']),
                'period_price': parse_money(p['period_price']),
                'periods': p['periods'],
            }, ensure_ascii=False)
        elif guidance_check['mode'] == '租赁':
            snapshot_lease_deposit_guidance = guidance_check['guidance'].get('deposit_guidance') or 0
            snapshot_box_monthly_guidance = guidance_check['guidance'].get('monthly_guidance') or 0

        if sales_mode in ('租赁', '以租代售'):
            if not customer_screenshot_path:
                conn.close()
                return jsonify({'success': False, 'message': '请先上传客户首次付款截图'}), 400
        if guidance_check['mode'] == '以租代售' and not guidance_check['guidance'].get('plan'):
            conn.close()
            return jsonify({'success': False, 'message': '请选择与该新车厢型一致的以租代售金融方案'}), 400
        if shortage > 0:
            if not first_payment_shortage_reason:
                conn.close()
                return jsonify({'success': False, 'message': '首付不足，请填写不足原因'}), 400
            if not first_payment_promised_date:
                conn.close()
                return jsonify({'success': False, 'message': '首付不足，请填写承诺归还时间'}), 400

    c.execute("""
        UPDATE sales_orders
        SET payment_date=?, customer_name=?, customer_phone=?, customer_id_card=?, sales_mode=?, vehicle_id=?, vin=?,
            is_new=?, vehicle_brand=?, lease_start_date=?,
            vehicle_category=?, vehicle_cab=?, vehicle_engine_battery=?, vehicle_power_battery=?,
            vehicle_gearbox=?, vehicle_box_type=?,
            car_type=?, vehicle_color=?, plate_number=?, tail_plate=?, lease_term=?, cargo_length=?, sale_total_price=?,
            payment_category=?, car_purchase_amount=?, vehicle_rent_amount=?, receiving_company=?,
            wechat_interest=?, wechat_registration_fee=?, wechat_purchase_tax=?,
            full_package=?, wechat_private_fee=?, gifted_items=?, deposit_amount=?, order_status=?,
            snapshot_guidance_price=?, snapshot_lease_installment_price=?, snapshot_sale_total_price=?,
            price_check_status=?, price_exception_reason=?, customer_plan_match_status=?, factory_plan_match_status=?,
            saved_at=?, expires_at=?, sales_advisor=?, remark=?,
            customer_screenshot_path=?, first_payment_received_amount=?, first_payment_shortage_amount=?,
            first_payment_shortage_reason=?, first_payment_promised_date=?, first_payment_check_status=?,
            order_exception_reason=?, finance_plan_id=?, snapshot_finance_plan=?,
            snapshot_lease_deposit_guidance=?, snapshot_box_monthly_guidance=?, vehicle_ids=?
        WHERE id=?
    """, (
        data.get('payment_date') or order['payment_date'] or datetime.now().strftime('%Y-%m-%d'),
        (data.get('customer_name') or order['customer_name'] or '').strip(),
        (data.get('customer_phone') or order['customer_phone'] or '').strip(),
        (data.get('customer_id_card') or order['customer_id_card'] or '').strip(),
        sales_mode,
        vehicle['id'] if vehicle else None,
        vin,
        order_vehicle.get('is_new') or (data.get('is_new') or order['is_new'] or '新车').strip(),
        order_vehicle.get('vehicle_brand') or (data.get('vehicle_brand') or order['vehicle_brand'] or '').strip(),
        (data.get('lease_start_date') or order['lease_start_date'] or '').strip(),
        (data.get('vehicle_category') or order['vehicle_category'] or '').strip(),
        (data.get('vehicle_cab') or order['vehicle_cab'] or '').strip(),
        (data.get('vehicle_engine_battery') or order['vehicle_engine_battery'] or '').strip(),
        (data.get('vehicle_power_battery') or order['vehicle_power_battery'] or '').strip(),
        (data.get('vehicle_gearbox') or order['vehicle_gearbox'] or '').strip(),
        vehicle_box_type,
        order_vehicle.get('car_type') or (data.get('car_type') or order['car_type'] or '').strip(),
        order_vehicle.get('vehicle_color') or (data.get('vehicle_color') or order['vehicle_color'] or '').strip(),
        order_vehicle.get('plate_number') or (data.get('plate_number') or order['plate_number'] or '').strip(),
        tail_plate,
        (data.get('lease_term') or order['lease_term'] or '').strip(),
        (data.get('cargo_length') or order['cargo_length'] or '').strip(),
        sale_total_price,
        (data.get('payment_category') or order['payment_category'] or '').strip(),
        order_vehicle.get('car_purchase_amount', parse_money(data.get('car_purchase_amount'), order['car_purchase_amount'])),
        lease_quote,
        (data.get('receiving_company') or order['receiving_company'] or '').strip(),
        parse_money(data.get('wechat_interest'), order['wechat_interest']),
        parse_money(data.get('wechat_registration_fee'), order['wechat_registration_fee']),
        parse_money(data.get('wechat_purchase_tax'), order['wechat_purchase_tax']),
        1 if data.get('full_package', order['full_package']) else 0,
        parse_money(data.get('wechat_private_fee'), order['wechat_private_fee']),
        (data.get('gifted_items') or order['gifted_items'] or '').strip(),
        parse_money(data.get('deposit_amount'), order['deposit_amount']),
        order_status,
        snapshot_guidance_price,
        snapshot_lease_price,
        snapshot_sale_price,
        price_check_status,
        price_exception_reason,
        '未生成',
        '未上传',
        saved_at,
        expires_at,
        (data.get('sales_advisor') or order['sales_advisor'] or user['display_name']).strip(),
        (data.get('remark') or order['remark'] or '').strip(),
        customer_screenshot_path,
        first_payment_received,
        round(max(0, shortage), 2) if submit_now else parse_money(order['first_payment_shortage_amount']),
        first_payment_shortage_reason,
        first_payment_promised_date,
        first_payment_check_status,
        order_exception_reason,
        finance_plan_id,
        snapshot_finance_plan,
        snapshot_lease_deposit_guidance,
        snapshot_box_monthly_guidance,
        json.dumps([v['id'] for v in vehicles_multi], ensure_ascii=False) if vehicles_multi else order['vehicle_ids'] if 'vehicle_ids' in order.keys() and order['vehicle_ids'] else None,
        order_id,
    ))
    if submit_now:
        if order_status == '待老板审批':
            create_approval_flow(conn, 'order_exception', order_id)
        else:
            ensure_default_sales_order_planning_if_needed(conn, order_id)
            create_approval_flow(conn, 'sale_payment', order_id)
        # 批量锁车（多车草稿提交锁定所有车辆）
        if vehicles_multi:
            vids = tuple(v['id'] for v in vehicles_multi)
            placeholders = ', '.join('?' for _ in vids)
            c.execute(f"UPDATE vehicles SET status='报单锁定中' WHERE id IN ({placeholders})", vids)
    log_audit(conn, '提交草稿报单' if submit_now else '更新报单草稿', 'sales_order', order_id,
              f'VIN:{vin} 状态:{order_status}', user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': order_id, 'message': '草稿已提交' if submit_now else '草稿已保存'})


@app.route('/api/sales-orders/<int:order_id>', methods=['DELETE'])
@require_role('销售')
def delete_sales_order_draft(order_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT order_status FROM sales_orders WHERE id=?", (order_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '草稿不存在'}), 404
    if row['order_status'] != '草稿':
        conn.close()
        return jsonify({'success': False, 'message': '只能删除草稿'}), 400
    c.execute("DELETE FROM sales_orders WHERE id=?", (order_id,))
    log_audit(conn, '删除报单草稿', 'sales_order', order_id, '销售删除未提交草稿', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '草稿已删除'})


@app.route('/api/sales-orders/<int:order_id>/copy-draft', methods=['POST'])
@require_role('销售')
def copy_sales_order_to_draft(order_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM sales_orders WHERE id=?", (order_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '报单不存在'}), 404
    if row['order_status'] not in ('已作废', '草稿'):
        conn.close()
        return jsonify({'success': False, 'message': '仅草稿或已作废报单可复制为草稿'}), 400
    now = datetime.now()
    c.execute("""
        INSERT INTO sales_orders
            (payment_date, customer_name, customer_phone, customer_id_card, sales_mode, vehicle_id, vin,
             is_new, vehicle_brand, lease_start_date,
             vehicle_category, vehicle_cab, vehicle_engine_battery, vehicle_power_battery,
             vehicle_gearbox, vehicle_box_type,
             car_type, vehicle_color, plate_number, tail_plate, lease_term, cargo_length, sale_total_price,
             payment_category, car_purchase_amount, vehicle_rent_amount, receiving_company,
             wechat_interest, wechat_registration_fee, wechat_purchase_tax,
             full_package, wechat_private_fee, gifted_items, deposit_amount, order_status,
             price_check_status, customer_plan_match_status, factory_plan_match_status,
             saved_at, expires_at, sales_advisor, remark, created_by)
        SELECT payment_date, customer_name, customer_phone, customer_id_card, sales_mode, vehicle_id, vin,
               is_new, vehicle_brand, lease_start_date,
               vehicle_category, vehicle_cab, vehicle_engine_battery, vehicle_power_battery,
               vehicle_gearbox, vehicle_box_type,
               car_type, vehicle_color, plate_number, tail_plate, lease_term, cargo_length, sale_total_price,
               payment_category, car_purchase_amount, vehicle_rent_amount, receiving_company,
               wechat_interest, wechat_registration_fee, wechat_purchase_tax,
               full_package, wechat_private_fee, gifted_items, deposit_amount, '草稿',
               '未提交', '未生成', '未上传',
               ?, ?, sales_advisor, remark, ?
        FROM sales_orders
        WHERE id=?
    """, (
        now.strftime('%Y-%m-%d %H:%M:%S'),
        (now + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S'),
        request.current_user['display_name'],
        order_id,
    ))
    draft_id = c.lastrowid
    log_audit(conn, '复制报单为草稿', 'sales_order', draft_id,
              f'来源报单:{order_id}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': draft_id, 'message': '已复制为新草稿'})


@app.route('/api/sales-orders/<int:order_id>/activate', methods=['POST'])
@require_role('财务')
def activate_sales_order(order_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM sales_orders WHERE id=?", (order_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '报单不存在'}), 404
    if row['order_status'] == '待价格特批':
        conn.close()
        return jsonify({'success': False, 'message': '成交价低于指导价，请先由老板完成价格审批'}), 400
    if row['order_status'] == '待老板审批':
        conn.close()
        return jsonify({'success': False, 'message': '报单存在价格异常或首付不足，请先由老板完成报单审批'}), 400
    if row['order_status'] == '已作废':
        conn.close()
        return jsonify({'success': False, 'message': '报单已作废，不能确认'}), 400
    if row['order_status'] == '已激活':
        conn.close()
        return jsonify({'success': False, 'message': '该报单已激活'}), 400
    if row['order_status'] != '待财务确认':
        conn.close()
        return jsonify({'success': False, 'message': f'当前报单状态为{row["order_status"]}，不能财务确认'}), 400

    blocker = sales_order_plan_activation_blocker(conn, order_id)
    if blocker:
        conn.close()
        return jsonify({'success': False, 'message': blocker}), 400

    # 20260804：财务确认 = 首付款对账 + 确认报单。须填银行流水号（≥4位）或上传公司回单。
    data = request.json or {}
    bank_serial = (data.get('bank_serial') or '').strip()
    bank_receipt_path = (data.get('bank_receipt_path') or data.get('receipt_path') or '').strip()
    if len(bank_serial) < 4 and not bank_receipt_path:
        conn.close()
        return jsonify({'success': False, 'message': '请填写银行流水号（至少4位）或上传公司收款回单'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE sales_orders
        SET order_status='已激活',
            finance_confirmed_by=?,
            finance_confirmed_at=?,
            finance_bank_serial=?,
            finance_bank_receipt_path=?
        WHERE id=?
    """, (
        request.current_user['display_name'],
        now,
        bank_serial or None,
        bank_receipt_path or None,
        order_id,
    ))
    c.execute("""
        UPDATE approval_flows
        SET status='已通过', operator_id=?, operator_name=?, comment=COALESCE(NULLIF(comment, ''), '财务确认报单'), acted_at=?
        WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
    """, (request.current_user['id'], request.current_user['display_name'], now, order_id))
    log_audit(conn, '确认销售报单', 'sales_order', order_id,
              f"财务确认报单 首付实收 ¥{parse_money(row['first_payment_received_amount'])} 流水号:{bank_serial or '未填'} 回单:{bank_receipt_path or '未传'}",
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '报单已确认，等待运营上传线下合同'})


@app.route('/api/sales-orders/<int:order_id>/planning-contract', methods=['GET', 'PUT'])
@require_role('运营', '财务')
def sales_order_planning_contract(order_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM sales_orders WHERE id=?", (order_id,))
    order = c.fetchone()
    if not order:
        conn.close()
        return jsonify({'success': False, 'message': '报单不存在'}), 404
    if request.method == 'PUT' and request.current_user['role'] == '财务':
        conn.close()
        return jsonify({'success': False, 'message': '客户还款计划由运营维护，财务确认报单不再上传厂家分期表'}), 403

    if request.method == 'GET' and not order['contract_id']:
        conn.close()
        return jsonify({
            'success': True,
            'contract': None,
            'repayments': [],
            'factory_repayments': [],
            'comparison': {
                'status': '未生成',
                'message': '请先保存客户还款计划',
            },
        })

    if request.method == 'GET':
        contract_id = order['contract_id']
        c.execute("SELECT * FROM contracts WHERE id=?", (contract_id,))
        contract = dict(c.fetchone())
        c.execute("SELECT * FROM repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
        repayments = [dict(row) for row in c.fetchall()]
        c.execute("SELECT * FROM factory_repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
        factory_rows = [dict(row) for row in c.fetchall()]
        comparison = build_customer_plan_summary(conn, contract_id)
        conn.commit()
        conn.close()
        return jsonify({
            'success': True,
            'contract': contract,
            'repayments': repayments,
            'factory_repayments': factory_rows,
            'comparison': comparison,
        })

    data = request.json or {}
    try:
        contract_id = ensure_sales_order_planning_contract(conn, order_id, data, reset_factory=True)
    except ValueError as exc:
        conn.close()
        return jsonify({'success': False, 'message': str(exc)}), 400
    if not contract_id:
        conn.close()
        return jsonify({'success': False, 'message': '当前报单不能生成分期计划'}), 400

    c.execute("SELECT * FROM contracts WHERE id=?", (contract_id,))
    contract = dict(c.fetchone())
    c.execute("SELECT * FROM repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
    repayments = [dict(row) for row in c.fetchall()]
    c.execute("SELECT * FROM factory_repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
    factory_rows = [dict(row) for row in c.fetchall()]
    comparison = build_customer_plan_summary(conn, contract_id)
    log_audit(conn, '生成报单客户还款计划', 'sales_order', order_id,
              f"合同壳:{contract_id} 客户期数:{len(repayments)} 比对状态:{comparison.get('status')}",
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({
        'success': True,
        'contract_id': contract_id,
        'contract': contract,
        'repayments': repayments,
        'factory_repayments': factory_rows,
        'comparison': comparison,
        'message': '客户还款计划已生成，财务可直接确认报单',
    })


@app.route('/api/sales-orders/<int:order_id>/void', methods=['POST'])
@login_required
def void_sales_order(order_id):
    data = request.json or {}
    user = request.current_user
    reason = (data.get('reason') or data.get('comment') or '').strip()
    if not reason:
        return jsonify({'success': False, 'message': '作废原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vehicle_id, vehicle_ids, contract_id, order_status FROM sales_orders WHERE id=?", (order_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '报单不存在'}), 404
    if row['order_status'] not in ('草稿', '待价格特批', '待财务确认'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["order_status"]}，不能作废'}), 400
    if user['role'] == '销售' and row['order_status'] not in ('草稿', '待价格特批', '待财务确认'):
        conn.close()
        return jsonify({'success': False, 'message': '销售只能撤回草稿或待审批报单'}), 403

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE sales_orders
        SET order_status='已作废', voided_at=?, void_reason=?, voided_by=?
        WHERE id=?
    """, (now, reason, user['display_name'], order_id))
    if row['vehicle_id'] and row['order_status'] != '草稿':
        # 释放所有关联车辆（SKU 改造：多车报单批量释放）
        try:
            vehicle_ids = json.loads(row['vehicle_ids']) if row['vehicle_ids'] else []
        except (TypeError, ValueError):
            vehicle_ids = []
        release_ids = list(vehicle_ids) if vehicle_ids else [row['vehicle_id']]
        for vid in release_ids:
            c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (vid,))
    # 终止报单计划合同壳（避免残留合同壳阻塞车辆重新报单）
    if row['contract_id']:
        c.execute("""
            UPDATE contracts SET contract_status='已终止', delivery_status='已终止'
            WHERE id=? AND contract_status='报单计划中'
        """, (row['contract_id'],))
    c.execute("""
        UPDATE approval_flows
        SET status='已取消', comment=COALESCE(NULLIF(comment, ''), ?), acted_at=COALESCE(acted_at, ?)
        WHERE ref_type IN ('price_exception', 'sale_payment') AND ref_id=? AND status='待审批'
    """, (reason, now, order_id))
    log_audit(conn, '作废销售报单', 'sales_order', order_id, reason, user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '报单已作废，车辆已释放'})


# ======================== 合同 CRUD ========================
@app.route('/api/contracts', methods=['GET'])
def get_contracts():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT c.*, v.vin, v.plate_number, v.car_type, v.status as vehicle_status,
               cu.name as customer_name, cu.phone as customer_phone,
               (SELECT ip.id FROM contract_initial_payments ip WHERE ip.contract_id=c.id ORDER BY ip.id DESC LIMIT 1) as initial_payment_id,
               (SELECT ip.status FROM contract_initial_payments ip WHERE ip.contract_id=c.id ORDER BY ip.id DESC LIMIT 1) as initial_payment_status,
               (SELECT ip.amount FROM contract_initial_payments ip WHERE ip.contract_id=c.id ORDER BY ip.id DESC LIMIT 1) as initial_payment_amount
        FROM contracts c
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY c.id ASC
    """)
    contracts = [dict(row) for row in c.fetchall()]
    # SKU 改造：多车合同带出关联车辆 VIN 列表（供前端展示）
    for ct in contracts:
        cv_rows = c.execute("""
            SELECT v.vin, v.plate_number, cv.is_primary
            FROM contract_vehicles cv
            JOIN vehicles v ON v.id = cv.vehicle_id
            WHERE cv.contract_id=?
            ORDER BY cv.is_primary DESC, cv.id ASC
        """, (ct['id'],)).fetchall()
        if cv_rows:
            ct['contract_vehicle_vins'] = [
                {'vin': r['vin'], 'plate_number': r['plate_number'], 'is_primary': r['is_primary']}
                for r in cv_rows
            ]
        else:
            ct['contract_vehicle_vins'] = []
    user = get_current_user()
    if user:
        contracts = redact_for_role(conn, user['role'], 'contracts', contracts)
    conn.close()
    return jsonify(contracts)


@app.route('/api/contracts', methods=['POST'])
@require_role('运营')
def add_contract():
    data = request.json or {}
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    try:
        vehicle_id = data.get('vehicle_id')
        if not vehicle_id:
            return jsonify({'success': False, 'message': '请选择车辆'}), 400
        sales_order_id = data.get('sales_order_id')
        order = None
        # 默认从请求中读取合同类型，若未提供则默认为租赁
        contract_type = data.get('contract_type') or '租赁'
        if sales_order_id:
            c.execute("SELECT * FROM sales_orders WHERE id=?", (sales_order_id,))
            order = c.fetchone()
            if not order:
                return jsonify({'success': False, 'message': '关联报单不存在'}), 404
            # 多车支持（SKU 改造）：报单 vehicle_ids（JSON）中须含当前主车
            order_vehicle_ids = []
            try:
                order_vehicle_ids = json.loads(order['vehicle_ids']) if order['vehicle_ids'] else []
            except (TypeError, ValueError):
                order_vehicle_ids = []
            if order_vehicle_ids:
                if vehicle_id not in order_vehicle_ids:
                    return jsonify({'success': False, 'message': '报单车辆与合同车辆不一致'}), 400
            elif order['vehicle_id'] != vehicle_id:
                return jsonify({'success': False, 'message': '报单车辆与合同车辆不一致'}), 400
            if order['order_status'] != '已激活':
                if order['order_status'] == '待价格特批':
                    return jsonify({'success': False, 'message': '成交价低于指导价，请先由老板完成价格审批'}), 400
                return jsonify({'success': False, 'message': f'当前报单状态为{order["order_status"]}，请先由财务确认报单后再上传线下合同'}), 400

        contract_file = (data.get('contract_file') or '').strip()
        requested_contract_id = data.get('supplement_contract_id') or (order['contract_id'] if order else None)
        existing_order_contract = None
        if sales_order_id and requested_contract_id:
            c.execute("""
                SELECT id, vehicle_id, contract_status, delivery_status, contract_file
                FROM contracts
                WHERE id=? AND sales_order_id=?
            """, (requested_contract_id, sales_order_id))
            existing_order_contract = c.fetchone()
        if sales_order_id and not existing_order_contract:
            c.execute("""
                SELECT id, vehicle_id, contract_status, delivery_status, contract_file
                FROM contracts
                WHERE sales_order_id=?
                  AND contract_status!='报单计划中'
                  AND contract_status NOT IN ('已结清', '已终止')
                ORDER BY id DESC LIMIT 1
            """, (sales_order_id,))
            existing_order_contract = c.fetchone()

        if existing_order_contract and existing_order_contract['contract_status'] != '报单计划中':
            if existing_order_contract['vehicle_id'] != vehicle_id:
                return jsonify({'success': False, 'message': '报单合同车辆与当前车辆不一致'}), 400
            if not contract_file:
                return jsonify({'success': False, 'message': '请上传要补充的合同附件'}), 400
            merged_contract_file = merge_attachment_csv(existing_order_contract['contract_file'], contract_file)
            c.execute("UPDATE contracts SET contract_file=? WHERE id=?", (merged_contract_file, existing_order_contract['id']))
            c.execute("UPDATE sales_orders SET contract_id=? WHERE id=?", (existing_order_contract['id'], sales_order_id))
            log_audit(conn, '补充线下合同附件', 'contract', existing_order_contract['id'],
                      f'新增附件:{contract_file} 保留原附件:{existing_order_contract["contract_file"] or ""}',
                      user['display_name'])
            conn.commit()
            return jsonify({
                'success': True,
                'id': existing_order_contract['id'],
                'contract_file': merged_contract_file,
                'message': '合同附件已补充，原合同数据已保留'
            })

        planning_contract_id = (
            existing_order_contract['id']
            if existing_order_contract and existing_order_contract['contract_status'] == '报单计划中'
            else None
        )
        c.execute("""
            SELECT id
            FROM contracts
            WHERE sales_order_id=? AND contract_status='报单计划中'
            ORDER BY id DESC LIMIT 1
        """, (sales_order_id,))
        planning = c.fetchone()
        planning_contract_id = planning_contract_id or (planning['id'] if planning else None)

        # ===== 校验：报单内所有车辆都不能重复签约；历史计划合同可复用 =====
        contract_vehicles_all = []
        try:
            contract_vehicles_all = json.loads(order['vehicle_ids']) if order and order['vehicle_ids'] else []
        except (TypeError, ValueError):
            contract_vehicles_all = []
        if not contract_vehicles_all:
            contract_vehicles_all = [vehicle_id]
        for cvid in contract_vehicles_all:
            if planning_contract_id:
                c.execute("""SELECT id, contract_type, contract_status, delivery_status
                             FROM contracts
                             WHERE vehicle_id=? AND contract_status NOT IN ('已结清', '已终止') AND id!=?""",
                          (cvid, planning_contract_id))
            else:
                c.execute("""SELECT id, contract_type, contract_status, delivery_status
                             FROM contracts WHERE vehicle_id=? AND contract_status NOT IN ('已结清', '已终止')""", (cvid,))
            existing = c.fetchone()
            if existing:
                status_desc = existing['delivery_status'] or existing['contract_status']
                return jsonify({'success': False, 'message': f'车辆（id={cvid}）已有未结清合同（状态: {status_desc}），不能重复签约'}), 400

        if not contract_file:
            return jsonify({'success': False, 'message': '请先上传线下签署的合同文档'}), 400

        c.execute("""
            SELECT COUNT(*) AS cnt, COALESCE(SUM(amount), 0) AS total, COALESCE(AVG(amount), 0) AS avg_amount
            FROM factory_repayments
            WHERE contract_id=?
        """, (planning_contract_id or -1,))
        factory_stats = c.fetchone()
        factory_count = int(factory_stats['cnt'] or 0) if factory_stats else 0
        factory_total = parse_money(factory_stats['total'] if factory_stats else 0)
        factory_avg = parse_money(factory_stats['avg_amount'] if factory_stats else 0)

        contract_type = contract_type_from_sales_mode(order['sales_mode']) if order else contract_type
        today = datetime.now().strftime('%Y-%m-%d')
        start_date = normalize_date(data.get('start_date')) or (normalize_date(order['payment_date']) if order else None) or today
        try:
            repayment_day_default = min(datetime.strptime(start_date, '%Y-%m-%d').day, 28)
        except ValueError:
            repayment_day_default = 1
        repayment_day = int(parse_money(data.get('repayment_day'), repayment_day_default) or repayment_day_default)
        repayment_day = min(max(repayment_day, 1), 28)
        plan_snapshot = {}
        if contract_type == '以租代售' and order and order['snapshot_finance_plan']:
            try:
                plan_snapshot = json.loads(order['snapshot_finance_plan'])
            except (TypeError, ValueError):
                plan_snapshot = {}
        loan_periods = int(parse_money(data.get('loan_periods'), 0) or 0)
        if plan_snapshot:
            loan_periods = parse_period_count(plan_snapshot.get('periods'), factory_count or 12)
        elif loan_periods <= 0:
            loan_periods = parse_period_count(
                order['lease_term'] if order else None,
                factory_count or 12,
            )
        total_price = parse_money(data.get('total_price'), 0)  # sale_total_price 已废弃（以租代售走金融方案）
        rent = parse_money(
            data.get('rent'),
            parse_money(
                plan_snapshot.get('period_price'),
                parse_money(order['vehicle_rent_amount']) if order else 0,
            ),
        )
        if plan_snapshot:
            rent = parse_money(plan_snapshot.get('period_price'), rent)
        if rent <= 0 and total_price > 0 and loan_periods > 0:
            rent = round(total_price / loan_periods, 2)
        monthly_payment = parse_money(data.get('monthly_payment'), factory_avg)
        customer_loan_amount = parse_money(data.get('customer_loan_amount'), 0)
        factory_guarantee_deposit = parse_money(data.get('factory_guarantee_deposit'), 0)
        default_factory_periods = factory_count or loan_periods
        factory_periods = int(parse_money(data.get('factory_periods'), default_factory_periods) or 0)
        factory_repayment_months = int(parse_money(data.get('factory_repayment_months'), factory_count or factory_periods) or 0)
        loan_amount = parse_money(data.get('loan_amount'), factory_total)
        order_payment_amount = parse_money(order['deposit_amount']) if order else 0
        if contract_type == '以租代售':
            deposit = parse_money(data.get('deposit'), 0)
            down_payment = parse_money(
                data.get('down_payment'),
                parse_money(
                    plan_snapshot.get('down_payment'),
                    order_payment_amount or parse_money(order['car_purchase_amount']) if order else 0,
                ),
            )
            if plan_snapshot:
                down_payment = parse_money(plan_snapshot.get('down_payment'), down_payment)
            business_mode = data.get('business_mode') or '以租代售'
        else:
            deposit = parse_money(data.get('deposit'), order_payment_amount)
            down_payment = parse_money(data.get('down_payment'), 0)
            business_mode = data.get('business_mode') or '转租'
        # 20260804：首付已在报单环节完成。从报单带首付状态（足额或特批通过 → 已收/部分已收）
        first_payment_status = '未校验'
        first_payment_shortage = 0
        first_payment_received = 0
        first_payment_screenshot = ''
        if order:
            first_payment_status = order['first_payment_check_status'] or '未校验'
            first_payment_shortage = parse_money(order['first_payment_shortage_amount'])
            first_payment_received = parse_money(order['first_payment_received_amount'])
            first_payment_screenshot = order['customer_screenshot_path'] or ''
        deposit_paid = parse_money(data.get('collected_deposit'), 0)
        down_payment_paid = parse_money(data.get('collected_down_payment'), 0)
        if order and first_payment_received > 0:
            remaining = first_payment_received
            if contract_type == '租赁':
                deposit_paid = min(remaining, deposit) if deposit > 0 else 0
                remaining = round(max(0, remaining - deposit_paid), 2)
            elif contract_type == '以租代售':
                down_payment_paid = min(remaining, down_payment) if down_payment > 0 else 0
                remaining = round(max(0, remaining - down_payment_paid), 2)
        # 租赁：collected_rent 记录首期租金已收部分
        collected_rent_init = 0
        if contract_type == '租赁' and order and first_payment_received > 0:
            remaining = round(max(0, first_payment_received - deposit_paid), 2)
            collected_rent_init = min(remaining, rent) if rent > 0 else 0
        payment_ok = first_payment_status in ('足额', '特批通过')
        # 报单已完成首付核验时可直接进入出库；没有关联报单的手工合同必须
        # 先走首次付款及财务回单审核，避免“待出库”与首次付款入口相互阻塞。
        delivery_status = '待出库' if order and payment_ok else '待首付款'
        deposit_status_val = '免收' if deposit == 0 else ('已收' if (payment_ok and deposit_paid >= deposit) else ('部分已收' if deposit_paid > 0 else '待收'))
        down_payment_status_val = '免收' if down_payment == 0 else ('已收' if (payment_ok and down_payment_paid >= down_payment) else ('部分已收' if down_payment_paid > 0 else '待收'))
        rental_method = data.get('rental_method') or ('经营租赁' if contract_type == '租赁' else contract_type)
        company = (data.get('company') or (order['receiving_company'] if order else '') or '').strip()
        yard = (data.get('yard') or '').strip()
        lease_bank_name = (data.get('lease_bank_name') or '').strip()
        lease_bank_card_no = (data.get('lease_bank_card_no') or '').strip()
        expected_profit_floor = parse_money(data.get('expected_profit_floor'), 0)
        expected_profit_ceiling = normalize_expected_profit_ceiling(data.get('expected_profit_ceiling'))

        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = start_dt + timedelta(days=30 * loan_periods) if loan_periods > 0 else start_dt

        if loan_periods <= 0:
            return jsonify({'success': False, 'message': '客户分期期数必须大于0'}), 400
        if customer_loan_amount > 0:
            min_customer_payment = customer_loan_amount / loan_periods
            if float(rent or 0) < min_customer_payment:
                return jsonify({
                    'success': False,
                    'message': f'每期贷款额不能低于客户贷款额度/期数（至少 ¥{round(min_customer_payment, 2)}）'
                }), 400
        if factory_periods <= 0:
            factory_periods = loan_periods
        if factory_repayment_months <= 0:
            factory_repayment_months = factory_periods

        # 如果提供了客户名但没有 customer_id，自动创建客户
        customer_id = data.get('customer_id')
        customer_name = (data.get('customer_name') or (order['customer_name'] if order else '') or '').strip()
        customer_phone = (data.get('customer_phone') or (order['customer_phone'] if order else '') or '').strip()
        if not customer_id and customer_name:
            c.execute("""
                SELECT id FROM customers
                WHERE name=? AND COALESCE(phone, '')=COALESCE(?, '')
                ORDER BY id DESC LIMIT 1
            """, (customer_name, customer_phone))
            existing_customer = c.fetchone()
            if existing_customer:
                customer_id = existing_customer['id']
            else:
                c.execute("INSERT INTO customers (name, phone) VALUES (?, ?)", (customer_name, customer_phone))
                customer_id = c.lastrowid

        # PRD: 价格快照 — 成交时复制当前基准价至合同
        c.execute("SELECT car_type, purchase_price FROM vehicles WHERE id=?", (vehicle_id,))
        vrow = c.fetchone()
        snap_guidance = resolve_guidance_price_for_vehicle(conn, vrow)[0] if vrow else 0
        snap_invoice = vrow['purchase_price'] if vrow else 0

        # 6.2 更新：合同线下签署，运营上传文档后不再走财务合同审批。
        contract_status = '执行中'

        if planning_contract_id:
            contract_id = planning_contract_id
            c.execute("""
                UPDATE contracts
                SET vehicle_id=?, customer_id=?, contract_type=?, business_mode=?, rental_method=?, repayment_day=?,
                    start_date=?, end_date=?, total_price=?, customer_loan_amount=?, loan_amount=?, monthly_payment=?,
                    rent=?, loan_periods=?, company=?, yard=?, lease_bank_name=?, lease_bank_card_no=?,
                    factory_guarantee_deposit=?, factory_repayment_months=?, factory_periods=?, deposit=?, down_payment=?,
                    down_payment_status=?, deposit_status=?, delivery_status=?,
                    contract_status=?, sales_order_id=?,
                    snapshot_guidance_price=?, snapshot_invoice_price=?,
                    snapshot_finance_plan=?, snapshot_lease_deposit_guidance=?, snapshot_box_monthly_guidance=?,
                    collected_deposit=?, collected_rent=?,
                    expected_profit_floor=?, expected_profit_ceiling=?, contract_file=?, created_by=?
                WHERE id=?
            """, (
                vehicle_id, customer_id, contract_type, business_mode,
                rental_method, repayment_day,
                start_date, end_dt.strftime('%Y-%m-%d'),
                total_price, customer_loan_amount, loan_amount,
                monthly_payment, rent, loan_periods, company, yard, lease_bank_name, lease_bank_card_no, factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                down_payment_status_val,
                deposit_status_val,
                delivery_status,
                contract_status, sales_order_id,
                snap_guidance, snap_invoice,
                order['snapshot_finance_plan'] if order else None,
                parse_money(order['snapshot_lease_deposit_guidance']) if order else 0,
                parse_money(order['snapshot_box_monthly_guidance']) if order else 0,
                deposit_paid,
                collected_rent_init,
                expected_profit_floor, expected_profit_ceiling,
                contract_file,
                user['display_name'],
                contract_id,
            ))
        else:
            c.execute('''
            INSERT INTO contracts (vehicle_id, customer_id, contract_type, business_mode, rental_method, repayment_day,
                                   start_date, end_date, total_price, customer_loan_amount, loan_amount, monthly_payment,
                                   rent, loan_periods, company, yard, lease_bank_name, lease_bank_card_no, factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                                   down_payment_status, deposit_status, delivery_status,
                                   contract_status, sales_order_id,
                                   snapshot_guidance_price, snapshot_invoice_price,
                                   snapshot_finance_plan, snapshot_lease_deposit_guidance, snapshot_box_monthly_guidance,
                                   collected_deposit, collected_rent,
                                   expected_profit_floor, expected_profit_ceiling, contract_file, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                vehicle_id, customer_id, contract_type, business_mode,
                rental_method, repayment_day,
                start_date, end_dt.strftime('%Y-%m-%d'),
                total_price, customer_loan_amount, loan_amount,
                monthly_payment, rent, loan_periods, company, yard, lease_bank_name, lease_bank_card_no, factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                down_payment_status_val,
                deposit_status_val,
                delivery_status,
                contract_status, sales_order_id,
                snap_guidance, snap_invoice,
                order['snapshot_finance_plan'] if order else None,
                parse_money(order['snapshot_lease_deposit_guidance']) if order else 0,
                parse_money(order['snapshot_box_monthly_guidance']) if order else 0,
                deposit_paid,
                collected_rent_init,
                expected_profit_floor, expected_profit_ceiling,
                contract_file,
                user['display_name'],
            ))
            contract_id = c.lastrowid

        # ===== 多车桥表（SKU 改造）：合同关联所有报单车辆，主车 vehicle_id 已写入 contracts =====
        for cvid in contract_vehicles_all:
            c.execute("""
                INSERT OR IGNORE INTO contract_vehicles (contract_id, vehicle_id, is_primary)
                VALUES (?, ?, ?)
            """, (contract_id, cvid, 1 if cvid == vehicle_id else 0))

        comparison = None
        planning_has_financial_activity = False
        if planning_contract_id:
            c.execute("""
                SELECT 1
                FROM repayments
                WHERE contract_id=?
                  AND (COALESCE(paid_amount, 0)>0 OR COALESCE(verified_amount, 0)>0)
                LIMIT 1
            """, (contract_id,))
            planning_has_financial_activity = bool(c.fetchone())
            if not planning_has_financial_activity:
                c.execute("DELETE FROM repayments WHERE contract_id=?", (contract_id,))
                generate_customer_repayment_plan(
                    conn,
                    contract_id,
                    contract_type,
                    start_date,
                    loan_periods,
                    rent,
                    repayment_day=repayment_day,
                    deposit=deposit,
                    down_payment=down_payment,
                )
            if factory_count > 0:
                comparison = compare_contract_repayment_plans(conn, contract_id)
            else:
                c.execute("UPDATE contracts SET customer_plan_match_status='已生成', plan_compare_summary=NULL WHERE id=?", (contract_id,))

        # 车辆状态暂不改变；首次付款审核完成后进入车管出库。
        c.execute("""
            UPDATE approval_flows
            SET status='已取消',
                comment=COALESCE(NULLIF(comment, ''), '合同改为线下签署上传，无需财务审批'),
                acted_at=COALESCE(acted_at, datetime('now','localtime'))
            WHERE ref_type='contract_delivery'
              AND ref_id=?
              AND status='待审批'
        """, (contract_id,))
        if sales_order_id:
            c.execute("""
                UPDATE sales_orders
                SET order_status='已激活',
                    contract_id=?,
                    customer_plan_match_status=?,
                    factory_plan_match_status=?,
                    plan_compare_summary=CASE WHEN ? THEN plan_compare_summary ELSE NULL END
                WHERE id=?
            """, (
                contract_id,
                (comparison or {}).get('status') or '已生成',
                '已上传' if factory_count else '未上传',
                1 if factory_count else 0,
                sales_order_id,
            ))

        # 20260804：首付核销（押金/首付款 period=0 行）+ 应收同步（按 sales_order_id 幂等）
        if order and payment_ok and first_payment_received > 0:
            today = datetime.now().strftime('%Y-%m-%d')
            # 核销 period=0 行（押金/首付款）
            c.execute("""
                UPDATE repayments
                SET status=CASE
                        WHEN MAX(COALESCE(paid_amount, 0), ?) < amount THEN '部分核销'
                        ELSE '已还款'
                    END,
                    paid_amount=MAX(COALESCE(paid_amount, 0), ?),
                    verified_amount=MAX(COALESCE(verified_amount, 0), ?),
                    paid_at=CASE
                        WHEN MAX(COALESCE(paid_amount, 0), ?) >= amount THEN COALESCE(paid_at, ?)
                        ELSE paid_at
                    END,
                    bank_serial=COALESCE(NULLIF(bank_serial,''), ?),
                    screenshot_path=COALESCE(NULLIF(screenshot_path,''), ?),
                    bank_receipt_path=COALESCE(NULLIF(bank_receipt_path,''), ?),
                    verified_by=?,
                    verified_at=?
                WHERE contract_id=? AND period=0
            """, (
                min(first_payment_received, parse_money(deposit) + parse_money(down_payment)),
                min(first_payment_received, parse_money(deposit) + parse_money(down_payment)),
                min(first_payment_received, parse_money(deposit) + parse_money(down_payment)),
                min(first_payment_received, parse_money(deposit) + parse_money(down_payment)),
                today,
                order['finance_bank_serial'] or '',
                first_payment_screenshot or None,
                order['finance_bank_receipt_path'] or None,
                user['display_name'],
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                contract_id,
            ))
            # 租赁报单的首次付款包含“押金 + 首期租金”。第 1 期在计划生成时已存在，
            # 必须同步核销，否则催收/对账会把已收的首期租金再次列为待收。
            if contract_type == '租赁' and collected_rent_init > 0:
                c.execute("""
                    UPDATE repayments
                    SET status=CASE
                            WHEN MAX(COALESCE(paid_amount, 0), ?) >= amount THEN '已还款'
                            ELSE '部分核销'
                        END,
                        paid_amount=MAX(COALESCE(paid_amount, 0), ?),
                        verified_amount=MAX(COALESCE(verified_amount, 0), ?),
                        paid_at=CASE
                            WHEN MAX(COALESCE(paid_amount, 0), ?) >= amount THEN COALESCE(paid_at, ?)
                            ELSE paid_at
                        END,
                        bank_serial=COALESCE(NULLIF(bank_serial,''), ?),
                        screenshot_path=COALESCE(NULLIF(screenshot_path,''), ?),
                        bank_receipt_path=COALESCE(NULLIF(bank_receipt_path,''), ?),
                        verified_by=?,
                        verified_at=?,
                        remark=COALESCE(NULLIF(remark, ''), '销售报单首次付款自动核销首期租金')
                    WHERE contract_id=? AND period=1
                """, (
                    collected_rent_init,
                    collected_rent_init,
                    collected_rent_init,
                    collected_rent_init,
                    today,
                    order['finance_bank_serial'] or '',
                    first_payment_screenshot or None,
                    order['finance_bank_receipt_path'] or None,
                    user['display_name'],
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    contract_id,
                ))
                first_rent_row = c.execute("""
                    SELECT id FROM repayments
                    WHERE contract_id=? AND period=1
                    ORDER BY id ASC LIMIT 1
                """, (contract_id,)).fetchone()
                if first_rent_row:
                    reverse_late_fee_accruals_from_payment_date(
                        conn,
                        first_rent_row['id'],
                        (order['finance_confirmed_at'] or today)[:10],
                        '销售报单首次付款自动核销首期租金',
                    )
            # 首付不足（特批通过）→ 同步应收（按 sales_order_id 幂等，防与审批通过时重复挂账）
            if first_payment_shortage > 0:
                create_or_update_receivable(
                    conn,
                    contract_id,
                    'initial_payment_shortfall',
                    first_payment_shortage,
                    sales_order_id=sales_order_id,
                    promised_repay_date=order['first_payment_promised_date'],
                    reason=order['first_payment_shortage_reason'] or '首付不足（报单特批通过）',
                    status='待归还',
                    created_by=user['display_name'],
                )

        log_audit(conn, '上传线下合同', 'contract', contract_id,
                  f'类型{contract_type} 车辆{vehicle_id} 月租{rent} 月供{monthly_payment} 期数{loan_periods} 合同附件:{contract_file}',
                  user['display_name'])
        conn.commit()
        message = '线下合同已上传，等待车管出库' if delivery_status == '待出库' else '线下合同已上传，请先发起首次付款审核'
        return jsonify({'success': True, 'id': contract_id, 'message': message})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()



# ======================== 客户还款（客户 → 公司）========================
@app.route('/api/contracts/<int:cid>/repayments', methods=['GET'])
def get_repayments(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM repayments WHERE contract_id = ? ORDER BY period ASC", (cid,))
    repayments = [dict(row) for row in c.fetchall()]
    user = get_current_user()
    if user:
        repayments = redact_for_role(conn, user['role'], 'repayments', repayments)
    conn.close()
    return jsonify(repayments)

@app.route('/api/repayments/<int:rid>/confirm', methods=['POST'])
@require_role('财务')
def confirm_repayment(rid):
    """PRD: 财务确认后，按 5.27 规则核销并调整计划表。"""
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    row, gate_message, gate_status = reconciliation_gate_for_repayment(conn, rid)
    if gate_message:
        conn.close()
        return jsonify({'success': False, 'message': gate_message}), gate_status
    if row['status'] == '已还款':
        conn.close()
        return jsonify({'success': False, 'message': '该笔客户还款已核销，请勿重复操作'}), 400
    old_status = row['status']
    amount = row['amount']
    contract_id = row['contract_id']

    received_amount = parse_money(data.get('received_amount'), amount)
    if received_amount <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '到账金额必须大于0'}), 400
    try:
        allocation_result = apply_waterfall_allocation(
            conn,
            rid,
            received_amount,
            request.current_user['display_name'],
            data.get('extra_alloc_periods')
        )
    except ValueError as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': str(e)}), 400
    log_audit(conn, '客户还款核销', 'repayment', rid,
             f'合同{contract_id} 应收{amount} 到账{received_amount} 原状态{old_status} 分配:{allocation_result["summary"]}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '核销成功', 'allocation': allocation_result})


# ======================== 厂家还款（公司 → 一汽解放）========================
@app.route('/api/contracts/<int:cid>/factory-repayments', methods=['GET'])
def get_factory_repayments(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM factory_repayments WHERE contract_id = ? ORDER BY period ASC", (cid,))
    repayments = [dict(row) for row in c.fetchall()]
    user = get_current_user()
    if user:
        repayments = redact_for_role(conn, user['role'], 'factory_repayments', repayments)
    conn.close()
    return jsonify(repayments)

@app.route('/api/factory-repayments/<int:rid>/confirm', methods=['POST'])
@require_role('财务')
def confirm_factory_repayment(rid):
    """PRD: 厂家月供确认后，自动核减贷款余额"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT contract_id, amount, status FROM factory_repayments WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '记录不存在'}), 404
    if row['status'] == '已还款':
        conn.close()
        return jsonify({'success': False, 'message': '该笔厂家月供已核销，请勿重复操作'}), 400
    old_status = row['status']
    amount = row['amount']
    contract_id = row['contract_id']

    c.execute("UPDATE factory_repayments SET status='已还款', paid_at=? WHERE id=?",
              (datetime.now().strftime('%Y-%m-%d'), rid))
    # 联动更新合同已付本金
    c.execute("UPDATE contracts SET paid_principal = paid_principal + ? WHERE id=?",
              (amount, contract_id))
    log_audit(conn, '厂家月供核销', 'factory_repayment', rid,
             f'合同{contract_id} 金额{amount} 原状态{old_status}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '厂家月供核销成功'})


@app.route('/api/contracts/<int:cid>/factory-repayments/import', methods=['POST'])
@require_role('运营', '财务')
def import_factory_repayments(cid):
    data = request.json or {}
    file_url = data.get('file_url', '').strip()
    if not file_url:
        return jsonify({'success': False, 'message': '请先上传厂家还款计划表'}), 400

    local_path = upload_url_to_path(file_url)
    if not os.path.exists(local_path):
        return jsonify({'success': False, 'message': '导入文件不存在，请重新上传'}), 400
    if not (local_path.lower().endswith('.xlsx') or local_path.lower().endswith('.pdf')):
        return jsonify({'success': False, 'message': '当前导入器仅支持 xlsx 或 PDF 格式'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, monthly_payment FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404

    try:
        rows, import_meta = parse_factory_plan_file(local_path)
    except Exception as e:
        conn.close()
        return jsonify({'success': False, 'message': f'解析计划表失败: {e}'}), 400
    import_meta['row_count'] = len(rows)

    fallback_amount = parse_money(contract['monthly_payment'])
    prepared_rows = []
    for row in rows:
        amount = parse_money(row['amount'], fallback_amount)
        if amount <= 0:
            amount = fallback_amount
        prepared_rows.append((cid, row['period'], row['due_date'], amount, '待还款', None, row.get('remark')))

    c.execute("DELETE FROM factory_repayments WHERE contract_id=?", (cid,))
    c.executemany("""
        INSERT INTO factory_repayments (contract_id, period, due_date, amount, status, paid_at, remark)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, prepared_rows)
    factory_total = round(sum(row[3] for row in prepared_rows), 2)
    factory_avg = round(factory_total / len(prepared_rows), 2) if prepared_rows else 0
    c.execute("""
        UPDATE contracts
        SET factory_periods=?,
            factory_repayment_months=?,
            monthly_payment=?,
            loan_amount=CASE WHEN COALESCE(loan_amount, 0) <= 0 THEN ? ELSE loan_amount END
        WHERE id=?
    """, (len(prepared_rows), len(prepared_rows), factory_avg, factory_total, cid))
    comparison = compare_contract_repayment_plans(conn, cid)
    log_audit(conn, '导入厂家还款计划', 'contract', cid,
              f"文件{os.path.basename(local_path)} 格式:{import_meta.get('source_format')} 导入{len(prepared_rows)}期 比对:{comparison.get('status')}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({
        'success': True,
        'message': f'已导入 {len(prepared_rows)} 期厂家还款计划',
        'comparison': comparison,
        'import_meta': import_meta,
    })


@app.route('/api/contracts/<int:cid>/plan-compare', methods=['GET', 'POST'])
@require_role('财务', '运营')
def compare_contract_plans(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM contracts WHERE id=?", (cid,))
    if not c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404
    comparison = compare_contract_repayment_plans(conn, cid)
    log_audit(conn, '还款计划比对', 'contract', cid,
              f"F3 比对状态:{comparison.get('status')}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'comparison': comparison})


@app.route('/api/contracts/<int:cid>/plan-compare/confirm-difference', methods=['POST'])
@require_role('财务')
def confirm_plan_compare_difference(cid):
    data = request.json or {}
    comment = (data.get('comment') or '').strip()
    if not comment:
        return jsonify({'success': False, 'message': '确认差异原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, customer_plan_match_status, plan_compare_summary, sales_order_id FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404

    current = contract['customer_plan_match_status'] or '未比对'
    if current == '未上传':
        conn.close()
        return jsonify({'success': False, 'message': '厂家计划未上传，不能确认差异'}), 400

    summary = {}
    if contract['plan_compare_summary']:
        try:
            summary = json.loads(contract['plan_compare_summary'])
        except ValueError:
            summary = {}
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 比对已通过：利差在合理区间，无需放行差异，这里仅登记一条财务人工确认记录。
    if current == '已通过':
        summary.update({
            'status': '已通过',
            'pass_confirmed_by': request.current_user['display_name'],
            'pass_confirmed_at': now,
            'pass_confirm_comment': comment,
        })
        summary_text = json.dumps(summary, ensure_ascii=False)
        c.execute("UPDATE contracts SET plan_compare_summary=? WHERE id=?", (summary_text, cid))
        if contract['sales_order_id']:
            c.execute("UPDATE sales_orders SET plan_compare_summary=? WHERE id=?",
                      (summary_text, contract['sales_order_id']))
        log_audit(conn, '确认还款计划比对通过', 'contract', cid, comment, request.current_user['display_name'])
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '已记录人工确认，可继续财务确认', 'comparison': summary})

    summary.update({
        'status': '差异已确认',
        'difference_confirmed_by': request.current_user['display_name'],
        'difference_confirmed_at': now,
        'difference_confirm_comment': comment,
    })
    summary_text = json.dumps(summary, ensure_ascii=False)
    c.execute("""
        UPDATE contracts
        SET customer_plan_match_status='差异已确认',
            plan_compare_summary=?
        WHERE id=?
    """, (summary_text, cid))
    if contract['sales_order_id']:
        c.execute("""
            UPDATE sales_orders
            SET customer_plan_match_status='差异已确认',
                factory_plan_match_status='已上传',
                plan_compare_summary=?
            WHERE id=?
        """, (summary_text, contract['sales_order_id']))
    log_audit(conn, '确认还款计划差异', 'contract', cid, comment, request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '计划差异已确认，可继续财务确认', 'comparison': summary})


# ======================== 利润核算 ========================
@app.route('/api/profit/by-vehicle', methods=['GET'])
@require_role('财务', '老板')
def get_profit_by_vehicle():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT v.id AS vehicle_id, v.vin, v.plate_number, v.car_type,
               v.purchase_price, v.tax_rate, v.estimated_residual_value,
               c.id AS contract_id, c.contract_type, c.business_mode, c.rent, c.monthly_payment,
               c.loan_periods, c.deposit, c.down_payment, c.contract_status,
               COALESCE((SELECT SUM(amount) FROM repayments WHERE contract_id=c.id AND period>=1), 0) AS customer_schedule_total,
               COALESCE((SELECT SUM(amount) FROM factory_repayments WHERE contract_id=c.id), 0) AS factory_schedule_total,
               COALESCE((SELECT SUM(COALESCE(paid_amount, 0)) FROM repayments WHERE contract_id=c.id), 0) AS customer_received,
               COALESCE((SELECT SUM(amount) FROM factory_repayments WHERE contract_id=c.id AND status='已还款'), 0) AS factory_paid,
               COALESCE((SELECT SUM(rebate_amount) FROM vehicle_rebates WHERE vehicle_id=v.id), 0) AS rebate_total
        FROM vehicles v
        JOIN contracts c ON c.vehicle_id = v.id
        WHERE (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY v.id ASC, c.id ASC
    """)
    rows = []
    for row in c.fetchall():
        d = dict(row)
        d['planned_single_vehicle_profit'] = round(
            parse_money(d['customer_schedule_total']) - parse_money(d['factory_schedule_total']), 2
        )
        d['realized_cash_profit'] = round(
            parse_money(d['customer_received']) - parse_money(d['factory_paid']) + parse_money(d['rebate_total']), 2
        )
        d['net_profit_with_asset'] = round(
            d['realized_cash_profit'] + parse_money(d['estimated_residual_value']) - parse_money(d['purchase_price']), 2
        )
        d['monthly_spread'] = round(parse_money(d['rent']) - parse_money(d['monthly_payment']), 2)
        d['profit'] = d['realized_cash_profit']
        rows.append(d)
    conn.close()
    return jsonify(rows)


@app.route('/api/profit/summary', methods=['GET'])
@require_role('财务', '老板')
def get_profit_summary():
    rows_response = get_profit_by_vehicle()
    rows = rows_response.get_json()
    summary = {
        'vehicle_count': len(rows),
        'planned_single_vehicle_profit': round(sum(parse_money(r.get('planned_single_vehicle_profit')) for r in rows), 2),
        'realized_cash_profit': round(sum(parse_money(r.get('realized_cash_profit')) for r in rows), 2),
        'net_profit_with_asset': round(sum(parse_money(r.get('net_profit_with_asset')) for r in rows), 2),
        'rebate_total': round(sum(parse_money(r.get('rebate_total')) for r in rows), 2),
        'customer_received': round(sum(parse_money(r.get('customer_received')) for r in rows), 2),
        'factory_paid': round(sum(parse_money(r.get('factory_paid')) for r in rows), 2),
    }
    return jsonify({'success': True, 'summary': summary, 'items': rows})


@app.route('/api/boss-dashboard/overview', methods=['GET'])
@require_role('老板')
def boss_dashboard_overview():
    check_overdue()
    conn = get_db()
    c = conn.cursor()
    metrics = build_dashboard_metrics(conn)
    overview_keys = [
        'vehicle_count', 'active_contract_count', 'open_order_count',
        'overdue_repayment_count', 'pending_invoice_count', 'pending_waiver_count',
        'pending_return_count', 'pending_lock_count', 'rebate_total',
        'customer_received', 'factory_paid', 'realized_cash_profit',
    ]
    overview = {key: metrics.get(key, 0) for key in overview_keys}
    overview['missing_guidance_count'] = unresolved_guidance_vehicle_count(conn)
    overview['data_as_of'] = metrics.get('data_as_of')

    c.execute("""
        SELECT ref_type, required_role, COUNT(*) AS cnt
        FROM approval_flows
        WHERE status='待审批'
        GROUP BY ref_type, required_role
        ORDER BY cnt DESC
    """)
    todos = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({'success': True, 'overview': overview, 'todos': todos})


@app.route('/api/sla/reminders', methods=['GET'])
@require_role('老板', '财务')
def get_sla_reminders():
    conn = get_db()
    c = conn.cursor()
    now = datetime.now()
    reminders = []
    c.execute("""
        SELECT id, ref_type, ref_id, required_role, step_label, created_at
        FROM approval_flows
        WHERE status='待审批'
        ORDER BY created_at ASC
    """)
    for row in c.fetchall():
        created_at = row['created_at']
        try:
            created_dt = datetime.strptime(created_at, '%Y-%m-%d %H:%M:%S')
        except (TypeError, ValueError):
            created_dt = now
        age_hours = (now - created_dt).total_seconds() / 3600
        threshold = 24 if row['ref_type'] == 'price_exception' else 48
        if age_hours >= threshold:
            item = dict(row)
            item['age_hours'] = round(age_hours, 1)
            item['threshold_hours'] = threshold
            item['message'] = f"{row['step_label'] or row['ref_type']} 已超时 {item['age_hours']} 小时"
            reminders.append(item)
    conn.close()
    return jsonify({'success': True, 'count': len(reminders), 'items': reminders})


@app.route('/api/ownership-transfers', methods=['GET'])
@require_role('运营', '财务', '老板')
def list_ownership_transfers():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT ot.*, v.vin, v.plate_number, cu.name AS customer_name
        FROM ownership_transfers ot
        LEFT JOIN vehicles v ON v.id = ot.vehicle_id
        LEFT JOIN contracts c ON c.id = ot.contract_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE (v.is_deleted IS NULL OR v.is_deleted = 0 OR v.id IS NULL)
        ORDER BY ot.id DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/ownership-transfers/pending', methods=['GET'])
@require_role('运营', '财务', '老板')
def list_pending_ownership_transfers():
    """列出已结清且可办理的以租代售过户事项，供首页待办使用。"""
    conn = get_db()
    c = conn.cursor()
    eligibility_sql = """
        NOT EXISTS (
            SELECT 1
            FROM repayments r
            WHERE r.contract_id=c.id
              AND r.period>=0
              AND NOT (
                  r.status IN ('已还款', '预抵')
                  OR COALESCE(r.paid_amount, 0) >= COALESCE(r.amount, 0)
              )
        )
        AND NOT EXISTS (
            SELECT 1
            FROM receivables rv
            WHERE rv.contract_id=c.id
              AND rv.status NOT IN ('已结清', '已取消')
              AND rv.amount > COALESCE(rv.paid_amount, 0)
        )
        AND NOT EXISTS (
            SELECT 1
            FROM contract_fee_items fi
            WHERE fi.contract_id=c.id
              AND fi.amount_due > COALESCE(fi.amount_paid, 0)
        )
        AND NOT EXISTS (
            SELECT 1
            FROM waivers w
            WHERE w.contract_id=c.id
              AND w.status IN ('待审批', '已通过')
        )
    """
    c.execute(f"""
        SELECT *
        FROM (
            SELECT
                c.id AS contract_id,
                c.vehicle_id,
                ot.id AS transfer_id,
                cu.name AS customer_name,
                v.plate_number,
                v.vin,
                c.contract_status,
                ot.status AS transfer_status,
                ot.settle_type,
                '待完成' AS pending_kind,
                ot.created_at
            FROM ownership_transfers ot
            JOIN contracts c ON c.id=ot.contract_id
            LEFT JOIN vehicles v ON v.id=c.vehicle_id
            LEFT JOIN customers cu ON cu.id=c.customer_id
            WHERE ot.status='待过户'
              AND c.contract_type='以租代售'
              AND (v.is_deleted IS NULL OR v.is_deleted=0 OR v.id IS NULL)
              AND {eligibility_sql}

            UNION ALL

            SELECT
                c.id AS contract_id,
                c.vehicle_id,
                NULL AS transfer_id,
                cu.name AS customer_name,
                v.plate_number,
                v.vin,
                c.contract_status,
                NULL AS transfer_status,
                'natural_settle' AS settle_type,
                '待发起' AS pending_kind,
                NULL AS created_at
            FROM contracts c
            LEFT JOIN vehicles v ON v.id=c.vehicle_id
            LEFT JOIN customers cu ON cu.id=c.customer_id
            WHERE c.contract_type='以租代售'
              AND COALESCE(c.contract_status, '') NOT IN ('已结清', '已终止')
              AND (v.is_deleted IS NULL OR v.is_deleted=0 OR v.id IS NULL)
              AND NOT EXISTS (
                  SELECT 1 FROM ownership_transfers ot WHERE ot.contract_id=c.id
              )
              AND {eligibility_sql}
        ) pending
        ORDER BY CASE pending_kind WHEN '待完成' THEN 0 ELSE 1 END, contract_id DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({'success': True, 'count': len(rows), 'items': rows})


@app.route('/api/completion-history', methods=['GET'])
@login_required
def get_completion_history():
    """统一查询已完结的租赁退车与以租代售过户记录。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT
            'rental_return' AS source_type,
            ri.id AS source_id,
            '租赁退车入库' AS completion_type,
            CASE
                WHEN COALESCE(v.status, '')='在库' THEN '车辆已入库（二手车）'
                WHEN COALESCE(v.status, '')='待维修' THEN '已退车结算，待维修（二手车）'
                ELSE '已退车结算（二手车）'
            END AS completion_result,
            COALESCE(NULLIF(ri.paid_out_at, ''), NULLIF(ri.created_at, '')) AS completed_at,
            ri.contract_id,
            ri.vehicle_id,
            COALESCE(cu.name, ri.customer_name, '') AS customer_name,
            cu.phone AS customer_phone,
            COALESCE(v.vin, ri.vin, '') AS vin,
            COALESCE(v.plate_number, ri.plate_number, '') AS plate_number,
            COALESCE(v.car_type, ri.car_type, '') AS car_type,
            COALESCE(c.contract_type, '租赁') AS contract_type,
            COALESCE(c.business_mode, '') AS business_mode,
            COALESCE(v.status, '') AS vehicle_status,
            '二手车' AS vehicle_condition,
            '实际退款' AS amount_label,
            COALESCE(ri.refund_paid_amount, ri.actual_refund, 0) AS amount_value,
            COALESCE(ri.refund_serial, '') AS bank_serial,
            COALESCE(ri.return_reason, '') AS completion_reason,
            COALESCE(ri.paid_out_by, '') AS handled_by,
            COALESCE(ri.needs_repair, 0) AS needs_repair,
            COALESCE(ri.repair_reason, '') AS repair_reason,
            '' AS settle_type,
            '' AS transfer_date,
            '' AS new_owner_name,
            '' AS document_path
        FROM return_inspections ri
        LEFT JOIN contracts c ON c.id=ri.contract_id
        LEFT JOIN customers cu ON cu.id=c.customer_id
        LEFT JOIN vehicles v ON v.id=ri.vehicle_id
        WHERE ri.status IN ('已完成', '已入库')

        UNION ALL

        SELECT
            'ownership_transfer' AS source_type,
            ot.id AS source_id,
            '以租代售过户' AS completion_type,
            '车辆已过户' AS completion_result,
            COALESCE(NULLIF(ot.completed_at, ''), NULLIF(ot.transfer_date, ''), NULLIF(ot.created_at, '')) AS completed_at,
            ot.contract_id,
            ot.vehicle_id,
            COALESCE(cu.name, '') AS customer_name,
            cu.phone AS customer_phone,
            COALESCE(v.vin, '') AS vin,
            COALESCE(v.plate_number, '') AS plate_number,
            COALESCE(v.car_type, '') AS car_type,
            COALESCE(c.contract_type, '以租代售') AS contract_type,
            COALESCE(c.business_mode, '') AS business_mode,
            COALESCE(v.status, '') AS vehicle_status,
            COALESCE(v.condition, '') AS vehicle_condition,
            '合同金额' AS amount_label,
            COALESCE(c.total_price, 0) AS amount_value,
            '' AS bank_serial,
            '' AS completion_reason,
            COALESCE(ot.completed_by, ot.created_by, '') AS handled_by,
            0 AS needs_repair,
            '' AS repair_reason,
            COALESCE(ot.settle_type, '') AS settle_type,
            COALESCE(ot.transfer_date, '') AS transfer_date,
            COALESCE(ot.new_owner_name, '') AS new_owner_name,
            COALESCE(ot.transfer_doc_path, '') AS document_path
        FROM ownership_transfers ot
        LEFT JOIN contracts c ON c.id=ot.contract_id
        LEFT JOIN customers cu ON cu.id=c.customer_id
        LEFT JOIN vehicles v ON v.id=ot.vehicle_id
        WHERE ot.status IN ('已完成', '已过户')
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()

    keyword = (request.args.get('keyword') or '').strip().lower()
    source_type = (request.args.get('source_type') or '').strip()
    date_start = (request.args.get('date_start') or '').strip()
    date_end = (request.args.get('date_end') or '').strip()
    if source_type:
        rows = [row for row in rows if row['source_type'] == source_type]
    if keyword:
        rows = [
            row for row in rows
            if keyword in ' '.join(str(row.get(key) or '').lower() for key in (
                'customer_name', 'customer_phone', 'vin', 'plate_number',
                'car_type', 'contract_id', 'completion_type',
            ))
        ]
    if date_start:
        rows = [row for row in rows if (row.get('completed_at') or '')[:10] >= date_start]
    if date_end:
        rows = [row for row in rows if (row.get('completed_at') or '')[:10] <= date_end]
    rows.sort(key=lambda row: (row.get('completed_at') or '', row.get('source_id') or 0), reverse=True)
    return jsonify(rows)


@app.route('/api/contracts/<int:cid>/early-settlement', methods=['GET', 'POST'])
@require_role('运营', '财务')
def early_settlement(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vehicle_id, contract_type, customer_id FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404
    if contract['contract_type'] != '以租代售':
        conn.close()
        return jsonify({'success': False, 'message': '仅以租代售合同支持提前结清'}), 400
    c.execute("SELECT COUNT(*) AS cnt FROM waivers WHERE contract_id=? AND status IN ('待审批', '已通过')", (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在审核中的减免申请，提前结清前请先关闭'}), 400

    c.execute("""
        SELECT id, period, amount, COALESCE(paid_amount,0) AS paid_amount
        FROM repayments
        WHERE contract_id=? AND period>=0
        ORDER BY period ASC
    """, (cid,))
    repayment_rows = [dict(row) for row in c.fetchall()]
    installment_outstanding = round(sum(max(0, parse_money(r['amount']) - parse_money(r['paid_amount'])) for r in repayment_rows), 2)
    c.execute("""
        SELECT COALESCE(SUM(amount_due - COALESCE(amount_paid, 0)), 0) AS total
        FROM contract_fee_items
        WHERE contract_id=? AND amount_due > COALESCE(amount_paid, 0)
    """, (cid,))
    fee_outstanding = round(parse_money(c.fetchone()['total']), 2)
    c.execute("""
        SELECT COALESCE(SUM(waive_amount), 0) AS total
        FROM waivers
        WHERE contract_id=?
          AND status='已生效'
          AND COALESCE(waiver_kind, '')!='late_fee'
    """, (cid,))
    waiver_credit = round(parse_money(c.fetchone()['total']), 2)
    c.execute("SELECT id, balance FROM customer_prepayments WHERE contract_id=? AND balance>0 ORDER BY id ASC", (cid,))
    prepayment_rows = [dict(row) for row in c.fetchall()]
    prepayment_balance = round(sum(parse_money(row['balance']) for row in prepayment_rows), 2)
    settlement_base = round(max(0, installment_outstanding + fee_outstanding - waiver_credit), 2)
    amount_due = round(max(0, settlement_base - prepayment_balance), 2)
    quote = {
        'success': True,
        'contract_id': cid,
        'outstanding_installments': installment_outstanding,
        'outstanding_fees': fee_outstanding,
        'approved_waiver_credit': waiver_credit,
        'customer_prepayment_balance': prepayment_balance,
        'settlement_base': settlement_base,
        'amount_due': amount_due,
        'note': '以租代售无押金项；未结清首付款作为 period=0 纳入提前结清，已核销部分不重复计算。',
    }
    if request.method == 'GET':
        conn.close()
        return jsonify(quote)

    data = request.json or {}
    bank_serial = (data.get('bank_serial') or '').strip()
    if len(bank_serial) < 4:
        conn.close()
        return jsonify({'success': False, 'message': '银行流水号至少填写4位'}), 400
    customer_screenshot_path = (data.get('customer_screenshot_path') or data.get('screenshot_path') or '').strip()
    received_amount = parse_money(data.get('received_amount'), amount_due)
    if received_amount < amount_due:
        conn.close()
        return jsonify({'success': False, 'message': f'提前结清收款不足，应收 ¥{amount_due}'}), 400
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    prepay_to_consume = min(prepayment_balance, settlement_base)
    remaining_prepay = prepay_to_consume
    for row in prepayment_rows:
        if remaining_prepay <= 0:
            break
        consume = min(remaining_prepay, parse_money(row['balance']))
        c.execute("UPDATE customer_prepayments SET balance=balance-? WHERE id=?", (consume, row['id']))
        remaining_prepay = round(remaining_prepay - consume, 2)
    settlement_pool = round(prepay_to_consume + amount_due, 2)
    c.execute("""
        SELECT *
        FROM contract_fee_items
        WHERE contract_id=? AND amount_due > COALESCE(amount_paid, 0)
        ORDER BY
            CASE fee_type
                WHEN 'insurance_fee' THEN 1
                WHEN 'service_fee' THEN 1
                WHEN 'maintenance_fee' THEN 2
                WHEN 'accident_fee' THEN 2
                WHEN 'penalty_fee' THEN 2
                WHEN 'late_fee' THEN 3
                ELSE 99
            END,
            id ASC
    """, (cid,))
    for item in c.fetchall():
        if settlement_pool <= 0:
            break
        outstanding_fee = max(0, parse_money(item['amount_due']) - parse_money(item['amount_paid']))
        pay_fee = min(settlement_pool, outstanding_fee)
        if pay_fee <= 0:
            continue
        c.execute("UPDATE contract_fee_items SET amount_paid=COALESCE(amount_paid,0)+? WHERE id=?", (pay_fee, item['id']))
        sync_fee_item_status(conn, item['id'])
        settlement_pool = round(settlement_pool - pay_fee, 2)
    for row in repayment_rows:
        due = max(0, parse_money(row['amount']) - parse_money(row['paid_amount']))
        if due <= 0:
            continue
        pay_amount = min(settlement_pool, due)
        if pay_amount <= 0:
            break
        new_paid = round(parse_money(row['paid_amount']) + pay_amount, 2)
        status = '已还款' if new_paid >= parse_money(row['amount']) else '部分核销'
        c.execute("""
            UPDATE repayments
            SET paid_amount=?,
                verified_amount=COALESCE(verified_amount,0)+?,
                status=?,
                paid_at=?,
                bank_serial=COALESCE(NULLIF(bank_serial,''), ?),
                screenshot_path=COALESCE(NULLIF(screenshot_path,''), ?),
                verified_by=COALESCE(NULLIF(verified_by,''), ?),
                verified_at=COALESCE(NULLIF(verified_at,''), ?),
                waterfall_summary=COALESCE(NULLIF(waterfall_summary,''), 'early settlement')
            WHERE id=?
        """, (
            new_paid,
            pay_amount,
            status,
            datetime.now().strftime('%Y-%m-%d'),
            bank_serial,
            customer_screenshot_path,
            request.current_user['display_name'],
            now,
            row['id'],
        ))
        settlement_pool = round(settlement_pool - pay_amount, 2)
    settle_contract_shortfall_receivables(conn, cid)
    extra = round(received_amount - amount_due, 2)
    if extra > 0:
        c.execute("""
            INSERT INTO customer_prepayments
                (contract_id, customer_id, amount, source_bank_serial, balance, remark)
            VALUES (?, ?, ?, ?, ?, '提前结清多收款挂账')
        """, (cid, contract['customer_id'], extra, bank_serial, extra))
    c.execute("""
        UPDATE contracts
        SET loan_balance='已结清',
            contract_status='已提前结清',
            early_settlement_amount=?,
            early_settlement_serial=?,
            early_settled_at=?,
            early_settled_by=?
        WHERE id=?
    """, (received_amount, bank_serial, now, request.current_user['display_name'], cid))
    transfer_id = None
    try:
        c.execute("""
            INSERT INTO ownership_transfers
                (contract_id, vehicle_id, settle_type, status, idempotency_key, created_by, created_at)
            VALUES (?, ?, 'early_settle', '待过户', ?, ?, ?)
        """, (
            cid,
            contract['vehicle_id'],
            f'ownership-{cid}-early-settle',
            request.current_user['display_name'],
            now,
        ))
        transfer_id = c.lastrowid
    except Exception as e:
        if 'UNIQUE constraint failed' not in str(e):
            conn.rollback()
            conn.close()
            return jsonify({'success': False, 'message': str(e)}), 400
        c.execute("SELECT id FROM ownership_transfers WHERE idempotency_key=?", (f'ownership-{cid}-early-settle',))
        existing_transfer = c.fetchone()
        transfer_id = existing_transfer['id'] if existing_transfer else None
    log_audit(conn, '提前结清收款', 'contract', cid,
              f'应收{amount_due} 到账{received_amount} 预收抵扣{prepay_to_consume} 流水{bank_serial}',
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    quote.update({'message': '提前结清已完成，已生成待过户单', 'received_amount': received_amount, 'transfer_id': transfer_id})
    return jsonify(quote)


@app.route('/api/contracts/<int:cid>/ownership-transfer', methods=['POST'])
@require_role('运营')
def create_ownership_transfer(cid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vehicle_id, contract_type, contract_status FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404
    if contract['contract_type'] != '以租代售':
        conn.close()
        return jsonify({'success': False, 'message': '仅以租代售合同可发起过户'}), 400

    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM waivers
        WHERE contract_id=? AND status IN ('待审批', '已通过')
    """, (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在审核中的减免申请，结清/过户前请先关闭'}), 400

    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM repayments
        WHERE contract_id=?
          AND period>=0
          AND NOT (
              status IN ('已还款', '预抵')
              OR COALESCE(paid_amount, 0) >= COALESCE(amount, 0)
          )
    """, (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '客户首付款或分期尚未全部结清，不能发起过户'}), 400
    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM receivables
        WHERE contract_id=?
          AND status NOT IN ('已结清', '已取消')
          AND amount > COALESCE(paid_amount, 0)
    """, (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在未结清挂账应收，不能发起过户'}), 400
    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM contract_fee_items
        WHERE contract_id=?
          AND amount_due > COALESCE(amount_paid, 0)
    """, (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在未结清合同费用，不能发起过户'}), 400

    idempotency_key = (data.get('idempotency_key') or f'ownership-{cid}').strip()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        c.execute("""
            INSERT INTO ownership_transfers
                (contract_id, vehicle_id, settle_type, status, idempotency_key, created_by, created_at)
            VALUES (?, ?, ?, '待过户', ?, ?, ?)
        """, (
            cid,
            contract['vehicle_id'],
            data.get('settle_type') or 'natural_settle',
            idempotency_key,
            request.current_user['display_name'],
            now,
        ))
        transfer_id = c.lastrowid
    except Exception as e:
        if 'UNIQUE constraint failed' not in str(e):
            conn.rollback()
            conn.close()
            return jsonify({'success': False, 'message': str(e)}), 400
        c.execute("SELECT id FROM ownership_transfers WHERE idempotency_key=?", (idempotency_key,))
        transfer_id = c.fetchone()['id']

    log_audit(conn, '发起过户', 'ownership_transfer', transfer_id,
              f'合同{cid} 结清类型:{data.get("settle_type") or "natural_settle"}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': transfer_id, 'message': '过户单已创建'})


@app.route('/api/ownership-transfers/<int:tid>/complete', methods=['POST'])
@require_role('运营')
def complete_ownership_transfer(tid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM ownership_transfers WHERE id=?", (tid,))
    transfer = c.fetchone()
    if not transfer:
        conn.close()
        return jsonify({'success': False, 'message': '过户单不存在'}), 404
    if transfer['status'] in ('已完成', '已过户'):
        conn.close()
        return jsonify({'success': False, 'message': '过户单已完成，请勿重复操作'}), 400
    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM repayments
        WHERE contract_id=?
          AND period>=0
          AND COALESCE(paid_amount, 0) < COALESCE(amount, 0)
    """, (transfer['contract_id'],))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '客户首付款或分期尚未全部结清，不能完成过户'}), 400
    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM receivables
        WHERE contract_id=?
          AND status NOT IN ('已结清', '已取消')
          AND amount > COALESCE(paid_amount, 0)
    """, (transfer['contract_id'],))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在未结清挂账应收，不能完成过户'}), 400
    c.execute("""
        SELECT COUNT(*) AS cnt
        FROM contract_fee_items
        WHERE contract_id=?
          AND amount_due > COALESCE(amount_paid, 0)
    """, (transfer['contract_id'],))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '存在未结清合同费用，不能完成过户'}), 400

    transfer_date = data.get('transfer_date') or datetime.now().strftime('%Y-%m-%d')
    completed_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE ownership_transfers
        SET status='已过户',
            transfer_date=?,
            new_owner_name=?,
            new_owner_id_card=?,
            transfer_doc_path=?,
            completed_by=?,
            completed_at=?
        WHERE id=?
    """, (
        transfer_date,
        (data.get('new_owner_name') or '').strip(),
        (data.get('new_owner_id_card') or '').strip(),
        (data.get('transfer_doc_path') or '').strip(),
        request.current_user['display_name'],
        completed_at,
        tid,
    ))
    c.execute("UPDATE contracts SET contract_status='已结清' WHERE id=?", (transfer['contract_id'],))
    if transfer['vehicle_id']:
        c.execute("UPDATE vehicles SET status='已过户' WHERE id=?", (transfer['vehicle_id'],))
    log_audit(conn, '完成过户', 'ownership_transfer', tid,
              f'过户日期:{transfer_date}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '过户已完成'})


# ======================== 风控逾期概览 ========================
@app.route('/api/risk/overdue', methods=['GET'])
@login_required
def get_overdue():
    check_overdue()
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT r.*, c.vehicle_id, v.vin, v.plate_number, v.car_type,
               cu.name as customer_name, cu.phone as customer_phone
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE (r.status LIKE '逾期%' OR r.status='部分核销')
          AND r.period >= 1
          AND COALESCE(c.contract_file, '')!=''
          AND c.delivery_status='已出库'
          AND (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY r.due_date ASC
    """)
    overdue = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(overdue)

@app.route('/api/risk/factory-overdue', methods=['GET'])
@login_required
def get_factory_overdue():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT fr.*, c.vehicle_id, v.vin, v.plate_number, v.car_type
        FROM factory_repayments fr
        JOIN contracts c ON c.id = fr.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        WHERE fr.status = '逾期'
          AND COALESCE(c.contract_file, '')!=''
          AND (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY fr.due_date ASC
    """)
    overdue = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(overdue)


@app.route('/api/risk/insurance-expiry', methods=['GET'])
@login_required
def get_insurance_expiry():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT id, vin, plate_number, car_type, company, insurance_expiry_date, annual_review_date, status
        FROM vehicles
        WHERE ((insurance_expiry_date IS NOT NULL AND insurance_expiry_date != '')
           OR (annual_review_date IS NOT NULL AND annual_review_date != ''))
          AND (is_deleted IS NULL OR is_deleted = 0)
        ORDER BY insurance_expiry_date ASC, annual_review_date ASC
    """)
    rows = []
    today = datetime.now().date()
    for row in c.fetchall():
        item = dict(row)
        expiry = item.get('insurance_expiry_date')
        annual = item.get('annual_review_date')
        insurance_days = None
        annual_days = None
        if expiry:
            insurance_days = (datetime.strptime(expiry, '%Y-%m-%d').date() - today).days
        if annual:
            annual_days = (datetime.strptime(annual, '%Y-%m-%d').date() - today).days
        item['insurance_days_left'] = insurance_days
        item['annual_review_days_left'] = annual_days
        item['need_warning'] = (insurance_days is not None and insurance_days <= 30) or (annual_days is not None and annual_days <= 30)
        rows.append(item)
    conn.close()
    return jsonify(rows)





# ======================== 账单汇总 ========================
@app.route('/api/bills/pending', methods=['GET'])
def get_pending_bills():
    """获取所有待核销与逾期的账单"""
    check_overdue()
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT r.*, c.vehicle_id, v.vin, v.plate_number, v.car_type,
               cu.name as customer_name, cu.phone as customer_phone,
               'customer' as bill_type
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE (r.status='待还款' OR r.status LIKE '逾期%' OR r.status IN ('临近还款','还款日','部分核销'))
          AND r.period >= 1
          AND COALESCE(c.contract_file, '')!=''
          AND c.delivery_status='已出库'
          AND (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY r.due_date ASC
    """)
    bills = [dict(row) for row in c.fetchall()]
    # T-3 黄色预警：为「临近还款」账单附带流程图 4.1 提醒文案。
    for bill in bills:
        if bill.get('status') == '临近还款':
            log = c.execute(
                "SELECT detail FROM audit_logs WHERE action='T-3还款提醒' AND target_type='repayment' AND target_id=? ORDER BY id DESC LIMIT 1",
                (bill['id'],),
            ).fetchone()
            bill['t3_reminder_text'] = (log['detail'] if log else '') or _t3_reminder_text(
                bill.get('customer_name'), bill.get('plate_number'), bill.get('vin'), bill.get('due_date'),
            )
    conn.close()
    return jsonify(bills)


@app.route('/api/contracts/<int:cid>/fee-items', methods=['GET'])
@login_required
def get_contract_fee_items(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT *
        FROM contract_fee_items
        WHERE contract_id=?
        ORDER BY
            CASE fee_type
                WHEN 'insurance_fee' THEN 1
                WHEN 'service_fee' THEN 1
                WHEN 'maintenance_fee' THEN 2
                WHEN 'accident_fee' THEN 2
                WHEN 'penalty_fee' THEN 2
                WHEN 'late_fee' THEN 3
                ELSE 99
            END,
            id DESC
    """, (cid,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/contracts/<int:cid>/fee-items', methods=['POST'])
@require_role('财务', '运营')
def create_contract_fee_item(cid):
    data = request.json or {}
    fee_type = data.get('fee_type', '').strip()
    if fee_type not in ('insurance_fee', 'service_fee', 'maintenance_fee', 'accident_fee', 'penalty_fee', 'late_fee'):
        return jsonify({'success': False, 'message': '费用类型不正确'}), 400

    amount_due = parse_money(data.get('amount_due'))
    if amount_due <= 0:
        return jsonify({'success': False, 'message': '费用金额必须大于0'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        INSERT INTO contract_fee_items
            (contract_id, fee_type, description, amount_due, due_date, created_by)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        cid, fee_type, data.get('description', '').strip(), amount_due,
        data.get('due_date', '').strip(), request.current_user['display_name']
    ))
    fee_item_id = c.lastrowid
    log_audit(conn, '新增合同费用项', 'contract_fee_item', fee_item_id,
              f"合同{cid} 类型{fee_type} 金额{amount_due}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': fee_item_id, 'message': '费用项已新增'})


# ======================== 对账核销流程（三位一体）========================
@app.route('/api/reconciliation/list', methods=['GET'])
def get_reconciliation_list():
    """获取所有需要对账的还款记录（含已核销的），用于对账单页面"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT r.id, r.contract_id, r.period, r.due_date, r.amount, r.status,
               r.paid_at, r.screenshot_path, r.bank_receipt_path, r.bank_serial,
               r.reported_amount, r.verified_by, r.verified_at, r.paid_amount, r.verified_amount, r.waterfall_summary,
               v.vin, v.plate_number, v.car_type,
               cu.name as customer_name, cu.phone as customer_phone
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE r.period >= 1
          AND r.status!='未激活'
          AND COALESCE(c.contract_file, '')!=''
          AND c.delivery_status='已出库'
        ORDER BY r.due_date DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    # 计算每条记录的核销步骤进度（运营发起对账 → 财务对账核销）
    for row in rows:
        step = 0
        if row.get('screenshot_path'): step = 1
        if row.get('bank_serial'): step = 2
        if row.get('status') == '已还款': step = 3
        row['reconciliation_step'] = step
    conn.close()
    return jsonify(rows)


@app.route('/api/reconciliation/<int:rid>/allocations', methods=['GET'])
@login_required
def get_reconciliation_allocations(rid):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT ra.*, fi.description as fee_description
        FROM reconciliation_allocations ra
        LEFT JOIN contract_fee_items fi ON fi.id = ra.fee_item_id
        WHERE ra.repayment_id=?
        ORDER BY ra.id ASC
    """, (rid,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/reconciliation/<int:rid>/screenshot', methods=['POST'])
@require_role('运营')
def upload_screenshot(rid):
    """运营发起对账：必须填写还款金额并上传支付凭证截图"""
    data = request.json
    screenshot_path = data.get('screenshot_path', '')
    if not screenshot_path:
        return jsonify({'success': False, 'message': '请上传支付凭证图片'}), 400
    reported_amount = parse_money(data.get('reported_amount'), 0)
    if reported_amount <= 0:
        return jsonify({'success': False, 'message': '请填写客户还款金额'}), 400
    conn = get_db()
    c = conn.cursor()
    row, gate_message, gate_status = reconciliation_gate_for_repayment(conn, rid)
    if gate_message:
        conn.close()
        return jsonify({'success': False, 'message': gate_message}), gate_status
    # 一笔运营发起的对账在财务核销前不能被下一笔回款覆盖。
    # 已部分核销时，bank_serial 会保留本次已核销的流水；运营重新上传补款凭证后
    # 会清空该轮核销标识，形成下一轮待核销记录。
    if row['screenshot_path'] and not (row['bank_serial'] or '').strip():
        conn.close()
        return jsonify({
            'success': False,
            'message': '当前对账已发起，正在等待财务核销；请先完成本次核销后再发起补款对账'
        }), 400
    c.execute("""
        UPDATE repayments
        SET screenshot_path=?,
            reported_amount=?,
            bank_receipt_path=NULL,
            bank_serial=NULL,
            verified_by=NULL,
            verified_at=NULL
        WHERE id=?
    """, (screenshot_path, reported_amount, rid))
    log_audit(conn, '发起对账', 'repayment', rid, f'凭证: {screenshot_path} 还款金额: ¥{reported_amount}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '对账已发起，等待财务核销'})


@app.route('/api/reconciliation/<int:rid>/receipt', methods=['POST'])
@require_role('财务')
def upload_receipt(rid):
    """步骤2：财务上传银行回单（可选，有流水号即可核销）"""
    data = request.json
    bank_receipt_path = (data.get('bank_receipt_path') or '').strip()
    bank_serial = (data.get('bank_serial') or '').strip()
    if not bank_receipt_path and len(bank_serial) < 4:
        return jsonify({'success': False, 'message': '请填写银行流水号（至少4位）或上传银行回单'}), 400
    conn = get_db()
    c = conn.cursor()
    row, gate_message, gate_status = reconciliation_gate_for_repayment(conn, rid)
    if gate_message:
        conn.close()
        return jsonify({'success': False, 'message': gate_message}), gate_status
    if not row['screenshot_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先由运营发起对账'}), 400
    if bank_receipt_path:
        c.execute("UPDATE repayments SET bank_receipt_path=? WHERE id=?", (bank_receipt_path, rid))
    log_audit(conn, '上传银行回单', 'repayment', rid, f'回单: {bank_receipt_path or "未上传"} 流水:{bank_serial}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '银行回单上传成功'})


@app.route('/api/reconciliation/<int:rid>/verify', methods=['POST'])
@require_role('财务')
def verify_reconciliation(rid):
    """财务对账：填写到账金额、银行流水号（可附银行回单），自动核销"""
    data = request.json
    bank_serial = data.get('bank_serial', '').strip()
    if not bank_serial:
        return jsonify({'success': False, 'message': '请输入银行流水号'}), 400
    conn = get_db()
    c = conn.cursor()
    row, gate_message, gate_status = reconciliation_gate_for_repayment(conn, rid)
    if gate_message:
        conn.close()
        return jsonify({'success': False, 'message': gate_message}), gate_status
    if row['status'] == '已还款':
        conn.close()
        return jsonify({'success': False, 'message': '该笔账单已核销，请勿重复操作'}), 400
    if not row['screenshot_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先由运营发起对账'}), 400
    if (row['bank_serial'] or '').strip():
        conn.close()
        return jsonify({'success': False, 'message': '本次对账已核销，请先由运营上传补款凭证后再继续核销'}), 400
    if len(bank_serial) < 4:
        conn.close()
        return jsonify({'success': False, 'message': '银行流水号至少填写4位'}), 400

    # 强制按期顺序结清：前序期未结清时禁止录入后续期
    c.execute("""
        SELECT period, status FROM repayments
        WHERE contract_id=? AND period < ? AND status NOT IN ('已还款', '预抵')
        ORDER BY period ASC LIMIT 1
    """, (row['contract_id'], row['period']))
    blocking = c.fetchone()
    if blocking:
        conn.close()
        return jsonify({'success': False, 'message': f'请先结清第{blocking["period"]}期（当前状态：{blocking["status"]}），不能跨期核销'}), 400

    user = request.current_user['display_name']
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    contract_id = row['contract_id']
    amount = row['amount']
    received_amount = parse_money(data.get('received_amount'), amount)
    if received_amount <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '到账金额必须大于0'}), 400

    # 可选：财务上传银行回单
    bank_receipt_path = (data.get('bank_receipt_path') or '').strip()
    if bank_receipt_path:
        c.execute("UPDATE repayments SET bank_receipt_path=? WHERE id=?", (bank_receipt_path, rid))

    # 更新流水号并自动核销
    c.execute("""UPDATE repayments SET bank_serial=?, verified_by=?, verified_at=? WHERE id=?""",
              (bank_serial, user, now, rid))
    try:
        allocation_result = apply_waterfall_allocation(
            conn,
            rid,
            received_amount,
            user,
            data.get('extra_alloc_periods')
        )
    except ValueError as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': str(e)}), 400
    log_audit(conn, '对账核销', 'repayment', rid,
             f'合同{contract_id} 应收{amount} 到账{received_amount} 流水号{bank_serial} 核销人{user} 分配:{allocation_result["summary"]}')
    conn.commit()
    conn.close()
    return jsonify({
        'success': True,
        'message': f'核销成功，流水号: {bank_serial}',
        'allocation': allocation_result,
    })


# ======================== 应收挂账归还（首付不足等差额） ========================
@app.route('/api/receivables', methods=['GET'])
@login_required
def get_receivables_list():
    conn = get_db()
    rows = conn.execute("""
        SELECT rv.*, v.vin, v.plate_number, v.car_type,
               cu.name as customer_name, cu.phone as customer_phone
        FROM receivables rv
        JOIN contracts c ON c.id = rv.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE rv.status != '已结清' AND rv.receivable_type = 'initial_payment_shortfall'
          AND (v.is_deleted IS NULL OR v.is_deleted = 0)
        ORDER BY rv.promised_repay_date ASC
    """).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route('/api/receivables/<int:rid>/screenshot', methods=['POST'])
@require_role('运营')
def receivable_upload_screenshot(rid):
    data = request.json or {}
    path = (data.get('screenshot_path') or '').strip()
    if not path:
        return jsonify({'success': False, 'message': '请上传还款截图'}), 400
    conn = get_db()
    conn.execute("UPDATE receivables SET screenshot_path=? WHERE id=?", (path, rid))
    log_audit(conn, '应收欠款上传截图', 'receivable', rid, f'运营上传还款凭证', request.current_user['display_name'])
    conn.commit(); conn.close()
    return jsonify({'success': True, 'message': '还款截图已上传'})


@app.route('/api/receivables/<int:rid>/settle', methods=['POST'])
@require_role('财务')
def settle_receivable_api(rid):
    data = request.json or {}
    bank_serial = (data.get('bank_serial') or '').strip()
    if not bank_serial:
        return jsonify({'success': False, 'message': '请填写银行流水号'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM receivables WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '记录不存在'}), 404
    if not row['screenshot_path']:
        conn.close()
        return jsonify({'success': False, 'message': '运营还未上传还款截图'}), 400
    outstanding = round(max(0, parse_money(row['amount']) - parse_money(row['paid_amount'])), 2)
    received_amount = parse_money(data.get('received_amount'), outstanding)
    if received_amount <= 0 or received_amount > outstanding:
        conn.close()
        return jsonify({'success': False, 'message': f'归还金额必须大于0且不超过剩余应收 ¥{outstanding}'}), 400
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        settled_amount = settle_receivable_payment(conn, rid, received_amount)
        applied_rows = apply_receivable_payment_to_repayments(
            conn,
            rid,
            settled_amount,
            request.current_user['display_name'],
            bank_serial,
            row['screenshot_path'],
        )
    except ValueError as exc:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': str(exc)}), 400
    c.execute("""
        UPDATE receivables
        SET bank_serial=?, verified_by=?, verified_at=?
        WHERE id=?
    """, (bank_serial, request.current_user['display_name'], now, rid))
    log_audit(conn, '应收欠款核销', 'receivable', rid,
              f"本次归还¥{settled_amount} 剩余¥{round(outstanding-settled_amount, 2)} 流水号{bank_serial} 同步分期:{applied_rows}",
              request.current_user['display_name'])
    conn.commit(); conn.close()
    return jsonify({
        'success': True,
        'message': '欠款已核销' if received_amount >= outstanding else '已登记部分归还',
        'received_amount': settled_amount,
        'remaining_amount': round(outstanding - settled_amount, 2),
        'repayment_allocations': applied_rows,
    })


# ======================== 审计日志查询 ========================
@app.route('/api/audit-logs', methods=['GET'])
def get_audit_logs():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT 100")
    logs = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(logs)





# ======================== 收取首付/押金（财务确认）========================
@app.route('/api/contracts/<int:cid>/collect-payment', methods=['POST'])
@require_role('财务')
def collect_payment(cid):
    data = request.json
    pay_type = data.get('type', 'deposit')  # deposit or down_payment
    received_amount = parse_money(data.get('amount'))
    conn = get_db()
    c = conn.cursor()
    if pay_type == 'down_payment':
        c.execute("SELECT down_payment FROM contracts WHERE id=?", (cid,))
        row = c.fetchone()
        expected = parse_money(row['down_payment']) if row else 0
        if expected > 0 and received_amount < expected:
            conn.close()
            return jsonify({'success': False, 'message': f'首付款到账不足，应收 ¥{round(expected, 2)}'}), 400
        c.execute("UPDATE contracts SET down_payment_status='已收' WHERE id=?", (cid,))
        log_audit(conn, '收取首付', 'contract', cid, f'财务确认收取首付 到账{received_amount or expected}')
    else:
        c.execute("SELECT deposit FROM contracts WHERE id=?", (cid,))
        row = c.fetchone()
        amount = parse_money(row['deposit']) if row else 0
        if amount > 0 and received_amount < amount:
            conn.close()
            return jsonify({'success': False, 'message': f'押金到账不足，应收 ¥{round(amount, 2)}'}), 400
        c.execute("UPDATE contracts SET deposit_status='已收', collected_deposit=? WHERE id=?", (amount, cid))
        log_audit(conn, '收取押金', 'contract', cid, f'财务确认收取押金 应收{amount} 到账{received_amount or amount}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '收款确认成功'})


# ======================== 合同首次付款（运营发起 → 财务收款确认）========================
@app.route('/api/contracts/<int:cid>/initial-payment', methods=['POST'])
@require_role('运营')
def create_initial_payment(cid):
    """线下合同上传后，运营发起首次付款/押金审核。"""
    data = request.json or {}
    user = request.current_user
    screenshot_path = data.get('customer_screenshot_path') or data.get('screenshot_path') or ''
    if not screenshot_path:
        return jsonify({'success': False, 'message': '请先上传客户付款截图'}), 400

    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("SELECT * FROM contracts WHERE id=?", (cid,))
        contract = c.fetchone()
        if not contract:
            return jsonify({'success': False, 'message': '合同不存在'}), 404
        if contract['delivery_status'] not in ('待首付款', '首付已驳回'):
            return jsonify({'success': False, 'message': f'当前状态为{contract["delivery_status"]}，不能发起首次付款'}), 400

        c.execute("""
            SELECT id, status FROM contract_initial_payments
            WHERE contract_id=? AND status IN ('待审批', '审批中', '已通过')
            ORDER BY id DESC LIMIT 1
        """, (cid,))
        existing = c.fetchone()
        if existing:
            return jsonify({'success': False, 'message': f'已有首次付款记录（状态: {existing["status"]}），请勿重复发起'}), 400

        default_amount = get_initial_payment_amount(contract)
        amount = parse_money(data.get('amount'), default_amount)
        reported_amount = parse_money(data.get('reported_amount'), 0)
        if amount <= 0:
            return jsonify({'success': False, 'message': '首次付款金额必须大于0'}), 400

        payment_type = initial_payment_label(contract)
        c.execute("""
            INSERT INTO contract_initial_payments
                (contract_id, payment_type, amount, customer_screenshot_path, status, requested_by,
                 shortage_reason, promised_repay_date, remark)
            VALUES (?, ?, ?, ?, '审批中', ?, ?, ?, ?)
        """, (
            cid,
            payment_type,
            amount,
            screenshot_path,
            user['display_name'],
            (data.get('shortage_reason') or '').strip(),
            normalize_date(data.get('promised_repay_date')),
            data.get('remark', ''),
        ))
        payment_id = c.lastrowid
        if reported_amount > 0:
            c.execute("""
                UPDATE contract_initial_payments
                SET received_amount=?, shortage_amount=?, shortage_status=?
                WHERE id=?
            """, (
                reported_amount,
                round(max(0, amount - reported_amount), 2),
                '待财务核对' if reported_amount < amount else '无欠款',
                payment_id,
            ))
        create_approval_flow(conn, 'initial_payment', payment_id)
        c.execute("UPDATE contracts SET delivery_status='首付审批中' WHERE id=?", (cid,))
        log_audit(conn, '发起首次付款', 'initial_payment', payment_id,
                  f'合同{cid} {payment_type} 金额{amount} 运营{user["display_name"]}', user['display_name'])
        conn.commit()
        return jsonify({'success': True, 'id': payment_id, 'message': '首次付款已提交，等待财务审核收款'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


@app.route('/api/initial-payments/<int:pid>/receipt', methods=['POST'])
@require_role('财务')
def upload_initial_payment_receipt(pid):
    """财务上传公司到账/银行回单（可选，填流水号即可），之后才能审核通过首付款。"""
    data = request.json or {}
    receipt_path = (data.get('bank_receipt_path') or data.get('receipt_path') or '').strip()
    bank_serial = (data.get('bank_serial') or '').strip()
    if not receipt_path and len(bank_serial) < 4:
        return jsonify({'success': False, 'message': '请填写银行流水号（至少4位）或上传公司收款回单'}), 400
    if 'received_amount' not in data:
        return jsonify({'success': False, 'message': '请填写公司实际到账金额'}), 400
    received_amount = parse_money(data.get('received_amount'))
    if received_amount <= 0:
        return jsonify({'success': False, 'message': '公司到账金额必须大于0'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM contract_initial_payments WHERE id=?", (pid,))
    payment = c.fetchone()
    if not payment:
        conn.close()
        return jsonify({'success': False, 'message': '首次付款记录不存在'}), 404
    if payment['status'] == '已通过':
        conn.close()
        return jsonify({'success': False, 'message': '该首次付款已审核通过，不能重复上传回单'}), 400
    expected_amount = parse_money(payment['amount'])
    shortage_amount = round(max(0, expected_amount - received_amount), 2)
    shortage_reason = (data.get('shortage_reason') or payment['shortage_reason'] or '').strip()
    promised_repay_date = normalize_date(data.get('promised_repay_date') or payment['promised_repay_date'])
    if shortage_amount > 0:
        if not shortage_reason:
            conn.close()
            return jsonify({'success': False, 'message': '首次付款不足，请填写不足原因'}), 400
        if not promised_repay_date:
            conn.close()
            return jsonify({'success': False, 'message': '首次付款不足，请填写剩余金额承诺归还时间'}), 400

    c.execute("""
        UPDATE contract_initial_payments
        SET bank_receipt_path=?, bank_serial=?, received_amount=?,
            shortage_amount=?, shortage_reason=?, promised_repay_date=?,
            shortage_status=CASE WHEN ?>0 THEN '待老板审批' ELSE '无欠款' END
        WHERE id=?
    """, (receipt_path, bank_serial, received_amount, shortage_amount, shortage_reason, promised_repay_date, shortage_amount, pid))
    log_audit(conn, '上传首次付款回单', 'initial_payment', pid,
              f'财务上传公司收款回单 {receipt_path} 流水:{bank_serial} 到账:{received_amount}',
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '公司收款回单已上传'})


@app.route('/api/vehicles/<int:vid>/repair/start', methods=['POST'])
@require_role('运营', '车管')
def start_vehicle_repair(vid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status FROM vehicles WHERE id=?", (vid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    if row['status'] == '维修中':
        conn.close()
        return jsonify({'success': False, 'message': '车辆已在维修中'}), 400
    if row['status'] not in ('租赁中', '以租代售', '待维修', '在库'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能进入维修'}), 400
    c.execute("UPDATE vehicles SET pre_repair_status=?, status='维修中' WHERE id=?", (row['status'], vid))
    log_audit(conn, '车辆维修开始', 'vehicle', vid,
              data.get('reason') or f'原状态:{row["status"]}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已进入维修中'})


@app.route('/api/vehicles/<int:vid>/repair/complete', methods=['POST'])
@require_role('运营', '车管')
def complete_vehicle_repair(vid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status, pre_repair_status FROM vehicles WHERE id=?", (vid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    if row['status'] not in ('待维修', '维修中'):
        conn.close()
        return jsonify({'success': False, 'message': '车辆当前不在待维修或维修中状态'}), 400
    previous = row['pre_repair_status'] or '在库'
    return_repair = c.execute("""
        SELECT id FROM return_inspections
        WHERE vehicle_id=? AND needs_repair=1 AND status IN ('已完成', '已入库')
          AND COALESCE(repair_completed_at, '')=''
        ORDER BY id DESC LIMIT 1
    """, (vid,)).fetchone()
    was_return_repair = bool(return_repair)
    repair_completion_note = (data.get('repair_completion_note') or '').strip()
    if was_return_repair and not repair_completion_note:
        conn.close()
        return jsonify({'success': False, 'message': '请填写维修完成情况'}), 400
    repair_cost_raw = data.get('repair_cost')
    if repair_cost_raw in (None, ''):
        repair_cost = 0
    else:
        repair_cost = parse_optional_money(repair_cost_raw)
        if repair_cost is None:
            conn.close()
            return jsonify({'success': False, 'message': '维修费用必须为有效数字'}), 400
    if repair_cost < 0:
        conn.close()
        return jsonify({'success': False, 'message': '维修费用不能小于0'}), 400
    repair_completed_at = (data.get('repair_completed_at') or datetime.now().strftime('%Y-%m-%d')).strip()
    next_status = '在库' if was_return_repair else previous
    if was_return_repair:
        # 租赁退车后的待维修车辆已发生实际使用，维修完成入库后仍须按二手车管理。
        c.execute(
            "UPDATE vehicles SET status=?, condition='二手车', pre_repair_status=NULL WHERE id=?",
            (next_status, vid),
        )
    else:
        c.execute("UPDATE vehicles SET status=?, pre_repair_status=NULL WHERE id=?", (next_status, vid))
    if return_repair:
        c.execute("""
            UPDATE return_inspections
            SET repair_completed_by=?, repair_completed_at=?,
                repair_completion_note=?, repair_cost=?
            WHERE id=?
        """, (
            request.current_user['display_name'],
            repair_completed_at,
            repair_completion_note,
            round(repair_cost, 2),
            return_repair['id'],
        ))
    log_audit(conn, '车辆维修完成', 'vehicle', vid,
              f'维修完成入库，恢复状态:{next_status}；完成日期:{repair_completed_at}；'
              f'维修费用:¥{repair_cost:.2f}；维修结果:{repair_completion_note or "未填写"}',
              request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '维修已完成', 'status': next_status})


# ======================== 退还车辆验收单 ========================
RETURN_FLEET_REQUIRED_FIELDS = [
    ('mileage', '公里数记录'),
    ('body_tire_clean', '车体/备胎清理情况'),
    ('accident_info', '出险情况'),
    ('insurance_surcharge', '保险上浮费支付情况'),
    ('violation_info', '违章情况'),
    ('etc_info', 'ETC情况'),
    ('maintenance_info', '维修保养情况'),
]

RETURN_OPERATOR_MONEY_FIELDS = [
    ('rent_late_fee', '租金延迟支付滞纳金'),
    ('return_late_fee', '退车应支付滞纳金'),
    ('deposit_rent_receivable', '押金应收租金'),
    ('deposit_paid', '押金支付金额'),
    ('total_deduction', '合计扣款'),
    ('actual_refund', '实际应退金额'),
]


def return_fleet_missing_fields(data):
    return [
        label for key, label in RETURN_FLEET_REQUIRED_FIELDS
        if not str(data.get(key) or '').strip()
    ]


def parse_return_operator_amounts(data):
    amounts = {}
    missing = []
    invalid = []
    for key, label in RETURN_OPERATOR_MONEY_FIELDS:
        raw = data.get(key)
        if raw is None or (isinstance(raw, str) and not raw.strip()):
            missing.append(label)
            continue
        try:
            amount = float(raw)
        except (TypeError, ValueError):
            invalid.append(label)
            continue
        if amount < 0:
            invalid.append(label)
            continue
        amounts[key] = round(amount, 2)
    return amounts, missing, invalid


@app.route('/api/return-inspections', methods=['GET'])
def get_return_inspections():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT ri.*, v.plate_number as vehicle_plate_number, v.vin as vehicle_vin, v.car_type as vehicle_car_type,
               v.company as vehicle_company, v.status as vehicle_status, v.condition as vehicle_condition,
               c.company as contract_company, c.yard as contract_yard,
               c.lease_bank_name, c.lease_bank_card_no
        FROM return_inspections ri
        LEFT JOIN vehicles v ON v.id = ri.vehicle_id
        LEFT JOIN contracts c ON c.id = ri.contract_id
        ORDER BY ri.id DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
    for row in rows:
        needs_repair = bool(row.get('needs_repair'))
        repair_completed = bool(str(row.get('repair_completed_at') or '').strip())
        vehicle_status = row.get('vehicle_status') or ''
        return_status = row.get('status') or ''

        # 维修是退车结算完成后的入库动作。退车流程尚未结束时，先将已报修
        # 的车辆显示在维修队列中，并把当前阻塞环节明确交给车管查看。
        row['repair_queue_status'] = ''
        row['repair_queue_actionable'] = False
        if needs_repair and not repair_completed:
            if vehicle_status == '维修中':
                row['repair_queue_status'] = '维修中'
                row['repair_queue_actionable'] = True
            elif vehicle_status == '待维修':
                row['repair_queue_status'] = '待维修入库'
                row['repair_queue_actionable'] = True
            elif return_status == '待运营填写':
                row['repair_queue_status'] = '待运营填写'
            elif return_status == '待财务复核':
                row['repair_queue_status'] = '待财务复核'
            elif return_status == '待领导审批':
                row['repair_queue_status'] = '待领导审批'
            elif return_status == '待出款':
                row['repair_queue_status'] = '待退还押金'
            elif return_status == '已驳回待销售修改':
                row['repair_queue_status'] = '待销售修改'
            else:
                row['repair_queue_status'] = return_status or '退车流程中'
    conn.close()
    return jsonify(rows)


@app.route('/api/return-inspections', methods=['POST'])
@require_role('销售')
def create_return_inspection():
    request_data = request.json or {}
    data = {
        'vehicle_id': request_data.get('vehicle_id'),
        'return_reason': request_data.get('return_reason', '到期退车'),
        'sales_advisor': (request_data.get('sales_advisor') or '').strip(),
        'remark': (request_data.get('remark') or '').strip(),
    }
    conn = get_db()
    c = conn.cursor()
    user = get_current_user()

    # 自动填充车辆信息
    vehicle_id = data.get('vehicle_id')
    if not vehicle_id:
        conn.close()
        return jsonify({'success': False, 'message': '请选择已经出库（租赁）的车辆'}), 400
    if vehicle_id:
        c.execute("""
            SELECT v.*, c.id as cid, c.deposit, c.start_date, c.end_date, c.lease_bank_name, c.lease_bank_card_no,
                   c.company as contract_company, c.yard as contract_yard, cu.name as cust_name
            FROM vehicles v
            LEFT JOIN contracts c ON c.vehicle_id=v.id AND c.contract_status NOT IN ('已结清', '已终止')
            LEFT JOIN customers cu ON cu.id=c.customer_id
            WHERE v.id=?
            ORDER BY c.id DESC
        """, (vehicle_id,))
        vrow = c.fetchone()
        if vrow:
            if (vrow['lock_status'] if 'lock_status' in vrow.keys() else '未锁') != '未锁':
                conn.close()
                return jsonify({'success': False, 'message': '车辆已锁或锁车流程中，不能发起退车'}), 400
            if user and user['role'] in ('销售', '老板') and vrow['status'] not in ('租赁中', '经营租赁'):
                conn.close()
                return jsonify({'success': False, 'message': '仅已经出库（租赁）车辆可以发起退车'}), 400
            if not data.get('plate_number'): data['plate_number'] = vrow['plate_number']
            if not data.get('vin'): data['vin'] = vrow['vin']
            if not data.get('car_type'): data['car_type'] = vrow['car_type']
            if not data.get('company'): data['company'] = vrow['company']
            if not data.get('customer_name'): data['customer_name'] = vrow['cust_name'] or ''
            if not data.get('contract_id'): data['contract_id'] = vrow['cid']
            if not data.get('rental_period') and vrow['start_date'] and vrow['end_date']:
                data['rental_period'] = f"{vrow['start_date']} ~ {vrow['end_date']}"
            data['lease_bank_name'] = vrow['lease_bank_name'] or ''
            data['lease_bank_card_no'] = vrow['lease_bank_card_no'] or ''
            if not data.get('refund_company_name'):
                data['refund_company_name'] = vrow['contract_company'] or vrow['company'] or ''
            if not data.get('refund_bank_name'):
                data['refund_bank_name'] = vrow['lease_bank_name'] or ''
            if not data.get('refund_bank_card_no'):
                data['refund_bank_card_no'] = vrow['lease_bank_card_no'] or ''
        else:
            conn.close()
            return jsonify({'success': False, 'message': '未找到对应车辆'}), 404

    status = '待车管验车'
    sales_status = '已登记'
    fleet_status = '待填写'
    operator_status = '待填写'
    finance_status = '待填写'

    c.execute("""INSERT INTO return_inspections
        (vehicle_id, contract_id, plate_number, customer_name, rental_period, vin, car_type, company,
         lease_bank_name, lease_bank_card_no, refund_company_name, refund_bank_name, refund_bank_card_no,
         return_reason, tool_triangle, tool_vest, tool_extinguisher, tool_wedge, tool_jack,
         doc_license, doc_keys, mileage, body_tire_clean,
         accident_info, insurance_surcharge, violation_info, etc_info, maintenance_info,
         rent_late_fee, return_late_fee, deposit_rent_receivable, deposit_paid,
         total_deduction, actual_refund, needs_repair, repair_reason, remark, status, sales_status, fleet_status, operator_status, finance_status,
         sales_advisor, created_by, inspected_by, inspected_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (vehicle_id, data.get('contract_id'),
         data.get('plate_number',''), data.get('customer_name',''), data.get('rental_period',''),
         data.get('vin',''), data.get('car_type',''), data.get('company',''),
         data.get('lease_bank_name',''), data.get('lease_bank_card_no',''),
         data.get('refund_company_name',''), data.get('refund_bank_name',''), data.get('refund_bank_card_no',''),
         data.get('return_reason','到期退车'),
         0, 0, 0, 0, 0,
         0, 0,
         '', '',
         '', '',
         '', '', '',
         0, 0,
         0, 0,
         0, 0,
         0, '', data.get('remark',''), status, sales_status, fleet_status, operator_status, finance_status,
         data.get('sales_advisor',''),
         user['display_name'] if user else '', '', ''))

    rid = c.lastrowid
    update_return_inspection_status(conn, rid)
    c.execute("UPDATE vehicles SET status='退车中' WHERE id=? AND status IN ('租赁中', '经营租赁')", (vehicle_id,))

    log_audit(conn, '退还车辆验收', 'return_inspection', rid,
              f"车牌{data.get('plate_number','')} 客户{data.get('customer_name','')} 原因{data.get('return_reason','')}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': rid, 'message': '验收单保存成功'})


@app.route('/api/return-inspections/<int:rid>', methods=['PUT'])
@login_required
def update_return_inspection(rid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    user = get_current_user()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404

    if user and user['role'] == '销售' and row['status'] != '已驳回待销售修改':
        conn.close()
        return jsonify({'success': False, 'message': '销售只能修改被老板驳回的退车单'}), 403

    fields = []
    values = []
    editable_fields = ['plate_number','customer_name','rental_period','vin','car_type','company','lease_bank_name','lease_bank_card_no',
                'refund_company_name','refund_bank_name','refund_bank_card_no',
                'return_reason','tool_triangle','tool_vest','tool_extinguisher','tool_wedge','tool_jack',
                'tool_kit','tent_pole','car_wash_fee','body_ad_clean','other_info',
                'doc_license','doc_keys','mileage','body_tire_clean',
                'accident_info','insurance_surcharge','violation_info','etc_info','maintenance_info',
                'rent_late_fee','return_late_fee','deposit_rent_receivable','deposit_paid',
                'total_deduction','actual_refund','needs_repair','repair_reason','remark','sales_advisor']
    if user and user['role'] == '销售':
        editable_fields = ['return_reason', 'remark', 'sales_advisor']
    for key in editable_fields:
        if key in data:
            fields.append(f"{key}=?")
            values.append(data[key])

    if user and user['role'] == '销售':
        # 驳回状态只能保存销售修改；须通过专用重提接口重新进入车管环节。
        fields.append("sales_status=?")
        values.append('待修改')
    if user and user['role'] == '车管':
        fields.append("fleet_status=?")
        values.append('已填写')
        fields.append("status=?")
        values.append('待运营填写')
    if user and user['role'] == '运营':
        fields.append("operator_status=?")
        values.append('已填写')
        fields.append("status=?")
        values.append('待财务复核')
    if user and user['role'] == '财务':
        fields.append("finance_status=?")
        values.append('已填写')
        fields.append("status=?")
        values.append('待领导审批')

    if fields:
        values.append(rid)
        c.execute(f"UPDATE return_inspections SET {','.join(fields)} WHERE id=?", values)
        if not (user and user['role'] == '销售'):
            update_return_inspection_status(conn, rid)
        conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/return-inspections/<int:rid>/fleet', methods=['POST'])
@require_role('车管')
def update_return_fleet(rid):
    data = request.json or {}
    missing = return_fleet_missing_fields(data)
    if missing:
        return jsonify({'success': False, 'message': f"请填写车管验车信息：{'、'.join(missing)}"}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待车管验车':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到车管验车环节'}), 400
    c.execute("""
        UPDATE return_inspections
        SET fleet_status='已填写',
            tool_triangle=?, tool_vest=?, tool_extinguisher=?, tool_wedge=?, tool_jack=?,
            tool_kit=?, tent_pole=?, car_wash_fee=?, body_ad_clean=?,
            doc_license=?, doc_keys=?,
            mileage=?, body_tire_clean=?, accident_info=?, insurance_surcharge=?,
            violation_info=?, etc_info=?, maintenance_info=?,
            other_info=?,
            needs_repair=?, repair_reason=?,
            status='待运营填写'
        WHERE id=?
    """, (
        1 if data.get('tool_triangle') else 0,
        1 if data.get('tool_vest') else 0,
        1 if data.get('tool_extinguisher') else 0,
        1 if data.get('tool_wedge') else 0,
        1 if data.get('tool_jack') else 0,
        1 if data.get('tool_kit') else 0,
        1 if data.get('tent_pole') else 0,
        1 if data.get('car_wash_fee') else 0,
        1 if data.get('body_ad_clean') else 0,
        1 if data.get('doc_license') else 0,
        1 if data.get('doc_keys') else 0,
        data.get('mileage', ''),
        data.get('body_tire_clean', ''),
        data.get('accident_info', ''),
        data.get('insurance_surcharge', ''),
        data.get('violation_info', ''),
        data.get('etc_info', ''),
        data.get('maintenance_info', ''),
        data.get('other_info', ''),
        1 if data.get('needs_repair') else 0,
        data.get('repair_reason', ''),
        rid,
    ))
    update_return_inspection_status(conn, rid)
    log_audit(conn, '车管验车', 'return_inspection', rid, f"车管填写验车单 {data.get('plate_number', '')}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车管验车内容已保存'})


@app.route('/api/return-inspections/<int:rid>/operator', methods=['POST'])
@require_role('运营')
def update_return_operator(rid):
    data = request.json or {}
    amounts, missing, invalid = parse_return_operator_amounts(data)
    if missing:
        return jsonify({'success': False, 'message': f"请填写运营结算信息：{'、'.join(missing)}"}), 400
    if invalid:
        return jsonify({'success': False, 'message': f"运营结算金额必须为不小于0的数字：{'、'.join(invalid)}"}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待运营填写':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到运营填写环节'}), 400
    c.execute("""
        UPDATE return_inspections
        SET operator_status='已填写',
            rent_late_fee=?, return_late_fee=?, deposit_rent_receivable=?, deposit_paid=?,
            total_deduction=?, actual_refund=?, remark=?, status='待财务复核'
        WHERE id=?
    """, (
        amounts['rent_late_fee'],
        amounts['return_late_fee'],
        amounts['deposit_rent_receivable'],
        amounts['deposit_paid'],
        amounts['total_deduction'],
        amounts['actual_refund'],
        data.get('remark', ''),
        rid,
    ))
    update_return_inspection_status(conn, rid)
    log_audit(conn, '运营查车辆数据', 'return_inspection', rid, f"运营填写退车数据 {data.get('plate_number', '')}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '运营数据已保存'})


@app.route('/api/return-inspections/<int:rid>/finance', methods=['POST'])
@require_role('财务')
def update_return_finance(rid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT ri.*, c.lease_bank_name, c.lease_bank_card_no, c.company as contract_company, v.plate_number
        FROM return_inspections ri
        LEFT JOIN contracts c ON c.id = ri.contract_id
        LEFT JOIN vehicles v ON v.id = ri.vehicle_id
        WHERE ri.id=?
    """, (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待财务复核':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到财务复核环节'}), 400
    expected_bank_name = (row['lease_bank_name'] or '').strip()
    expected_bank_card_no = (row['lease_bank_card_no'] or '').strip()
    bank_name = (data.get('refund_bank_name') or '').strip()
    bank_card_no = (data.get('refund_bank_card_no') or '').strip()
    if expected_bank_card_no and bank_card_no and bank_card_no != expected_bank_card_no:
        conn.close()
        return jsonify({'success': False, 'message': '退款银行卡号必须与租赁时银行卡一致'}), 400
    if expected_bank_name and bank_name and bank_name != expected_bank_name:
        conn.close()
        return jsonify({'success': False, 'message': '退款开户银行必须与租赁时开户银行一致'}), 400
    c.execute("""
        UPDATE return_inspections
        SET finance_status='已填写',
            refund_company_name=?, refund_bank_name=?, refund_bank_card_no=?,
            lease_bank_name=COALESCE(lease_bank_name, ?),
            lease_bank_card_no=COALESCE(lease_bank_card_no, ?),
            finance_approved=1,
            finance_approved_by=?,
            finance_approved_at=?,
            status='待领导审批'
        WHERE id=?
    """, (
        data.get('refund_company_name', row['contract_company'] if 'contract_company' in row.keys() else ''),
        bank_name or expected_bank_name,
        bank_card_no or expected_bank_card_no,
        expected_bank_name,
        expected_bank_card_no,
        request.current_user['display_name'],
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        rid,
    ))
    update_return_inspection_status(conn, rid)
    log_audit(conn, '财务复核退车', 'return_inspection', rid, f"财务复核退车单 {row['plate_number']}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '财务复核已保存'})


def return_inspection_to_sales_for_revision(c, rid, operator_name, now, reason):
    """老板驳回退车后，保留单据和车辆占用，回到销售修改并重提。"""
    c.execute("""
        UPDATE return_inspections
        SET status='已驳回待销售修改',
            sales_status='待修改',
            fleet_status='待填写',
            operator_status='待填写',
            finance_status='待填写',
            finance_approved=0,
            finance_approved_by=NULL,
            finance_approved_at=NULL,
            boss_approved=0,
            boss_approved_by=NULL,
            boss_approved_at=NULL,
            rejected_by=?,
            rejected_at=?,
            reject_reason=?
        WHERE id=?
    """, (operator_name, now, reason, rid))
    c.execute("""
        UPDATE vehicles
        SET status='退车中'
        WHERE id=(SELECT vehicle_id FROM return_inspections WHERE id=?)
    """, (rid,))


@app.route('/api/return-inspections/<int:rid>/boss-approve', methods=['POST'])
@require_role('老板')
def boss_approve_return(rid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待领导审批':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到领导审批环节'}), 400
    batch_no = create_approval_flow(conn, 'return_stock', rid)
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE approval_flows
        SET status='已通过',
            operator_id=?,
            operator_name=?,
            comment=?,
            acted_at=?
        WHERE batch_no=? AND required_role='老板'
    """, (
        request.current_user['id'],
        request.current_user['display_name'],
        (request.get_json(silent=True) or {}).get('comment', ''),
        now,
        batch_no,
    ))
    c.execute("""
        UPDATE return_inspections
        SET boss_approved=1,
            boss_approved_by=?,
            boss_approved_at=?,
            status='待出款'
        WHERE id=?
    """, (
        request.current_user['display_name'],
        now,
        rid,
    ))
    update_return_inspection_status(conn, rid)
    log_audit(conn, '退车领导审批', 'return_inspection', rid, f"老板审批通过 {row['plate_number']}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '领导审批已通过'})


@app.route('/api/return-inspections/<int:rid>/boss-reject', methods=['POST'])
@require_role('老板')
def boss_reject_return(rid):
    data = request.get_json(silent=True) or {}
    reason = (data.get('reason') or data.get('comment') or '').strip()
    if not reason:
        return jsonify({'success': False, 'message': '驳回原因不能为空'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT plate_number, status FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待领导审批':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到领导审批环节'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    batch_no = create_approval_flow(conn, 'return_stock', rid)
    c.execute("""
        UPDATE approval_flows
        SET status='已驳回',
            operator_id=?,
            operator_name=?,
            comment=?,
            acted_at=?
        WHERE batch_no=? AND required_role='老板'
    """, (
        request.current_user['id'],
        request.current_user['display_name'],
        reason,
        now,
        batch_no,
    ))
    return_inspection_to_sales_for_revision(
        c,
        rid,
        request.current_user['display_name'],
        now,
        reason,
    )
    log_audit(conn, '退车领导驳回', 'return_inspection', rid,
              f"老板驳回 {row['plate_number']} 原因:{reason}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已驳回，退车单已退回销售修改'})


@app.route('/api/return-inspections/<int:rid>/resubmit', methods=['POST'])
@require_role('销售')
def resubmit_return_inspection(rid):
    data = request.get_json(silent=True) or {}
    note = (data.get('resubmit_note') or '').strip()
    if not note:
        return jsonify({'success': False, 'message': '请填写销售修改说明后再提交'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '已驳回待销售修改':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单不在销售修改重提环节'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE return_inspections
        SET return_reason=COALESCE(NULLIF(?, ''), return_reason),
            sales_advisor=COALESCE(NULLIF(?, ''), sales_advisor),
            remark=COALESCE(NULLIF(?, ''), remark),
            sales_status='已登记',
            fleet_status='待填写',
            operator_status='待填写',
            finance_status='待填写',
            finance_approved=0,
            finance_approved_by=NULL,
            finance_approved_at=NULL,
            boss_approved=0,
            boss_approved_by=NULL,
            boss_approved_at=NULL,
            rejected_by=NULL,
            rejected_at=NULL,
            reject_reason=NULL,
            resubmitted_by=?,
            resubmitted_at=?,
            resubmit_note=?,
            status='待车管验车'
        WHERE id=?
    """, (
        (data.get('return_reason') or '').strip(),
        (data.get('sales_advisor') or '').strip(),
        (data.get('remark') or '').strip(),
        request.current_user['display_name'],
        now,
        note,
        rid,
    ))
    c.execute("""
        UPDATE approval_flows
        SET status='已取消',
            comment=COALESCE(NULLIF(comment, ''), '老板驳回后已由销售重新提交'),
            acted_at=COALESCE(acted_at, ?)
        WHERE ref_type='return_stock' AND ref_id=? AND status='待审批'
    """, (now, rid))
    c.execute("""
        UPDATE vehicles
        SET status='退车中'
        WHERE id=(SELECT vehicle_id FROM return_inspections WHERE id=?)
    """, (rid,))
    log_audit(conn, '退车销售重提', 'return_inspection', rid,
              f'销售修改说明:{note}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已重新提交，等待车管验车'})


@app.route('/api/return-inspections/<int:rid>/pay', methods=['POST'])
@require_role('财务')
def pay_return_refund(rid):
    data = request.get_json(silent=True) or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待出款':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到退还押金环节'}), 400
    if not row['boss_approved']:
        conn.close()
        return jsonify({'success': False, 'message': '请先完成领导审批'}), 400
    refund_due = round(max(0, parse_money(row['actual_refund'])), 2)
    refund_serial = (data.get('refund_serial') or '').strip()
    refund_paid_amount = data.get('refund_paid_amount', refund_due)
    try:
        refund_paid_amount = float(refund_paid_amount)
        if refund_paid_amount < 0:
            raise ValueError
    except (TypeError, ValueError):
        conn.close()
        return jsonify({'success': False, 'message': '请填写有效的退款金额'}), 400
    if refund_due > 0 and not refund_serial:
        conn.close()
        return jsonify({'success': False, 'message': '请填写银行流水号'}), 400
    if round(refund_paid_amount, 2) != refund_due:
        conn.close()
        return jsonify({'success': False, 'message': f'退款金额必须与退车单应退金额一致（应退 ¥{refund_due}）'}), 400
    if row['contract_id']:
        settlement_date = datetime.now().strftime('%Y-%m-%d')
        close_covered_initial_payment_shortfalls(
            conn,
            row['contract_id'],
            request.current_user['display_name'],
            '退车结算前复核历史首期回款',
        )
        c.execute("""
            SELECT COUNT(*) AS cnt
            FROM repayments
            WHERE contract_id=?
              AND COALESCE(paid_amount, 0) < COALESCE(amount, 0)
              AND (
                  period=0
                  OR COALESCE(due_date, '')=''
                  OR due_date<=?
              )
        """, (row['contract_id'], settlement_date))
        if c.fetchone()['cnt'] > 0:
            conn.close()
            return jsonify({'success': False, 'message': '客户仍有已到期未结清首付款或租金，不能完成退车结算'}), 400
        c.execute("""
            SELECT COUNT(*) AS cnt
            FROM receivables
            WHERE contract_id=?
              AND status NOT IN ('已结清', '已取消')
              AND amount > COALESCE(paid_amount, 0)
        """, (row['contract_id'],))
        if c.fetchone()['cnt'] > 0:
            conn.close()
            return jsonify({'success': False, 'message': '存在未结清挂账应收，不能完成退车结算'}), 400
        c.execute("""
            SELECT COUNT(*) AS cnt
            FROM contract_fee_items
            WHERE contract_id=?
              AND amount_due > COALESCE(amount_paid, 0)
        """, (row['contract_id'],))
        if c.fetchone()['cnt'] > 0:
            conn.close()
            return jsonify({'success': False, 'message': '存在未结清合同费用，不能完成退车结算'}), 400
        c.execute("""
            UPDATE repayments
            SET status='已取消',
                remark=CASE
                    WHEN COALESCE(remark, '')='' THEN ?
                    WHEN instr(remark, ?) > 0 THEN remark
                    ELSE remark || '；' || ?
                END
            WHERE contract_id=?
              AND period>=1
              AND COALESCE(paid_amount, 0) < COALESCE(amount, 0)
              AND COALESCE(due_date, '') > ?
              AND status NOT IN ('已还款', '已取消')
        """, (
            f'退车结算取消（退车单#{rid}）',
            f'退车结算取消（退车单#{rid}）',
            f'退车结算取消（退车单#{rid}）',
            row['contract_id'],
            settlement_date,
        ))
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE return_inspections
        SET paid_out=1,
            paid_out_by=?,
            paid_out_at=?,
            refund_serial=?,
            refund_paid_amount=?,
            status='已完成'
        WHERE id=?
    """, (
        request.current_user['display_name'],
        now,
        refund_serial,
        refund_paid_amount,
        rid,
    ))
    if row['vehicle_id']:
        next_vehicle_status = '待维修' if row['needs_repair'] else '在库'
        # 租赁车辆完成退车结算后即使进入维修，也已经发生过实际使用；
        # 成色必须转为二手车，维修完成仅改变库存状态，不能恢复成新车。
        c.execute(
            "UPDATE vehicles SET status=?, condition='二手车' WHERE id=?",
            (next_vehicle_status, row['vehicle_id']),
        )
        returned_vehicle = c.execute(
            "SELECT * FROM vehicles WHERE id=?",
            (row['vehicle_id'],),
        ).fetchone()
        if returned_vehicle:
            ensure_sku_for_vehicle(
                conn,
                dict(returned_vehicle),
                request.current_user['display_name'],
            )
    if row['contract_id']:
        c.execute("UPDATE contracts SET contract_status='已结清', delivery_status='已完成' WHERE id=?", (row['contract_id'],))
    action = '退车退还押金' if refund_due > 0 else '退车结算完成（无退款）'
    log_audit(conn, action, 'return_inspection', rid, f"退车结算完成 {row['plate_number']} 应退¥{refund_due}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '押金已退还，退车完成'})


# ======================== 退车流程合成（审批中心展示用）========================
_RETURN_FLOW_STEPS = [
    {'step': 1, 'role': '销售', 'label': '销售登记'},
    {'step': 2, 'role': '车管', 'label': '车管验车'},
    {'step': 3, 'role': '运营', 'label': '运营查车'},
    {'step': 4, 'role': '财务', 'label': '财务复核'},
    {'step': 5, 'role': '老板', 'label': '领导审批'},
    {'step': 6, 'role': '财务', 'label': '退还押金'},
]

def _synthesize_return_items(conn, user):
    """从 return_inspections 合成审批中心所需的退车流程卡片列表（纯展示，操作在退车验工页面完成）。"""
    c = conn.cursor()
    c.execute("""
        SELECT ri.*, v.vin as vehicle_vin, v.plate_number as vehicle_plate_number, v.car_type as vehicle_car_type
        FROM return_inspections ri
        LEFT JOIN vehicles v ON v.id = ri.vehicle_id
        ORDER BY ri.id DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
    items = []
    for row in rows:
        is_rejected = row.get('status') == '已驳回待销售修改'
        done = {
            1: row.get('sales_status') == '已登记',
            2: row.get('fleet_status') == '已填写',
            3: row.get('operator_status') == '已填写',
            4: bool(row.get('finance_approved')),
            5: bool(row.get('boss_approved')),
            6: bool(row.get('paid_out')),
        }
        actors = {
            1: (row.get('created_by', ''), row.get('created_at', '')),
            2: (row.get('inspected_by', ''), row.get('inspected_at', '')),
            3: ('', ''),
            4: (row.get('finance_approved_by', ''), row.get('finance_approved_at', '')),
            5: (row.get('boss_approved_by', ''), row.get('boss_approved_at', '')),
            6: (row.get('paid_out_by', ''), row.get('paid_out_at', '')),
        }
        current_step = next((cfg['step'] for cfg in _RETURN_FLOW_STEPS if not done[cfg['step']]), 0)
        overall_status = '已驳回' if is_rejected else ('已完成' if current_step == 0 else '审批中')
        steps = [
            {
                'id': 0, 'ref_type': 'return_stock', 'ref_id': row['id'],
                'step_order': cfg['step'], 'required_role': cfg['role'], 'step_label': cfg['label'],
                'status': '已通过' if done[cfg['step']] else '待审批',
                'operator_name': actors[cfg['step']][0], 'acted_at': actors[cfg['step']][1], 'comment': '',
            }
            for cfg in _RETURN_FLOW_STEPS
        ]
        current_role = next((cfg['role'] for cfg in _RETURN_FLOW_STEPS if cfg['step'] == current_step), None)
        is_user_step = (current_role == user['role'])
        is_user_related = (row.get('created_by') == user['display_name'])
        if user['role'] != '老板' and not (is_user_step or is_user_related):
            continue
        items.append({
            'ref_type': 'return_stock', 'ref_id': row['id'],
            'vehicle_id': row.get('vehicle_id'),
            'plate_number': row.get('plate_number') or row.get('vehicle_plate_number', ''),
            'car_type': row.get('car_type') or row.get('vehicle_car_type', ''),
            'vin': row.get('vin') or row.get('vehicle_vin', ''),
            'customer_name': row.get('customer_name', ''),
            'contract_type': '退车入库',
            'created_at': row.get('created_at', ''),
            'return_reason': row.get('return_reason', ''),
            'reject_reason': row.get('reject_reason', ''),
            'delivery_status': row.get('status', ''),
            'requested_by': row.get('created_by', ''),
            'created_by': row.get('created_by', ''),
            'steps': steps, 'current_step': current_step, 'overall_status': overall_status,
            # fill remaining expected fields with defaults
            'total_price': 0, 'rent': 0, 'monthly_payment': 0, 'loan_periods': 0,
            'deposit': 0, 'down_payment': 0, 'start_date': '', 'end_date': '',
            'contract_file': '', 'business_mode': '', 'reason': '', 'overdue_days': 0,
        })
    return items


# ======================== 通用审批流程 API ========================
@app.route('/api/approvals', methods=['GET'])
@login_required
def get_approvals():
    """获取审批列表，支持按 ref_type 和 role 过滤"""
    user = request.current_user
    ref_type = request.args.get('ref_type', '')
    conn = get_db()
    c = conn.cursor()
    params = []
    query = "SELECT * FROM approval_flows"
    if ref_type:
        query += " WHERE ref_type = ?"
        params.append(ref_type)
    query += " ORDER BY id DESC"

    c.execute(query, params)
    rows = [dict(r) for r in c.fetchall()]

    grouped = {}
    latest_batches = {}
    for row in rows:
        key = (row['ref_type'], row['ref_id'])
        if key not in latest_batches:
            latest_batches[key] = row['batch_no']
            grouped[key] = []
        if row['batch_no'] == latest_batches[key]:
            grouped[key].append(row)

    def build_item(base_ref_type, base_ref_id):
        item = {
            'ref_type': base_ref_type,
            'ref_id': base_ref_id,
            'vehicle_id': None,
            'plate_number': '',
            'car_type': '',
            'vin': '',
            'customer_name': '',
            'customer_phone': '',
            'contract_type': '',
            'delivery_status': '',
            'total_price': 0,
            'rent': 0,
            'monthly_payment': 0,
            'loan_periods': 0,
            'deposit': 0,
            'down_payment': 0,
            'start_date': '',
            'end_date': '',
            'contract_file': '',
            'business_mode': '',
            'created_at': '',
            'return_reason': '',
            'reason': '',
            'overdue_days': 0,
            'requested_by': '',
            'snapshot_guidance_price': 0,
            'price_check_status': '',
            'price_exception_reason': '',
            'customer_plan_match_status': '',
            'factory_plan_match_status': '',
            'plan_compare_summary': '',
            'steps': [],
            'current_step': 0,
            'overall_status': '已完成',
        }

        if base_ref_type == 'contract_delivery':
            c.execute("""
                SELECT c.id as contract_id, c.*, v.vin, v.plate_number, v.car_type,
                       cu.name as customer_name, cu.phone as customer_phone
                FROM contracts c
                JOIN vehicles v ON v.id = c.vehicle_id
                LEFT JOIN customers cu ON cu.id = c.customer_id
                WHERE c.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'contract_id': row['contract_id'],
                'vehicle_id': row['vehicle_id'],
                'plate_number': row['plate_number'],
                'car_type': row['car_type'],
                'vin': row['vin'],
                'customer_name': row.get('customer_name', ''),
                'customer_phone': row.get('customer_phone', ''),
                'contract_type': row.get('contract_type', ''),
                'delivery_status': row.get('delivery_status', ''),
                'total_price': row.get('total_price', 0),
                'rent': row.get('rent', 0),
                'monthly_payment': row.get('monthly_payment', 0),
                'loan_periods': row.get('loan_periods', 0),
                'deposit': row.get('deposit', 0),
                'down_payment': row.get('down_payment', 0),
                'start_date': row.get('start_date', ''),
                'end_date': row.get('end_date', ''),
                'contract_file': row.get('contract_file', ''),
                'business_mode': row.get('business_mode', ''),
                'created_at': row.get('created_at', ''),
                'requested_by': row.get('created_by', ''),
                'delivery_photo_path': row.get('delivery_photo_path', ''),
                'delivery_document_path': row.get('delivery_document_path', ''),
            })
            return item

        if base_ref_type in ('price_exception', 'sale_payment', 'order_exception'):
            c.execute("""
                SELECT so.*, v.vin as vehicle_vin, v.plate_number as vehicle_plate_number,
                       v.car_type as vehicle_car_type
                FROM sales_orders so
                LEFT JOIN vehicles v ON v.id = so.vehicle_id
                WHERE so.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'vehicle_id': row.get('vehicle_id'),
                'plate_number': row.get('plate_number') or row.get('vehicle_plate_number', ''),
                'car_type': row.get('car_type') or row.get('vehicle_car_type', ''),
                'vin': row.get('vin') or row.get('vehicle_vin', ''),
                'customer_name': row.get('customer_name', ''),
                'customer_phone': row.get('customer_phone', ''),
                'contract_type': row.get('sales_mode', '销售报单'),
                'delivery_status': row.get('order_status', ''),
                'total_price': 0,  # sale_total_price 已废弃（整车销售下线）
                'snapshot_guidance_price': row.get('snapshot_guidance_price', 0),
                'snapshot_lease_installment_price': row.get('snapshot_lease_installment_price', 0),
                'price_check_status': row.get('price_check_status', ''),
                'price_exception_reason': row.get('price_exception_reason', ''),
                'customer_plan_match_status': row.get('customer_plan_match_status', ''),
                'factory_plan_match_status': row.get('factory_plan_match_status', ''),
                'plan_compare_summary': row.get('plan_compare_summary', ''),
                'business_mode': row.get('sales_mode', ''),
                'created_at': row.get('created_at', ''),
                'requested_by': row.get('created_by', ''),
                # 20260804 大改版字段
                'order_exception_reason': row.get('order_exception_reason', ''),
                'first_payment_received_amount': row.get('first_payment_received_amount', 0),
                'first_payment_shortage_amount': row.get('first_payment_shortage_amount', 0),
                'first_payment_shortage_reason': row.get('first_payment_shortage_reason', ''),
                'promised_repay_date': row.get('first_payment_promised_date', ''),
                'first_payment_check_status': row.get('first_payment_check_status', ''),
                'customer_screenshot_path': row.get('customer_screenshot_path', ''),
                'snapshot_finance_plan': row.get('snapshot_finance_plan', ''),
                'snapshot_lease_deposit_guidance': row.get('snapshot_lease_deposit_guidance', 0),
                'snapshot_box_monthly_guidance': row.get('snapshot_box_monthly_guidance', 0),
                'refund_id': row.get('refund_id'),
                # 关键财务字段：从销售报单字段映射到合同展示字段
                'rent': row.get('vehicle_rent_amount', 0),
                'monthly_payment': row.get('vehicle_rent_amount', 0),
                'loan_periods': int(''.join(filter(str.isdigit, str(row.get('lease_term') or '0'))) or 0),
                'deposit': row.get('deposit_amount', 0),
                'down_payment': row.get('deposit_amount', 0),
                'start_date': row.get('lease_start_date', ''),
                'end_date': '',
                'follow_up_role': '财务' if row.get('order_status') == '待财务确认' else '',
            })
            # 联查退款单状态
            if row.get('refund_id'):
                refund = c.execute("SELECT status, refund_amount FROM order_refunds WHERE id=?", (row['refund_id'],)).fetchone()
                if refund:
                    item['refund_status'] = refund['status']
                    item['refund_amount'] = refund['refund_amount']
            return item

        if base_ref_type in ('initial_payment', 'initial_payment_shortage'):
            c.execute("""
                SELECT ip.id as initial_payment_id, ip.payment_type, ip.amount as initial_payment_amount,
                       ip.received_amount as initial_payment_received_amount,
                       ip.shortage_amount as initial_payment_shortage_amount,
                       ip.shortage_reason, ip.promised_repay_date, ip.shortage_status,
                       ip.customer_screenshot_path, ip.bank_receipt_path, ip.status as initial_payment_status,
                       ip.requested_by, ip.remark as initial_payment_remark, ip.created_at as initial_payment_created_at,
                       c.id as contract_id, c.*, v.vin, v.plate_number, v.car_type,
                       cu.name as customer_name, cu.phone as customer_phone
                FROM contract_initial_payments ip
                JOIN contracts c ON c.id = ip.contract_id
                JOIN vehicles v ON v.id = c.vehicle_id
                LEFT JOIN customers cu ON cu.id = c.customer_id
                WHERE ip.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'contract_id': row['contract_id'],
                'vehicle_id': row['vehicle_id'],
                'plate_number': row['plate_number'],
                'car_type': row['car_type'],
                'vin': row['vin'],
                'customer_name': row.get('customer_name', ''),
                'customer_phone': row.get('customer_phone', ''),
                'contract_type': row.get('contract_type', ''),
                'delivery_status': row.get('delivery_status', ''),
                'total_price': row.get('total_price', 0),
                'rent': row.get('rent', 0),
                'monthly_payment': row.get('monthly_payment', 0),
                'loan_periods': row.get('loan_periods', 0),
                'deposit': row.get('deposit', 0),
                'down_payment': row.get('down_payment', 0),
                'start_date': row.get('start_date', ''),
                'end_date': row.get('end_date', ''),
                'contract_file': row.get('contract_file', ''),
                'business_mode': row.get('business_mode', ''),
                'created_at': row.get('initial_payment_created_at') or row.get('created_at', ''),
                'initial_payment_id': row.get('initial_payment_id'),
                'initial_payment_type': row.get('payment_type', ''),
                'initial_payment_amount': row.get('initial_payment_amount', 0),
                'initial_payment_received_amount': row.get('initial_payment_received_amount', 0),
                'initial_payment_shortage_amount': row.get('initial_payment_shortage_amount', 0),
                'shortage_reason': row.get('shortage_reason', ''),
                'promised_repay_date': row.get('promised_repay_date', ''),
                'shortage_status': row.get('shortage_status', ''),
                'initial_payment_status': row.get('initial_payment_status', ''),
                'customer_screenshot_path': row.get('customer_screenshot_path', ''),
                'bank_receipt_path': row.get('bank_receipt_path', ''),
                'requested_by': row.get('requested_by', ''),
                'initial_payment_remark': row.get('initial_payment_remark', ''),
                'delivery_photo_path': row.get('delivery_photo_path', ''),
                'delivery_document_path': row.get('delivery_document_path', ''),
            })
            return item

        if base_ref_type == 'lock_request':
            c.execute("""
                SELECT lr.*, v.vin, v.plate_number, v.car_type,
                       cu.name as customer_name, cu.phone as customer_phone,
                       c.contract_type, c.business_mode
                FROM lock_requests lr
                JOIN vehicles v ON v.id = lr.vehicle_id
                LEFT JOIN contracts c ON c.id = lr.contract_id
                LEFT JOIN customers cu ON cu.id = c.customer_id
                WHERE lr.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'vehicle_id': row['vehicle_id'],
                'plate_number': row.get('plate_number', ''),
                'car_type': row.get('car_type', ''),
                'vin': row.get('vin', ''),
                'customer_name': row.get('customer_name', ''),
                'customer_phone': row.get('customer_phone', ''),
                'contract_type': row.get('contract_type', ''),
                'business_mode': row.get('business_mode', ''),
                'created_at': row.get('created_at', ''),
                'reason': row.get('reason', ''),
                'overdue_days': row.get('overdue_days', 0),
                'requested_by': row.get('requested_by', ''),
            })
            return item

        if base_ref_type == 'return_stock':
            c.execute("""
                SELECT ri.*, v.vin as vehicle_vin, v.plate_number as vehicle_plate_number, v.car_type as vehicle_car_type
                FROM return_inspections ri
                LEFT JOIN vehicles v ON v.id = ri.vehicle_id
                WHERE ri.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'vehicle_id': row.get('vehicle_id'),
                'plate_number': row.get('plate_number') or row.get('vehicle_plate_number', ''),
                'car_type': row.get('car_type') or row.get('vehicle_car_type', ''),
                'vin': row.get('vin') or row.get('vehicle_vin', ''),
                'customer_name': row.get('customer_name', ''),
                'contract_type': '退车入库',
                'created_at': row.get('created_at', ''),
                'return_reason': row.get('return_reason', ''),
                'delivery_status': row.get('status', ''),
            })
            return item

        if base_ref_type == 'invoice':
            c.execute("""
                SELECT ir.*, c.contract_type, cu.name as customer_name
                FROM invoice_requests ir
                JOIN contracts c ON c.id = ir.contract_id
                LEFT JOIN customers cu ON cu.id = c.customer_id
                WHERE ir.id = ?
            """, (base_ref_id,))
            row = c.fetchone()
            if not row:
                return item
            row = dict(row)
            item.update({
                'contract_id': row.get('contract_id'),
                'customer_name': row.get('customer_name', ''),
                'contract_type': row.get('contract_type', ''),
                'amount': row.get('amount', 0),
                'invoice_entity_name': row.get('invoice_entity_name', ''),
                'invoice_entity_tax_no': row.get('invoice_entity_tax_no', ''),
                'invoice_no': row.get('invoice_no', ''),
                'receiving_company': row.get('receiving_company', ''),
                'applied_by': row.get('applied_by', ''),
                'applied_at': row.get('applied_at', ''),
                'delivery_status': row.get('status', ''),
                'requested_by': row.get('applied_by', ''),
            })
            return item

        return item

    result = []
    for (base_ref_type, base_ref_id), steps in grouped.items():
        # 退车入库改由 return_inspections 合成完整流程（不再依赖只在老板审批时才创建的 approval_flows）
        if base_ref_type == 'return_stock':
            continue
        steps.sort(key=lambda s: s['step_order'])
        item = build_item(base_ref_type, base_ref_id)
        item['steps'] = steps
        for s in steps:
            if s['status'] == '已驳回':
                item['overall_status'] = '已驳回'
                break
            if s['status'] == '待审批' and item['current_step'] == 0:
                item['current_step'] = s['step_order']
                item['overall_status'] = '审批中'
        pending_step = next((s for s in steps if s['status'] == '待审批' and s['step_order'] == item['current_step']), None)
        is_user_step = pending_step and pending_step['required_role'] == user['role']
        is_user_related = (
            item.get('requested_by') == user['display_name']
            or item.get('created_by') == user['display_name']
            or item.get('follow_up_role') == user['role']
        )
        if user['role'] != '老板' and not (is_user_step or is_user_related):
            continue
        result.append(item)

    # 合成退车入库流程（销售一发起即可见，全程 6 步状态展示）
    if not ref_type or ref_type == 'return_stock':
        result.extend(_synthesize_return_items(conn, user))

    result.sort(key=lambda x: x['created_at'] or '', reverse=True)

    conn.close()
    return jsonify(result)


@app.route('/api/approvals/<int:flow_id>/approve', methods=['POST'])
@login_required
def approve_step(flow_id):
    """审批通过"""
    data = request.json or {}
    user = request.current_user
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT * FROM approval_flows WHERE id=?", (flow_id,))
    flow = c.fetchone()
    if not flow:
        conn.close()
        return jsonify({'success': False, 'message': '审批记录不存在'}), 404

    # 校验是否轮到该角色审批
    if flow['status'] != '待审批':
        conn.close()
        return jsonify({'success': False, 'message': '该步骤已处理'}), 400

    # 检查前序步骤是否都已通过
    c.execute("SELECT * FROM approval_flows WHERE batch_no=? AND step_order<? AND status NOT IN ('已通过', '已取消')",
              (flow['batch_no'], flow['step_order']))
    if c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '前序审批未完成'}), 400

    if user['role'] != flow['required_role']:
        conn.close()
        return jsonify({'success': False, 'message': f'需要{flow["required_role"]}角色审批'}), 403

    # 财务确认报单必须先完成厂家分期表上传与比对。这里放在写审批状态之前，
    # 避免审批中心绕过报单列表按钮直接把节点改成已通过。
    if flow['ref_type'] == 'sale_payment' and flow['required_role'] == '财务':
        blocker = sales_order_plan_activation_blocker(conn, flow['ref_id'])
        if blocker:
            conn.close()
            return jsonify({'success': False, 'message': blocker}), 400

    # 首次付款的财务审核必须先上传公司到账/银行回单。
    if flow['ref_type'] == 'initial_payment' and flow['required_role'] == '财务':
        c.execute("SELECT bank_receipt_path, bank_serial, received_amount FROM contract_initial_payments WHERE id=?", (flow['ref_id'],))
        payment = c.fetchone()
        if not payment or not payment['bank_receipt_path']:
            conn.close()
            return jsonify({'success': False, 'message': '请先上传公司收款回单，再进行财务审核'}), 400
        if not payment['bank_serial'] or parse_money(payment['received_amount']) <= 0:
            conn.close()
            return jsonify({'success': False, 'message': '请先登记银行流水号和到账金额，再进行财务审核'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("UPDATE approval_flows SET status='已通过', operator_id=?, operator_name=?, comment=?, acted_at=? WHERE id=?",
              (user['id'], user['display_name'], data.get('comment', ''), now, flow_id))

    # 检查是否所有步骤都已通过 → 触发后续业务逻辑
    c.execute("SELECT COUNT(*) as cnt FROM approval_flows WHERE batch_no=? AND status='待审批'",
              (flow['batch_no'],))
    pending = c.fetchone()['cnt']

    ref_type = flow['ref_type']
    ref_id = flow['ref_id']
    message = '审批通过'

    if pending == 0:
        # 所有步骤都已通过
        if ref_type == 'price_exception':
            c.execute("""
                UPDATE sales_orders
                SET order_status='待财务确认',
                    price_check_status='已通过',
                    boss_price_approved_by=?,
                    boss_price_approved_at=?
                WHERE id=?
            """, (user['display_name'], now, ref_id))
            ensure_default_sales_order_planning_if_needed(conn, ref_id)
            c.execute("""
                SELECT id FROM approval_flows
                WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
                LIMIT 1
            """, (ref_id,))
            if not c.fetchone():
                create_approval_flow(conn, 'sale_payment', ref_id)
            c.execute("""
                INSERT INTO audit_logs (action, target_type, target_id, detail, operator, ip_address)
                VALUES (?,?,?,?,?,?)
            """, (
                '待财务确认报单提醒',
                'sales_order',
                ref_id,
                '低于指导价报单已由老板通过，等待财务确认报单',
                user['display_name'],
                request.headers.get('X-Forwarded-For', request.remote_addr or '')
            ))
            message = '价格特批通过，等待财务确认报单'
        elif ref_type == 'order_exception':
            # 报单异常审批（20260804）：价格异常 + 首付不足合并一次审批
            c.execute("SELECT * FROM sales_orders WHERE id=?", (ref_id,))
            o = c.fetchone()
            shortage = parse_money(o['first_payment_shortage_amount']) if o else 0
            c.execute("""
                UPDATE sales_orders
                SET order_status='待财务确认',
                    price_check_status='已通过',
                    boss_price_approved_by=?,
                    boss_price_approved_at=?,
                    first_payment_check_status=CASE WHEN ? > 0 THEN '特批通过' ELSE first_payment_check_status END
                WHERE id=?
            """, (user['display_name'], now, shortage, ref_id))
            # 首付不足 → 挂应收（承诺期内不计息，复用 receivables 机制）
            if o and shortage > 0:
                contract_id = ensure_sales_order_planning_contract(conn, ref_id)
                if contract_id:
                    create_or_update_receivable(
                        conn,
                        contract_id,
                        'initial_payment_shortfall',
                        shortage,
                        sales_order_id=ref_id,
                        promised_repay_date=o['first_payment_promised_date'],
                        reason=o['first_payment_shortage_reason'] or '首付不足（报单特批通过）',
                        status='待归还',
                        created_by=user['display_name'],
                    )
            ensure_default_sales_order_planning_if_needed(conn, ref_id)
            c.execute("""
                SELECT id FROM approval_flows
                WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
                LIMIT 1
            """, (ref_id,))
            if not c.fetchone():
                create_approval_flow(conn, 'sale_payment', ref_id)
            message = '报单异常审批通过，等待财务确认报单'
        elif ref_type == 'contract_delivery':
            # 流程图要求：销售发起的合同（租赁/销售/以租代售）审批全走完后，
            # 不能直接出库，必须先由运营发起首次付款。
            c.execute("UPDATE contracts SET delivery_status='待首付款' WHERE id=?", (ref_id,))
            message = '合同审批通过，请运营发起首次付款/押金审核'
        elif ref_type == 'sale_payment':
            blocker = sales_order_plan_activation_blocker(conn, ref_id)
            if blocker:
                conn.rollback()
                conn.close()
                return jsonify({'success': False, 'message': blocker}), 400
            bank_serial = (data.get('bank_serial') or '').strip()
            bank_receipt_path = (data.get('bank_receipt_path') or data.get('receipt_path') or '').strip()
            c.execute("""
                SELECT finance_bank_serial, finance_bank_receipt_path
                FROM sales_orders
                WHERE id=?
            """, (ref_id,))
            payment_proof = c.fetchone()
            if not payment_proof:
                conn.rollback()
                conn.close()
                return jsonify({'success': False, 'message': '销售报单不存在'}), 404
            bank_serial = bank_serial or (payment_proof['finance_bank_serial'] or '').strip()
            bank_receipt_path = bank_receipt_path or (payment_proof['finance_bank_receipt_path'] or '').strip()
            if len(bank_serial) < 4 and not bank_receipt_path:
                conn.rollback()
                conn.close()
                return jsonify({
                    'success': False,
                    'message': '请填写银行流水号（至少4位）或上传公司收款回单',
                }), 400
            c.execute("""
                UPDATE sales_orders
                SET order_status='已激活',
                    finance_confirmed_by=?,
                    finance_confirmed_at=?,
                    finance_bank_serial=?,
                    finance_bank_receipt_path=?
                WHERE id=? AND order_status='待财务确认'
            """, (user['display_name'], now, bank_serial or None, bank_receipt_path or None, ref_id))
            log_audit(
                conn,
                '确认销售报单',
                'sales_order',
                ref_id,
                f"审批中心财务确认报单 流水号:{bank_serial or '未填'} 回单:{bank_receipt_path or '未传'}",
                user['display_name'],
            )
            message = '报单已确认，等待运营上传线下合同'
        elif ref_type == 'initial_payment':
            result = finalize_initial_payment(conn, ref_id, user['display_name'], now, allow_shortage=False)
            if result.get('needs_shortage_approval'):
                c.execute("UPDATE contract_initial_payments SET status='待老板审批', shortage_status='待老板审批' WHERE id=?", (ref_id,))
                c.execute("""
                    UPDATE contracts
                    SET delivery_status='首付不足待审批'
                    WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
                """, (ref_id,))
                create_approval_flow(conn, 'initial_payment_shortage', ref_id)
                message = f'首次付款到账不足 ¥{round(result["shortage_amount"], 2)}，已提交老板审批是否允许不足额出库'
            else:
                message = '首次付款审核完成，等待车管出库'
        elif ref_type == 'initial_payment_shortage':
            finalize_initial_payment(conn, ref_id, user['display_name'], now, allow_shortage=True)
            message = '老板已同意不足额出库，差额已挂账应收，等待车管出库'
        elif ref_type == 'return_stock':
            # 退车领导审批通过 → 待财务退还押金
            c.execute("SELECT vehicle_id FROM return_inspections WHERE id=?", (ref_id,))
            ri = c.fetchone()
            if ri:
                c.execute("""
                    UPDATE return_inspections
                    SET boss_approved=1,
                        boss_approved_by=?,
                        boss_approved_at=?,
                        status='待出款'
                    WHERE id=?
                """, (user['display_name'], now, ref_id))
            message = '领导审批通过，等待财务退还押金'
        elif ref_type == 'invoice':
            # 发票审批通过 → 进入待开票，财务可开票
            c.execute("""
                UPDATE invoice_requests
                SET status='待开票',
                    boss_approved_by=?,
                    boss_approved_at=?
                WHERE id=?
            """, (user['display_name'], now, ref_id))
            message = '发票审批通过，等待财务开票'
    else:
        # 更新合同状态显示当前审批进度
        if ref_type == 'contract_delivery':
            c.execute("UPDATE contracts SET delivery_status='审批中' WHERE id=?", (ref_id,))
        elif ref_type == 'price_exception':
            c.execute("UPDATE sales_orders SET order_status='待价格特批' WHERE id=?", (ref_id,))
        elif ref_type == 'initial_payment':
            c.execute("""
                UPDATE contracts
                SET delivery_status='首付审批中'
                WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
            """, (ref_id,))
        elif ref_type == 'initial_payment_shortage':
            c.execute("""
                UPDATE contracts
                SET delivery_status='首付不足待审批'
                WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
            """, (ref_id,))

    log_audit(conn, '审批通过', ref_type, ref_id,
              f'{user["display_name"]}({user["role"]}) 通过 {flow["step_label"]}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': message})


# 报单 → 关联车辆 id 列表（SKU 改造：多车支持）
def order_vehicle_ids(conn, order):
    """返回报单关联的所有车辆 id 列表（含主车）。"""
    if not order:
        return []
    try:
        ids = json.loads(order['vehicle_ids']) if order['vehicle_ids'] else []
    except (TypeError, ValueError):
        ids = []
    if ids:
        return [int(i) for i in ids]
    return [order['vehicle_id']] if order['vehicle_id'] else []


@app.route('/api/approvals/<int:flow_id>/reject', methods=['POST'])
@login_required
def reject_step(flow_id):
    """审批驳回 — 当前步骤标记驳回，后续步骤标记已取消，父实体回退"""
    data = request.json or {}
    user = request.current_user
    comment = data.get('comment', '')
    if not comment:
        return jsonify({'success': False, 'message': '驳回原因不能为空'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM approval_flows WHERE id=?", (flow_id,))
    flow = c.fetchone()
    if not flow:
        conn.close()
        return jsonify({'success': False, 'message': '审批记录不存在'}), 404
    if flow['status'] != '待审批':
        conn.close()
        return jsonify({'success': False, 'message': '该步骤已处理'}), 400
    if user['role'] != flow['required_role']:
        conn.close()
        return jsonify({'success': False, 'message': f'需要{flow["required_role"]}角色操作'}), 403

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 当前步骤标记驳回
    c.execute("UPDATE approval_flows SET status='已驳回', operator_id=?, operator_name=?, comment=?, acted_at=? WHERE id=?",
              (user['id'], user['display_name'], comment, now, flow_id))
    # 后续步骤标记已取消
    c.execute("UPDATE approval_flows SET status='已取消' WHERE batch_no=? AND step_order>?",
              (flow['batch_no'], flow['step_order']))

    ref_type = flow['ref_type']
    ref_id = flow['ref_id']
    # 更新父实体状态
    if ref_type == 'order_exception':
        # 20260804：驳回后报单作废，但车辆保持'报单锁定中'等待退款闭环（退款完成才回在库）
        c.execute("SELECT vehicle_id FROM sales_orders WHERE id=?", (ref_id,))
        order = c.fetchone()
        c.execute("""
            UPDATE sales_orders
            SET order_status='已作废',
                price_check_status='已驳回',
                first_payment_check_status='已驳回',
                voided_at=?,
                void_reason=?,
                voided_by=?
            WHERE id=?
        """, (now, comment, user['display_name'], ref_id))
        # 取消同单其他待审批流（sale_payment 等）
        c.execute("""
            UPDATE approval_flows SET status='已取消', acted_at=COALESCE(acted_at, ?)
            WHERE ref_type IN ('order_exception', 'sale_payment') AND ref_id=? AND status='待审批'
        """, (now, ref_id))
    elif ref_type == 'price_exception':
        c.execute("SELECT vehicle_id, vehicle_ids FROM sales_orders WHERE id=?", (ref_id,))
        order = c.fetchone()
        c.execute("""
            UPDATE sales_orders
            SET order_status='已作废',
                price_check_status='已驳回',
                voided_at=?,
                void_reason=?,
                voided_by=?
            WHERE id=?
        """, (now, comment, user['display_name'], ref_id))
        if order:
            for vid in order_vehicle_ids(conn, order):
                c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (vid,))
    elif ref_type == 'sale_payment':
        c.execute("SELECT vehicle_id, vehicle_ids, contract_id FROM sales_orders WHERE id=?", (ref_id,))
        order = c.fetchone()
        c.execute("""
            UPDATE sales_orders
            SET order_status='已作废',
                voided_at=?,
                void_reason=?,
                voided_by=?
            WHERE id=?
        """, (now, comment, user['display_name'], ref_id))
        if order:
            for vid in order_vehicle_ids(conn, order):
                c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (vid,))
            # 终止报单计划合同壳（SKU 改造：多车报单的合同壳也需清理）
            if order['contract_id']:
                c.execute("""
                    UPDATE contracts SET contract_status='已终止', delivery_status='已终止'
                    WHERE id=? AND contract_status='报单计划中'
                """, (order['contract_id'],))
    elif ref_type == 'contract_delivery':
        c.execute("UPDATE contracts SET delivery_status='已驳回' WHERE id=?", (ref_id,))
    elif ref_type == 'initial_payment':
        c.execute("UPDATE contract_initial_payments SET status='已驳回' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE contracts
            SET delivery_status='首付已驳回'
            WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
        """, (ref_id,))
    elif ref_type == 'initial_payment_shortage':
        c.execute("""
            SELECT ip.contract_id, c.vehicle_id, c.sales_order_id
            FROM contract_initial_payments ip
            JOIN contracts c ON c.id = ip.contract_id
            WHERE ip.id=?
        """, (ref_id,))
        payment = c.fetchone()
        c.execute("UPDATE contract_initial_payments SET status='已驳回', shortage_status='老板驳回' WHERE id=?", (ref_id,))
        if payment:
            c.execute("""
                UPDATE contracts
                SET delivery_status='已终止', contract_status='已终止'
                WHERE id=?
            """, (payment['contract_id'],))
            # 释放合同关联的所有车辆（SKU 改造：多车）
            cv = c.execute("SELECT vehicle_id FROM contract_vehicles WHERE contract_id=?", (payment['contract_id'],)).fetchall()
            if cv:
                for r in cv:
                    c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status IN ('报单锁定中','待出库')", (r['vehicle_id'],))
            elif payment['vehicle_id']:
                c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status IN ('报单锁定中','待出库')", (payment['vehicle_id'],))
            if payment['sales_order_id']:
                c.execute("""
                    UPDATE sales_orders
                    SET order_status='已作废', voided_at=?, void_reason=?, voided_by=?
                    WHERE id=?
                """, (now, comment or '首次付款不足老板驳回，退款终止', user['display_name'], payment['sales_order_id']))
    elif ref_type == 'return_stock':
        return_inspection_to_sales_for_revision(c, ref_id, user['display_name'], now, comment)
    elif ref_type == 'invoice':
        c.execute("""
            UPDATE invoice_requests
            SET status='已驳回',
                voided_by=?,
                voided_at=?,
                void_reason=?
            WHERE id=?
        """, (user['display_name'], now, comment or '发票审批驳回', ref_id))

    log_audit(conn, '审批驳回', ref_type, ref_id,
              f'{user["display_name"]}({user["role"]}) 驳回 {flow["step_label"]} 原因:{comment}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'已驳回: {comment}'})


@app.route('/api/approvals/<int:ref_id>/resubmit', methods=['POST'])
@login_required
def resubmit_approval(ref_id):
    """驳回后重新提交审批"""
    data = request.json or {}
    ref_type = data.get('ref_type', 'contract_delivery')
    conn = get_db()
    c = conn.cursor()

    if ref_type == 'price_exception':
        conn.close()
        return jsonify({'success': False, 'message': '已作废报单不可重提，请复制为草稿后重新提交'}), 400
    elif ref_type == 'contract_delivery':
        c.execute("UPDATE contracts SET delivery_status='待首付款' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE approval_flows
            SET status='已取消',
                comment=COALESCE(NULLIF(comment, ''), '合同无需重新审批，已恢复到首次付款环节'),
                acted_at=COALESCE(acted_at, datetime('now','localtime'))
            WHERE ref_type='contract_delivery' AND ref_id=? AND status='待审批'
        """, (ref_id,))
        log_audit(conn, '恢复合同首次付款', ref_type, ref_id, '合同无需重新审批')
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': '合同无需重新审批，已恢复到首次付款环节'})
    elif ref_type == 'sale_payment':
        c.execute("UPDATE contracts SET delivery_status='待审批' WHERE id=?", (ref_id,))
    elif ref_type == 'initial_payment':
        c.execute("UPDATE contract_initial_payments SET status='审批中' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE contracts
            SET delivery_status='首付审批中'
            WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
        """, (ref_id,))
    elif ref_type == 'return_stock':
        conn.close()
        return jsonify({'success': False, 'message': '退车驳回后请由销售在退车单中修改并重新提交'}), 400
    elif ref_type == 'invoice':
        # 发票驳回后可重提：状态回到待审批，重新走老板审批
        c.execute("UPDATE invoice_requests SET status='待审批' WHERE id=?", (ref_id,))

    create_approval_flow(conn, ref_type, ref_id)
    log_audit(conn, '重新提交审批', ref_type, ref_id, '驳回后重新提交')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已重新提交审批'})


# ======================== J3 滞纳金减免审批 ========================
@app.route('/api/contracts/<int:cid>/waivers', methods=['POST'])
@require_role('销售')
def create_late_fee_waiver(cid):
    """J3/K ① 销售发起减免申请 → waivers(status='待审批')，推送老板待办。
    输入：target_period_list(JSON 数组或 'all')、waive_amount(全额减免时为空)、reason、attachment_path。"""
    data = request.json or {}
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM contracts WHERE id=?", (cid,))
    if not c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404

    reason = (data.get('reason') or '').strip()
    if not reason:
        conn.close()
        return jsonify({'success': False, 'message': '减免原因不能为空'}), 400

    waiver_kind = data.get('waiver_kind') or 'late_fee'
    if waiver_kind not in ('late_fee', 'rent'):
        conn.close()
        return jsonify({'success': False, 'message': '减免类型只能为 late_fee 或 rent'}), 400

    periods = data.get('target_period_list')
    if isinstance(periods, list):
        period_text = json.dumps(periods, ensure_ascii=False)
    elif periods in ('all', '全部', None, ''):
        period_text = 'all'
    else:
        period_text = str(periods)

    waive_amount = data.get('waive_amount')
    waive_amount = parse_money(waive_amount) if waive_amount not in (None, '') else None

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        INSERT INTO waivers
            (contract_id, waiver_kind, target_period_list, waive_amount, reason,
             attachment_path, status, sales_applied_by, sales_applied_at)
        VALUES (?, ?, ?, ?, ?, ?, '待审批', ?, ?)
    """, (cid, waiver_kind, period_text, waive_amount, reason,
          (data.get('attachment_path') or '').strip() or None,
          user['display_name'], now))
    waiver_id = c.lastrowid
    kind_label = '租金减免' if waiver_kind == 'rent' else '滞纳金减免'
    log_audit(conn, f'发起{kind_label}', 'waiver', waiver_id,
              f"合同{cid} 期数{period_text} 金额{'全额' if waive_amount is None else waive_amount} 原因:{reason}",
              user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '减免申请已提交，等待老板审批', 'waiver_id': waiver_id})


@app.route('/api/waivers/<int:wid>/approve', methods=['POST'])
@require_role('老板')
def approve_late_fee_waiver(wid):
    """J3 ② 老板审批减免：通过 → status='已通过' 推送财务待办；驳回 → status='已驳回'。"""
    data = request.json or {}
    user = request.current_user
    decision = (data.get('decision') or 'approve').strip()
    comment = (data.get('comment') or '').strip()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM waivers WHERE id=?", (wid,))
    w = c.fetchone()
    if not w:
        conn.close()
        return jsonify({'success': False, 'message': '减免申请不存在'}), 404
    if w['status'] != '待审批':
        conn.close()
        return jsonify({'success': False, 'message': '该申请已处理'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if decision in ('reject', '驳回'):
        if not comment:
            conn.close()
            return jsonify({'success': False, 'message': '驳回原因不能为空'}), 400
        c.execute("UPDATE waivers SET status='已驳回', boss_approved_by=?, boss_approved_at=?, revoke_reason=? WHERE id=?",
                  (user['display_name'], now, comment, wid))
        kind_label = '租金减免' if w['waiver_kind'] == 'rent' else '滞纳金减免'
        log_audit(conn, f'驳回{kind_label}', 'waiver', wid, f"原因:{comment}", user['display_name'])
        msg = '减免申请已驳回'
    else:
        c.execute("UPDATE waivers SET status='已通过', boss_approved_by=?, boss_approved_at=? WHERE id=?",
                  (user['display_name'], now, wid))
        kind_label = '租金减免' if w['waiver_kind'] == 'rent' else '滞纳金减免'
        log_audit(conn, f'通过{kind_label}', 'waiver', wid, comment or '老板审批通过，等待财务执行', user['display_name'])
        msg = '减免申请已通过，等待财务执行'
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': msg})


@app.route('/api/waivers/<int:wid>/execute', methods=['POST'])
@require_role('财务')
def execute_late_fee_waiver(wid):
    """J3 ③ 财务执行减免（status 必须='已通过'）：
    - 命中期数（target_period_list）的 late_fee_ledger 标 waived=1、waived_amount；
    - 同步 contract_fee_items(late_fee).amount_due -= waive_amount（由 _sync_late_fee_item 按 net 重算保证幂等）；
    - 写 reconciliation_allocations(allocation_type='late_fee_waiver')；
    - waivers.status='已生效'，finance_reviewed_by/at 落库。"""
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM waivers WHERE id=?", (wid,))
    w = c.fetchone()
    if not w:
        conn.close()
        return jsonify({'success': False, 'message': '减免申请不存在'}), 404
    if w['status'] != '已通过':
        conn.close()
        return jsonify({'success': False, 'message': '仅"已通过"的减免可由财务执行'}), 400

    cid = w['contract_id']
    if w['waiver_kind'] == 'rent':
        requested = w['waive_amount']
        if requested in (None, '') or parse_money(requested) <= 0:
            conn.close()
            return jsonify({'success': False, 'message': '租金减免必须填写大于0的减免金额'}), 400
        try:
            periods = parse_target_periods(w['target_period_list'])
        except (TypeError, ValueError):
            conn.close()
            return jsonify({'success': False, 'message': '目标期数格式不正确'}), 400
        params = [cid]
        period_clause = ''
        if periods is not None:
            placeholders = ','.join('?' * len(periods))
            period_clause = f' AND period IN ({placeholders})'
            params.extend(periods)
        c.execute(f"""
            SELECT id, period, due_date, amount, paid_amount
            FROM repayments
            WHERE contract_id=?
              AND period>=1
              {period_clause}
            ORDER BY period ASC, id ASC
        """, params)
        rows = c.fetchall()
        available_rows = []
        for row in rows:
            reducible = max(0.0, parse_money(row['amount']) - parse_money(row['paid_amount']))
            if reducible > 0:
                available_rows.append((row, reducible))
        available = round(sum(item[1] for item in available_rows), 2)
        waive_total = round(min(parse_money(requested), available), 2)
        if waive_total <= 0:
            conn.close()
            return jsonify({'success': False, 'message': '目标期数无可减免租金'}), 400

        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        remaining = waive_total
        for idx, (row, reducible) in enumerate(available_rows):
            if remaining <= 0:
                break
            alloc = remaining if idx == len(available_rows) - 1 else min(reducible, remaining)
            alloc = round(min(alloc, reducible), 2)
            if alloc <= 0:
                continue
            new_amount = round(parse_money(row['amount']) - alloc, 2)
            paid = parse_money(row['paid_amount'])
            new_status = repayment_status_after_allocation(row['due_date'], paid, new_amount)
            c.execute("""
                UPDATE repayments
                SET amount=?,
                    status=?,
                    waterfall_summary=COALESCE(NULLIF(waterfall_summary, ''), 'rent waiver applied')
                WHERE id=?
            """, (new_amount, new_status, row['id']))
            sync_period_shortfall_receivable(
                conn,
                row['id'],
                max(0, new_amount - paid),
                row['due_date'],
                f'租金减免#{wid}',
            )
            c.execute("""
                INSERT INTO reconciliation_allocations
                    (repayment_id, contract_id, fee_item_id, allocation_type, allocated_amount, note, created_by)
                VALUES (?, ?, NULL, 'rent_waiver', ?, ?, ?)
            """, (row['id'], cid, alloc, f"K租金减免#{wid} 第{row['period']}期", user['display_name']))
            remaining = round(remaining - alloc, 2)

        c.execute("UPDATE waivers SET status='已生效', waive_amount=?, finance_reviewed_by=?, finance_reviewed_at=? WHERE id=?",
                  (waive_total, user['display_name'], now, wid))
        log_audit(conn, '执行租金减免', 'waiver', wid,
                  f"合同{cid} 减免¥{waive_total}", user['display_name'])
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'已减免租金 ¥{waive_total}', 'waived_amount': waive_total})

    # 解析命中期数（'all' 或 JSON 数组）
    period_filter = w['target_period_list']
    periods = None
    if period_filter and period_filter not in ('all', '全部'):
        try:
            parsed = json.loads(period_filter)
            periods = [int(p) for p in parsed] if isinstance(parsed, list) else [int(parsed)]
        except (ValueError, TypeError):
            periods = None

    # 取每个 repayment 的最新累计行（cumulative_amount 为该期滞纳金毛额口径）
    c.execute("""
        SELECT lf.id, lf.repayment_id, lf.period, lf.cumulative_amount
        FROM late_fee_ledger lf
        JOIN (SELECT repayment_id, MAX(accrued_date) md FROM late_fee_ledger
              WHERE contract_id=? GROUP BY repayment_id) m
          ON m.repayment_id = lf.repayment_id AND m.md = lf.accrued_date
        WHERE lf.contract_id=?
    """, (cid, cid))
    latest_rows = [r for r in c.fetchall()
                   if periods is None or (r['period'] in periods)]
    if not latest_rows:
        conn.close()
        return jsonify({'success': False, 'message': '命中期数无已计提滞纳金，无可减免'}), 400

    # 各 repayment 已减免额 → 计算可减免余额
    matched_rids = [r['repayment_id'] for r in latest_rows]
    placeholders = ','.join('?' * len(matched_rids))
    c.execute(f"SELECT COALESCE(SUM(waived_amount),0) AS w FROM late_fee_ledger "
              f"WHERE repayment_id IN ({placeholders}) AND waived=1", matched_rids)
    already = parse_money(c.fetchone()['w'])
    gross = round(sum(parse_money(r['cumulative_amount']) for r in latest_rows), 2)
    available = round(max(0.0, gross - already), 2)
    if available <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '该滞纳金已全部减免，无可减免余额'}), 400

    requested = w['waive_amount']
    waive_total = available if requested in (None, '') else min(parse_money(requested), available)
    waive_total = round(waive_total, 2)
    if waive_total <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '减免金额必须大于 0'}), 400

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # 按期数余额比例分摊减免额到各 repayment 最新累计行
    remaining = waive_total
    per_rid_avail = []
    for r in latest_rows:
        c.execute("SELECT COALESCE(SUM(waived_amount),0) AS w FROM late_fee_ledger "
                  "WHERE repayment_id=? AND waived=1", (r['repayment_id'],))
        rid_already = parse_money(c.fetchone()['w'])
        avail = round(max(0.0, parse_money(r['cumulative_amount']) - rid_already), 2)
        if avail > 0:
            per_rid_avail.append((r, avail))
    for idx, (r, avail) in enumerate(per_rid_avail):
        if remaining <= 0:
            break
        # 最后一笔吸收四舍五入余差
        alloc = remaining if idx == len(per_rid_avail) - 1 else min(avail, remaining)
        alloc = round(min(alloc, avail), 2)
        if alloc <= 0:
            continue
        c.execute("UPDATE late_fee_ledger SET waived=1, waived_amount=COALESCE(waived_amount,0)+? WHERE id=?",
                  (alloc, r['id']))
        c.execute("""
            INSERT INTO reconciliation_allocations
                (repayment_id, contract_id, fee_item_id, allocation_type, allocated_amount, note, created_by)
            VALUES (?, ?, NULL, 'late_fee_waiver', ?, ?, ?)
        """, (r['repayment_id'], cid, alloc,
              f"J3滞纳金减免#{wid} 第{r['period']}期", user['display_name']))
        remaining = round(remaining - alloc, 2)

    # 同步 contract_fee_items(late_fee) 应收（按 net=毛额-已减免 重算，幂等）
    c.execute("""
        SELECT COALESCE(SUM(latest.cum),0) AS gross FROM (
            SELECT lf.cumulative_amount AS cum FROM late_fee_ledger lf
            JOIN (SELECT repayment_id, MAX(accrued_date) md FROM late_fee_ledger
                  WHERE contract_id=? GROUP BY repayment_id) m
              ON m.repayment_id=lf.repayment_id AND m.md=lf.accrued_date
            WHERE lf.contract_id=?
        ) latest
    """, (cid, cid))
    gross_all = parse_money(c.fetchone()['gross'])
    _sync_late_fee_item(c, cid, gross_all)

    c.execute("UPDATE waivers SET status='已生效', waive_amount=?, finance_reviewed_by=?, finance_reviewed_at=? WHERE id=?",
              (waive_total, user['display_name'], now, wid))
    log_audit(conn, '执行滞纳金减免', 'waiver', wid,
              f"合同{cid} 减免¥{waive_total}（毛额¥{gross} 已减¥{already}）", user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'已减免滞纳金 ¥{waive_total}', 'waived_amount': waive_total})


@app.route('/api/waivers/<int:wid>/revoke', methods=['POST'])
@require_role('销售', '老板')
def revoke_late_fee_waiver(wid):
    """J3 撤销：仅"待审批/已通过(未生效)"可撤销；"已生效"需走反向减免登记，此处拒绝。"""
    data = request.json or {}
    user = request.current_user
    reason = (data.get('reason') or '').strip()
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM waivers WHERE id=?", (wid,))
    w = c.fetchone()
    if not w:
        conn.close()
        return jsonify({'success': False, 'message': '减免申请不存在'}), 404
    if w['status'] == '已生效':
        conn.close()
        return jsonify({'success': False, 'message': '已生效的减免不可撤销，请通过反向减免登记补登'}), 400
    if w['status'] not in ('待审批', '已通过'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态({w["status"]})不可撤销'}), 400

    c.execute("UPDATE waivers SET status='已撤销', revoke_reason=? WHERE id=?",
              (reason or '申请方撤回', wid))
    log_audit(conn, '撤销滞纳金减免', 'waiver', wid, reason or '申请方撤回', user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '减免申请已撤销'})


@app.route('/api/contracts/<int:cid>/waivers', methods=['GET'])
@login_required
def list_contract_waivers(cid):
    """合同减免审批链路（J5 展示用）。"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM waivers WHERE contract_id=? ORDER BY id DESC", (cid,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'success': True, 'waivers': rows})


@app.route('/api/contracts/<int:cid>/approval-status', methods=['GET'])
def get_contract_approval_status(cid):
    """获取合同的审批状态（兼容历史合同审批 + 首次付款/旧支付核对）"""
    conn = get_db()
    delivery = get_approval_status(conn, 'contract_delivery', cid)
    c = conn.cursor()
    c.execute("SELECT sales_order_id FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    sale_payment_ref_id = contract['sales_order_id'] if contract and contract['sales_order_id'] else cid
    payment = get_approval_status(conn, 'sale_payment', sale_payment_ref_id)
    if not payment['steps'] and sale_payment_ref_id != cid:
        payment = get_approval_status(conn, 'sale_payment', cid)
    c.execute("SELECT id FROM contract_initial_payments WHERE contract_id=? ORDER BY id DESC LIMIT 1", (cid,))
    initial_row = c.fetchone()
    initial_payment = get_approval_status(conn, 'initial_payment', initial_row['id']) if initial_row else {'overall_status': '未开始', 'steps': []}
    conn.close()
    return jsonify({'delivery': delivery, 'payment': payment, 'initial_payment': initial_payment})


@app.route('/api/contracts/<int:cid>/delivery-files', methods=['POST'])
@require_role('车管')
def save_delivery_files(cid):
    """保存出库照片和出库单照片"""
    data = request.json or {}
    user = request.current_user
    delivery_photo_path = (data.get('delivery_photo_path') or '').strip()
    delivery_document_path = (data.get('delivery_document_path') or '').strip()
    if not delivery_photo_path and not delivery_document_path:
        return jsonify({'success': False, 'message': '请至少上传一项出库资料'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, vehicle_id, delivery_status FROM contracts WHERE id=?", (cid,))
    contract = c.fetchone()
    if not contract:
        conn.close()
        return jsonify({'success': False, 'message': '合同不存在'}), 404

    updates = []
    params = []
    if delivery_photo_path:
        updates.append("delivery_photo_path=?")
        params.append(delivery_photo_path)
    if delivery_document_path:
        updates.append("delivery_document_path=?")
        params.append(delivery_document_path)
    params.append(cid)
    c.execute(f"UPDATE contracts SET {', '.join(updates)} WHERE id=?", params)

    saved_fields = []
    if delivery_photo_path:
        saved_fields.append('出库照片')
    if delivery_document_path:
        saved_fields.append('出库单照片')
    log_audit(
        conn,
        '保存出库资料',
        'contract',
        cid,
        f'{user["display_name"]}({user["role"]}) 上传/更新 {"/".join(saved_fields)}'
    )
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '出库资料已保存'})


# ======================== 车辆出库（车管确认 - 需审批通过后）========================
@app.route('/api/vehicles/<int:vid>/deliver', methods=['POST'])
@require_role('车管')
def deliver_vehicle(vid):
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    # 验证合同是否已审批通过（delivery_status='待出库'）
    c.execute("""
        SELECT id, contract_type, delivery_status, delivery_photo_path, delivery_document_path
        FROM contracts
        WHERE vehicle_id=?
        ORDER BY id DESC
        LIMIT 1
    """, (vid,))
    ct = c.fetchone()
    if not ct:
        conn.close()
        return jsonify({'success': False, 'message': '该车辆没有关联合同，无法出库'}), 400
    if ct['delivery_status'] != '待出库':
        conn.close()
        status = ct['delivery_status'] or '未知'
        if status == '已出库':
            msg = '该车辆已出库，无需重复操作（请刷新页面查看最新状态）'
        else:
            msg = f'当前交付状态为「{status}」，尚未到出库步骤，无法出库'
        return jsonify({'success': False, 'message': msg}), 400
    if not ct['delivery_photo_path'] or not ct['delivery_document_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先上传出库照片和出库单照片，再执行出库'}), 400

    c.execute("UPDATE contracts SET delivery_status='已出库', delivery_date=? WHERE id=?",
              (datetime.now().strftime('%Y-%m-%d'), ct['id']))
    c.execute("""
        UPDATE repayments
        SET status='待还款'
        WHERE contract_id=?
          AND period>=1
          AND status='未激活'
    """, (ct['id'],))

    # 更新车辆状态：仅 租赁 / 以租代售 两种合同类型（整车销售已下线）
    # 多车支持（SKU 改造）：合同关联的所有车辆批量更新状态
    contract_type = ct['contract_type']
    status_map = {'以租代售': '以租代售', '租赁': '租赁中'}
    new_status = status_map.get(contract_type, '租赁中')
    cv_rows = c.execute("SELECT vehicle_id FROM contract_vehicles WHERE contract_id=?", (ct['id'],)).fetchall()
    if cv_rows:
        deliver_vids = [r['vehicle_id'] for r in cv_rows]
    else:
        deliver_vids = [vid]
    for dvid in deliver_vids:
        c.execute("UPDATE vehicles SET status=? WHERE id=?", (new_status, dvid))

    log_audit(conn, '车辆出库', 'vehicle', vid, f'{user["display_name"]}({user["role"]}) 确认出库 合同类型:{contract_type} 关联车辆:{len(deliver_vids)}台')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已出库'})


# ======================== 逾期催促 T+N ========================
@app.route('/api/repayments/<int:rid>/urge', methods=['POST'])
@login_required
def urge_repayment(rid):
    """完成 T+3 运营或 T+7 销售催收任务，证据是锁车的前置条件。"""
    user = request.current_user
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()

    c.execute("""SELECT r.*, c.vehicle_id, c.id as cid FROM repayments r
                 JOIN contracts c ON c.id = r.contract_id WHERE r.id=?""", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '还款记录不存在'}), 404
    if int(row['period'] or 0) <= 0 or row['status'] == '已还款':
        conn.close()
        return jsonify({'success': False, 'message': '该记录不需要催收'}), 400
    if not row['due_date']:
        conn.close()
        return jsonify({'success': False, 'message': '还款日期缺失，无法催收'}), 400

    try:
        overdue_days = max(0, (datetime.now().date() - datetime.strptime(row['due_date'], '%Y-%m-%d').date()).days)
    except ValueError:
        conn.close()
        return jsonify({'success': False, 'message': '还款日期格式错误'}), 400

    if overdue_days >= 7 and user['role'] == '销售':
        urge_type = '销售催款'
        urge_day = 7
    elif overdue_days >= 3 and user['role'] == '运营':
        urge_type = '运营催款'
        urge_day = 3
    else:
        conn.close()
        return jsonify({'success': False, 'message': f'当前逾期{overdue_days}天，不满足催款条件或角色不匹配'}), 400

    evidence_path = (data.get('evidence_path') or data.get('screenshot_path') or '').strip()
    result = (data.get('result') or '').strip() or '已联系，待跟进'
    remark = (data.get('remark') or '').strip()
    promised_repay_date = (data.get('promised_repay_date') or '').strip() or None
    if not evidence_path:
        conn.close()
        return jsonify({'success': False, 'message': '请上传催收证据后再提交'}), 400

    c.execute("""
        INSERT OR IGNORE INTO urge_records
            (repayment_id, contract_id, vehicle_id, urge_type, urge_day, status, result, remark)
        VALUES (?, ?, ?, ?, ?, '待执行', '待跟进', ?)
    """, (
        rid,
        row['cid'],
        row['vehicle_id'],
        urge_type,
        urge_day,
        f'补建逾期 T+{urge_day} 催收任务',
    ))
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE urge_records
        SET status='已完成',
            result=?,
            remark=?,
            evidence_path=?,
            promised_repay_date=?,
            operator_id=?,
            operator_name=?,
            closed_by=?,
            completed_at=?
        WHERE repayment_id=? AND urge_day=?
    """, (
        result,
        remark or None,
        evidence_path,
        promised_repay_date,
        user['id'],
        user['display_name'],
        user['display_name'],
        now,
        rid,
        urge_day,
    ))

    log_audit(conn, '催款', 'repayment', rid,
              f'{user["display_name"]}({urge_type}) 完成逾期{overdue_days}天催收，证据:{evidence_path}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'{urge_type} 已完成'})


@app.route('/api/repayments/<int:rid>/urge-records', methods=['GET'])
@login_required
def get_urge_records(rid):
    """获取某笔还款的催促记录"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM urge_records WHERE repayment_id=? ORDER BY urge_day ASC, created_at DESC", (rid,))
    records = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(records)


# ======================== H4 锁车/解锁（标记式审批流，平台不远程控车）========================
@app.route('/api/lock-requests', methods=['GET'])
@login_required
def get_lock_requests():
    """获取锁车/解锁申请列表（含 action 与车辆 lock_status）"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT lr.*, v.plate_number, v.car_type, v.lock_status, cu.name as customer_name
                 FROM lock_requests lr
                 JOIN vehicles v ON v.id = lr.vehicle_id
                 LEFT JOIN contracts c ON c.id = lr.contract_id
                 LEFT JOIN customers cu ON cu.id = c.customer_id
                 ORDER BY lr.id DESC""")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/lock-requests', methods=['POST'])
@require_role('销售')
def create_lock_request():
    """H4-1 销售发起锁车申请（diff≥7，且对应期数已有 H2 催款记录）"""
    data = request.json or {}
    user = request.current_user
    conn = get_db()
    c = conn.cursor()

    repayment_id = data.get('repayment_id')
    c.execute("""
        SELECT r.id, r.contract_id, r.due_date, r.status, c.vehicle_id
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        WHERE r.id = ?
    """, (repayment_id,))
    repayment = c.fetchone()
    if not repayment:
        conn.close()
        return jsonify({'success': False, 'message': '还款记录不存在'}), 404

    overdue_days = max(0, (datetime.now() - datetime.strptime(repayment['due_date'], '%Y-%m-%d')).days)
    if overdue_days < 7 and user['role'] != '老板':
        conn.close()
        return jsonify({'success': False, 'message': '逾期满 7 天（diff≥7）后才能发起锁车申请'}), 400

    # 前置校验：销售已完成 T+7 催收并留存证据。
    c.execute("""
        SELECT id FROM urge_records
        WHERE repayment_id=?
          AND urge_day=7
          AND status='已完成'
          AND COALESCE(evidence_path, '')!=''
        LIMIT 1
    """, (repayment_id,))
    if not c.fetchone() and user['role'] != '老板':
        conn.close()
        return jsonify({'success': False, 'message': '请先完成并留存该期 T+7 销售催收证据，再发起锁车申请'}), 400

    # 同车已存在进行中的锁车/解锁流程则拒绝
    c.execute("""
        SELECT id FROM lock_requests
        WHERE vehicle_id=? AND status IN ('待运营审核', '待老板审批', '已批准待执行')
        ORDER BY id DESC LIMIT 1
    """, (repayment['vehicle_id'],))
    if c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '该车已存在进行中的锁车/解锁流程'}), 400

    c.execute("SELECT lock_status FROM vehicles WHERE id=?", (repayment['vehicle_id'],))
    vrow = c.fetchone()
    if vrow and vrow['lock_status'] == '车辆已锁':
        conn.close()
        return jsonify({'success': False, 'message': '该车已处于「车辆已锁」状态'}), 400

    c.execute("""INSERT INTO lock_requests
                 (vehicle_id, contract_id, repayment_id, action, reason, overdue_days, status, requested_by)
                 VALUES (?, ?, ?, 'lock', ?, ?, '待运营审核', ?)""",
              (repayment['vehicle_id'], repayment['contract_id'], repayment_id,
               data.get('reason', '逾期锁车'), overdue_days, user['display_name']))
    lr_id = c.lastrowid
    c.execute("UPDATE vehicles SET lock_status='锁车流程中' WHERE id=?", (repayment['vehicle_id'],))

    log_audit(conn, '发起锁车', 'vehicle', repayment['vehicle_id'],
              f'逾期{overdue_days}天 发起人:{user["display_name"]} → 待运营审核')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '锁车申请已提交，等待运营审核', 'lock_request_id': lr_id})


def _lock_rollback_status(action):
    """驳回时 lock_status 回退：lock→未锁；unlock→车辆已锁。"""
    return '车辆已锁' if action == 'unlock' else '未锁'


@app.route('/api/lock-requests/<int:lr_id>/ops-review', methods=['POST'])
@require_role('运营')
def ops_review_lock_request(lr_id):
    """H4 运营审核（待运营审核 → 待老板审批 / 已驳回）"""
    data = request.json or {}
    user = request.current_user
    decision = data.get('decision', 'approve')
    comment = (data.get('comment') or '').strip()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM lock_requests WHERE id=?", (lr_id,))
    lr = c.fetchone()
    if not lr:
        conn.close()
        return jsonify({'success': False, 'message': '锁车/解锁申请不存在'}), 404
    if lr['status'] != '待运营审核':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为「{lr["status"]}」，无法运营审核'}), 400

    if decision in ('reject', '驳回'):
        if not comment:
            conn.close()
            return jsonify({'success': False, 'message': '驳回必须填写原因'}), 400
        c.execute("UPDATE lock_requests SET status='已驳回', ops_approved_by=?, ops_approved_at=?, reject_reason=? WHERE id=?",
                  (user['display_name'], now, comment, lr_id))
        c.execute("UPDATE vehicles SET lock_status=? WHERE id=?", (_lock_rollback_status(lr['action']), lr['vehicle_id']))
        log_audit(conn, '运营驳回锁车' if lr['action'] == 'lock' else '运营驳回解锁',
                  'vehicle', lr['vehicle_id'], f'{user["display_name"]} 驳回 原因:{comment}')
        msg = '已驳回'
    else:
        c.execute("UPDATE lock_requests SET status='待老板审批', ops_approved_by=?, ops_approved_at=? WHERE id=?",
                  (user['display_name'], now, lr_id))
        log_audit(conn, '运营通过锁车' if lr['action'] == 'lock' else '运营通过解锁',
                  'vehicle', lr['vehicle_id'], f'{user["display_name"]} 通过 → 待老板审批')
        msg = '运营审核通过，等待老板审批'
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': msg})


@app.route('/api/lock-requests/<int:lr_id>/boss-approve', methods=['POST'])
@require_role('老板')
def boss_approve_lock_request(lr_id):
    """H4 老板审批（待老板审批 → 已批准待执行 / 已驳回）"""
    data = request.json or {}
    user = request.current_user
    decision = data.get('decision', 'approve')
    comment = (data.get('comment') or '').strip()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM lock_requests WHERE id=?", (lr_id,))
    lr = c.fetchone()
    if not lr:
        conn.close()
        return jsonify({'success': False, 'message': '锁车/解锁申请不存在'}), 404
    if lr['status'] != '待老板审批':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为「{lr["status"]}」，无法老板审批'}), 400

    if decision in ('reject', '驳回'):
        if not comment:
            conn.close()
            return jsonify({'success': False, 'message': '驳回必须填写原因'}), 400
        c.execute("UPDATE lock_requests SET status='已驳回', boss_approved_by=?, boss_approved_at=?, reject_reason=? WHERE id=?",
                  (user['display_name'], now, comment, lr_id))
        c.execute("UPDATE vehicles SET lock_status=? WHERE id=?", (_lock_rollback_status(lr['action']), lr['vehicle_id']))
        log_audit(conn, '老板驳回锁车' if lr['action'] == 'lock' else '老板驳回解锁',
                  'vehicle', lr['vehicle_id'], f'{user["display_name"]} 驳回 原因:{comment}')
        msg = '已驳回'
    else:
        c.execute("UPDATE lock_requests SET status='已批准待执行', boss_approved_by=?, boss_approved_at=? WHERE id=?",
                  (user['display_name'], now, lr_id))
        action_label = '锁车' if lr['action'] == 'lock' else '解锁'
        log_audit(conn, f'老板通过{action_label}', 'vehicle', lr['vehicle_id'],
                  f'{user["display_name"]} 通过 → 已批准待执行，等待销售线下{action_label}')
        msg = f'老板审批通过，等待销售到第三方控车系统{action_label}后回平台确认'
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': msg})


@app.route('/api/lock-requests/<int:lr_id>/confirm', methods=['POST'])
@require_role('销售')
def confirm_lock_action(lr_id):
    """H4 销售线下操作完成后回平台确认（已批准待执行 → 已完成）。
    lock → lock_status='车辆已锁' 记 locked_by/at；unlock → lock_status='未锁' 记 unlocked_by/at。"""
    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM lock_requests WHERE id=?", (lr_id,))
    lr = c.fetchone()
    if not lr:
        conn.close()
        return jsonify({'success': False, 'message': '锁车/解锁申请不存在'}), 404
    if lr['status'] != '已批准待执行':
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为「{lr["status"]}」，无法确认执行'}), 400

    if lr['action'] == 'lock':
        c.execute("UPDATE lock_requests SET status='已完成', locked_by=?, locked_at=? WHERE id=?",
                  (user['display_name'], now, lr_id))
        c.execute("UPDATE vehicles SET lock_status='车辆已锁' WHERE id=?", (lr['vehicle_id'],))
        log_audit(conn, '确认锁车', 'vehicle', lr['vehicle_id'], f'{user["display_name"]} 第三方系统锁车成功 → 车辆已锁')
        msg = '车辆已锁'
    else:
        c.execute("UPDATE lock_requests SET status='已完成', unlocked_by=?, unlocked_at=? WHERE id=?",
                  (user['display_name'], now, lr_id))
        c.execute("UPDATE vehicles SET lock_status='未锁' WHERE id=?", (lr['vehicle_id'],))
        log_audit(conn, '确认解锁', 'vehicle', lr['vehicle_id'], f'{user["display_name"]} 第三方系统解锁成功 → 未锁')
        msg = '车辆已解锁'
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': msg})


@app.route('/api/unlock-requests', methods=['POST'])
@require_role('销售')
def create_unlock_request():
    """H4-2 销售发起解锁申请（车辆需处于「车辆已锁」）"""
    data = request.json or {}
    user = request.current_user
    conn = get_db()
    c = conn.cursor()

    vehicle_id = data.get('vehicle_id')
    c.execute("SELECT id, lock_status FROM vehicles WHERE id=?", (vehicle_id,))
    v = c.fetchone()
    if not v:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    if v['lock_status'] != '车辆已锁':
        conn.close()
        return jsonify({'success': False, 'message': f'车辆当前为「{v["lock_status"]}」，仅「车辆已锁」可发起解锁'}), 400

    c.execute("""
        SELECT id FROM lock_requests
        WHERE vehicle_id=? AND status IN ('待运营审核', '待老板审批', '已批准待执行')
        ORDER BY id DESC LIMIT 1
    """, (vehicle_id,))
    if c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '该车已存在进行中的锁车/解锁流程'}), 400

    c.execute("SELECT id FROM contracts WHERE vehicle_id=? ORDER BY id DESC LIMIT 1", (vehicle_id,))
    crow = c.fetchone()
    contract_id = crow['id'] if crow else None

    c.execute("""INSERT INTO lock_requests
                 (vehicle_id, contract_id, action, reason, status, requested_by)
                 VALUES (?, ?, 'unlock', ?, '待运营审核', ?)""",
              (vehicle_id, contract_id, data.get('reason', '客户已结清/还款到位'), user['display_name']))
    lr_id = c.lastrowid
    c.execute("UPDATE vehicles SET lock_status='开锁流程中' WHERE id=?", (vehicle_id,))

    log_audit(conn, '发起解锁', 'vehicle', vehicle_id, f'发起人:{user["display_name"]} → 待运营审核')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '解锁申请已提交，等待运营审核', 'lock_request_id': lr_id})


# ======================== H1 逐日催收统一入口 ========================
@app.route('/api/jobs/daily-collect', methods=['POST'])
@login_required
def trigger_daily_collect():
    """H1 统一入口：手动/调度触发逐日催收。
    (job='daily-collect', run_date) 当天幂等；force=true 可强制重跑当天。"""
    force = bool((request.get_json(silent=True) or {}).get('force', False))
    result = run_daily_collect(force=force)
    return jsonify({'success': True, **result})


def _daily_collect_scheduler():
    """单机零依赖调度：每日 00:30 触发 run_daily_collect。
    run_daily_collect 按 (job,run_date) 幂等，漏跑由惰性兜底/下次循环补跑。"""
    while True:
        now = datetime.now()
        nxt = now.replace(hour=0, minute=30, second=0, microsecond=0)
        if nxt <= now:
            nxt += timedelta(days=1)
        _time.sleep(max(1, (nxt - now).total_seconds()))
        try:
            run_daily_collect(force=False)
        except Exception as e:
            print(f'[daily-collect] scheduler error: {e}')


def start_scheduler():
    global _SCHEDULER_STARTED
    if _SCHEDULER_STARTED:
        return
    _SCHEDULER_STARTED = True
    t = threading.Thread(target=_daily_collect_scheduler, daemon=True)
    t.start()


@app.before_request
def ensure_scheduler_started():
    if app.config.get('TESTING'):
        return
    start_scheduler()


# ======================== 启动 ========================
if __name__ == '__main__':
    init_db()
    migration_conn = get_db()
    try:
        migrate_legacy_pending_approvals(migration_conn)
        migration_conn.commit()
    finally:
        migration_conn.close()
    seed_data()
    # SKU 改造：字典种子数据（静态配置 + 存量车辆维度值灌入，幂等）
    seed_conn = get_db()
    try:
        seed_data_dictionaries(seed_conn)
        sku_created = backfill_skus(seed_conn)
        if sku_created:
            print(f"SKU backfilled: {sku_created} entries")
    finally:
        seed_conn.close()
    # 启动补跑一次（漏跑兜底），并启动每日 00:30 调度线程
    try:
        run_daily_collect(force=False)
    except Exception as e:
        print(f'[daily-collect] startup catch-up error: {e}')
    start_scheduler()
    app.run(port=49165, debug=True, use_reloader=False)
