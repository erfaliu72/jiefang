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
import zipfile
from xml.sax.saxutils import escape as xml_escape

app = Flask(__name__)
app.secret_key = 'jinjuyuan-secret-2024'
CORS(app, supports_credentials=True)

# 上传文件目录
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)
CONTRACT_TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'contract_templates', 'raw')
CONTRACT_OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'contract_templates', 'generated')
os.makedirs(CONTRACT_OUTPUT_DIR, exist_ok=True)

CONTRACT_TEMPLATE_FILES = {
    '销售': '【金聚源】车辆买卖合同（新车）-陈律师-20250905.docx',
    '租赁': '2026年  金聚源 车辆租赁合同.docx',
    '以租代售': '金聚源 （以租代购）-陈律师-20250905.docx',
}

_DOCX_READY = None
_PDF_READY = None
_SCHEDULER_STARTED = False

# ======================== PRD 角色权限矩阵 ========================
# 每个角色可访问的页面 — 车辆列表全员可见
ROLE_PAGES = {
    '老板': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return', 'profit', 'settings'],
    '运营': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return'],
    '财务': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'profit', 'return'],
    '车管': ['dashboard', 'assets', 'approvals', 'return'],
    '销售': ['dashboard', 'orders', 'assets', 'approvals', 'risk', 'return'],
}

# 每个角色可执行的操作
ROLE_ACTIONS = {
    '老板': ['*'],
    '运营': ['view_contracts', 'view_overdue', 'lock_vehicle', 'execute_lock', 'confirm_repayment', 'initiate_return', 'view_orders'],
    '财务': ['view_contracts', 'confirm_repayment', 'confirm_factory', 'view_bills', 'view_profit', 'upload_receipt', 'collect_payment', 'verify_return', 'upload_initial_receipt', 'activate_order'],
    '车管': ['add_vehicle', 'update_vehicle', 'activate_vehicle', 'return_inspect', 'deliver_vehicle', 'return_stock'],
    '销售': ['create_contract', 'view_contracts', 'upload_screenshot', 'view_overdue', 'initiate_return', 'request_lock', 'initiate_initial_payment', 'create_order'],
}

ROLE_HIDDEN_FIELDS = {
    '销售': {
        'vehicles': ['purchase_price', 'tax_rate', 'estimated_residual_value', 'guidance_price'],
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
        'vehicles': ['purchase_price', 'tax_rate', 'guidance_price'],
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
        ],
        'contracts': [
            'loan_amount', 'monthly_payment', 'factory_guarantee_deposit', 'paid_principal',
            'loan_balance', 'collected_deposit', 'collected_rent', 'deposit', 'down_payment',
            'expected_profit_floor', 'expected_profit_ceiling', 'snapshot_guidance_price',
            'snapshot_invoice_price',
        ],
        'sales_orders': ['snapshot_guidance_price', 'snapshot_lease_installment_price', 'snapshot_sale_total_price'],
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

# ======================== 通用审批流程配置 ========================
APPROVAL_CONFIGS = {
    'price_exception': [
        {'step': 1, 'role': '老板', 'label': '低于指导价审批'},
    ],
    'sale_payment': [
        {'step': 1, 'role': '财务', 'label': '财务确认报单'},
    ],
    'initial_payment': [
        {'step': 1, 'role': '财务', 'label': '财务审核收款'},
    ],
    'return_stock': [
        {'step': 1, 'role': '老板', 'label': '领导审批'},
    ],
}


def normalize_approval_step_label(ref_type, step_order, required_role, step_label):
    """兼容历史审批流文案，确保页面展示使用当前标准名称。"""
    if ref_type == 'initial_payment' and step_order == 1 and required_role == '财务':
        return '财务审核收款'
    if ref_type == 'return_stock' and step_order == 1 and required_role == '运营':
        return '运营查车辆数据'
    if ref_type == 'return_stock' and required_role == '老板':
        return '领导审批'
    if ref_type == 'price_exception':
        return '低于指导价审批'
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
    conn.execute("UPDATE sales_orders SET order_status='已激活' WHERE order_status='已转合同'")
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


def current_guidance_price(conn, vehicle_id):
    row = conn.execute("SELECT guidance_price FROM vehicles WHERE id=?", (vehicle_id,)).fetchone()
    return float(row['guidance_price'] or 0) if row else 0.0


def normalize_sales_mode(value):
    mode = (value or '').strip()
    if mode in ('以租代购', '以租代售'):
        return '以租代售'
    if mode in ('整车销售', '卖车', '销售'):
        return '销售'
    return '租赁'


def contract_type_from_sales_mode(value):
    return normalize_sales_mode(value)


def resolve_guidance_prices_for_vehicle(conn, vehicle):
    """返回 6.2 车型双指导价；单车旧指导价只作为 legacy 信息，不参与双价判定。"""
    car_type = (vehicle['car_type'] if vehicle and 'car_type' in vehicle.keys() else '') or ''
    legacy_price = parse_money(vehicle['guidance_price'] if vehicle and 'guidance_price' in vehicle.keys() else 0)
    result = {
        'lease_installment_price': 0.0,
        'sale_total_price': 0.0,
        'legacy_guidance_price': legacy_price,
        'source': '单车指导价' if legacy_price > 0 else '',
    }
    if car_type:
        row = conn.execute("""
            SELECT guidance_price, lease_installment_price, sale_total_price
            FROM model_guidance_prices
            WHERE car_type=?
        """, (car_type,)).fetchone()
        if row:
            lease_price = parse_money(row['lease_installment_price'])
            sale_price = parse_money(row['sale_total_price'])
            legacy_model_price = parse_money(row['guidance_price'])
            result.update({
                'lease_installment_price': lease_price,
                'sale_total_price': sale_price,
                'legacy_guidance_price': legacy_model_price or legacy_price,
                'source': '车型双指导价' if (lease_price or sale_price) else '车型指导价',
            })
    return result


def has_model_dual_guidance(conn, car_type):
    if not car_type:
        return False
    row = conn.execute("""
        SELECT lease_installment_price, sale_total_price
        FROM model_guidance_prices
        WHERE car_type=?
    """, (car_type,)).fetchone()
    return bool(row and parse_money(row['lease_installment_price']) > 0 and parse_money(row['sale_total_price']) > 0)


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
    """历史调用兼容：返回以租代售整车价优先的单值。"""
    prices = resolve_guidance_prices_for_vehicle(conn, vehicle)
    price = prices['sale_total_price'] or prices['lease_installment_price'] or prices['legacy_guidance_price']
    return price, prices['source'] or '指导价'


def unresolved_guidance_vehicle_count(conn):
    row = conn.execute("""
        SELECT COUNT(*) as cnt
        FROM vehicles v
        LEFT JOIN model_guidance_prices mgp ON mgp.car_type = v.car_type
        WHERE COALESCE(v.guidance_price, 0) <= 0
          AND (
              COALESCE(mgp.lease_installment_price, 0) <= 0
              OR COALESCE(mgp.sale_total_price, 0) <= 0
          )
          AND COALESCE(mgp.guidance_price, 0) <= 0
          AND COALESCE(v.status, '') IN ('在库', '报单锁定中')
    """).fetchone()
    return row['cnt'] if row else 0


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
    if contract_type == '销售':
        return float(value('total_price') or value('down_payment') or 0)
    if contract_type == '以租代售':
        return float(value('down_payment') or 0)
    return float(value('deposit') or 0) + float(value('rent') or 0)


def initial_payment_label(contract):
    if not contract:
        return '首付款审核'
    if contract['contract_type'] == '销售':
        return '卖车付款审核'
    if contract['contract_type'] == '以租代售':
        return '首付款审核'
    return '押金及首次支付审核'


def parse_money(value, default=0):
    try:
        if value in (None, ''):
            return float(default)
        if isinstance(value, str):
            value = value.replace(',', '').strip()
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def normalize_date(value):
    if not value:
        return ''
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value)[:10]


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


def repayment_status_after_allocation(due_date, paid_amount, expected_amount):
    if paid_amount >= expected_amount:
        return '已还款'
    if paid_amount > 0:
        return '部分核销'
    if due_date and due_date < datetime.now().strftime('%Y-%m-%d'):
        diff = (datetime.now().date() - datetime.strptime(due_date, '%Y-%m-%d').date()).days
        return f'逾期{diff}日' if diff >= 1 else '逾期'
    return '待还款'


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
    """核销客户回款，并按 5.27 调整客户还款计划表。

    费用项仍按 T+7 顺序优先分配；进入租金分期后：
    - 少于本期标准租金：差额顺延到下一期，本期关闭；
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

    next_period_adjusted = None
    if rent_allocated < rent_outstanding:
        shortfall = round(rent_outstanding - rent_allocated, 2)
        next_row = find_next_adjustable_repayment(c, contract_id, repayment['period'])
        if next_row:
            c.execute("UPDATE repayments SET amount=COALESCE(amount, 0)+? WHERE id=?", (shortfall, next_row['id']))
            next_period_adjusted = {
                'repayment_id': next_row['id'],
                'added_amount': shortfall,
            }
            allocation_lines.append(f"shortfall carried to next ¥{shortfall}")

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
            new_status = '预抵' if new_paid >= parse_money(row['amount']) else '部分核销'
            c.execute("""
                UPDATE repayments
                SET paid_amount=?,
                    verified_amount=COALESCE(verified_amount, 0)+?,
                    status=?,
                    paid_at=CASE WHEN ?='预抵' THEN ? ELSE paid_at END,
                    bank_serial=COALESCE(NULLIF(bank_serial, ''), ?),
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
                new_status,
                datetime.now().strftime('%Y-%m-%d'),
                repayment['bank_serial'] if 'bank_serial' in repayment.keys() else '',
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
    if paid_amount < expected_amount and next_period_adjusted:
        status = '已还款'
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

    return {
        'status': status,
        'paid_amount': paid_amount,
        'verified_amount': round(parse_money(repayment['verified_amount']) + received_amount, 2),
        'rent_allocated': rent_allocated,
        'remaining_unallocated': remaining,
        'next_period_adjusted': next_period_adjusted,
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
    try:
        from pypdf import PdfReader
    except Exception as e:
        raise ValueError(f'缺少 PDF 解析依赖 pypdf: {e}')

    reader = PdfReader(local_path)
    text = '\n'.join((page.extract_text() or '') for page in reader.pages)
    if not text.strip():
        raise ValueError('PDF 未提取到文本，请确认不是扫描件')

    metadata = {}
    patterns = {
        'lease_contract_no': r'租赁合同号\s+([A-Za-z0-9_-]+)',
        'lessee_name': r'承租人姓名\s+([^\s]+)',
        'principal_amount': r'计息本金\(元\)\s+([\d,]+(?:\.\d+)?)',
        'lease_months': r'租赁期限\(月\)\s+(\d+)',
        'plate_number': r'主车车牌号\s+([^\s]+)',
        'expected_repayment_day': r'预计每期还租日\s+(\d+)',
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            metadata[key] = match.group(1)

    rows = []
    row_pattern = re.compile(
        r'^\s*(\d+)\s+'
        r'(\d{4}-\d{1,2}-\d{1,2})\s+'
        r'([\d.]+)\s+'
        r'([\d,]+(?:\.\d+)?)\s+'
        r'([\d,]+(?:\.\d+)?)\s+'
        r'([\d,]+(?:\.\d+)?)\s+'
        r'([\d,]+(?:\.\d+)?)\s+'
        r'([\d,]+(?:\.\d+)?)\s+'
        r'(已归还|未归还)\s*$'
    )
    for line in text.splitlines():
        compact = ' '.join(str(line).strip().split())
        match = row_pattern.match(compact)
        if not match:
            continue
        period, due_date, rate, principal, interest, subsidy, penalty, amount, status = match.groups()
        rows.append({
            'period': int(period),
            'due_date': normalize_date(due_date),
            'amount': parse_money(amount, 0),
            'principal': parse_money(principal, 0),
            'interest': parse_money(interest, 0),
            'subsidy': parse_money(subsidy, 0),
            'penalty': parse_money(penalty, 0),
            'rate': parse_money(rate, 0),
            'source_status': status,
        })

    if not rows:
        raise ValueError('PDF 中未解析到还租计划明细')

    metadata['source_format'] = 'factory_plan_pdf'
    metadata['row_count'] = len(rows)
    metadata['total_amount'] = round(sum(parse_money(row['amount']) for row in rows), 2)
    metadata['paid_amount'] = round(sum(parse_money(row['amount']) for row in rows if row.get('source_status') == '已归还'), 2)
    metadata['unpaid_amount'] = round(metadata['total_amount'] - metadata['paid_amount'], 2)
    return rows, metadata


def parse_factory_plan_file(local_path):
    lower_path = local_path.lower()
    if lower_path.endswith('.xlsx'):
        return parse_factory_plan_sheet(local_path), {'source_format': 'xlsx'}
    if lower_path.endswith('.pdf'):
        return parse_factory_plan_pdf(local_path)
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
    ceiling = parse_money(contract['expected_profit_ceiling'], 999999999)

    customer_due_dates = [row['due_date'] for row in customer_rows if row.get('due_date')]
    factory_due_dates = [row['due_date'] for row in factory_rows if row.get('due_date')]
    status = '已通过' if factory_rows and floor <= spread_total <= ceiling else '差异待处理'
    if not factory_rows:
        status = '未上传'

    summary = {
        'status': status,
        'customer_total': customer_total,
        'factory_total': factory_total,
        'spread_total': spread_total,
        'expected_profit_floor': floor,
        'expected_profit_ceiling': ceiling,
        'spread_range_passed': bool(factory_rows and floor <= spread_total <= ceiling),
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


def sales_order_plan_activation_blocker(conn, order_id):
    """If a sales order already has a linked installment contract, E2 must respect F2/F3."""
    c = conn.cursor()
    c.execute("""
        SELECT id, contract_type, customer_plan_match_status
        FROM contracts
        WHERE sales_order_id=?
        ORDER BY id DESC
        LIMIT 1
    """, (order_id,))
    contract = c.fetchone()
    if not contract or contract['contract_type'] == '销售':
        return None

    c.execute("SELECT COUNT(*) AS cnt FROM repayments WHERE contract_id=? AND period>=1", (contract['id'],))
    customer_count = c.fetchone()['cnt']
    c.execute("SELECT COUNT(*) AS cnt FROM factory_repayments WHERE contract_id=?", (contract['id'],))
    factory_count = c.fetchone()['cnt']
    if customer_count <= 0:
        return '客户还款计划未生成，不能财务确认报单'
    if factory_count <= 0:
        return '厂家还款计划未上传，不能财务确认报单'

    status = contract['customer_plan_match_status'] or '未比对'
    if status not in ('已通过', '差异已确认'):
        return f'还款计划比对状态为{status}，请先完成 F3 比对或确认差异'
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
    digits = ''.join(ch for ch in text if ch.isdigit())
    if digits:
        return max(1, int(digits))
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
    if contract_type == '销售':
        return None

    data = overrides or {}
    vehicle_id = order['vehicle_id']
    customer_name = data.get('customer_name') or order['customer_name']
    customer_phone = data.get('customer_phone') or order['customer_phone']
    c.execute("SELECT car_type, guidance_price, invoice_price FROM vehicles WHERE id=?", (vehicle_id,))
    vrow = c.fetchone()
    snap_guidance = resolve_guidance_price_for_vehicle(conn, vrow)[0] if vrow else 0
    snap_invoice = vrow['invoice_price'] if vrow else 0

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
    loan_periods = parse_period_count(data.get('loan_periods') or order['lease_term'], 12)
    repayment_day = int(data.get('repayment_day') or 1)
    rent = parse_money(data.get('rent'), parse_money(order['vehicle_rent_amount']) or parse_money(order['sale_total_price']))
    deposit = parse_money(data.get('deposit'), parse_money(order['deposit_amount']) if contract_type == '租赁' else 0)
    down_payment = parse_money(data.get('down_payment'), parse_money(order['car_purchase_amount']) if contract_type == '以租代售' else 0)
    monthly_payment = parse_money(data.get('monthly_payment'), 0)
    factory_periods = int(data.get('factory_periods') or loan_periods)
    factory_repayment_months = int(data.get('factory_repayment_months') or factory_periods)
    customer_loan_amount = parse_money(data.get('customer_loan_amount'), 0)
    loan_amount = parse_money(data.get('loan_amount'), 0)
    total_price = parse_money(data.get('total_price'), parse_money(order['sale_total_price']))
    company = data.get('company') or order['receiving_company'] or ''
    yard = data.get('yard') or ''
    business_mode = data.get('business_mode') or order['sales_mode'] or contract_type
    rental_method = data.get('rental_method') or ('经营租赁' if contract_type == '租赁' else contract_type)
    expected_profit_floor = parse_money(data.get('expected_profit_floor'), 0)
    expected_profit_ceiling = parse_money(data.get('expected_profit_ceiling'), 999999999)
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
        c.execute("""
            UPDATE contracts
            SET vehicle_id=?, customer_id=?, contract_type=?, business_mode=?, rental_method=?, repayment_day=?,
                start_date=?, end_date=?, total_price=?, customer_loan_amount=?, loan_amount=?, monthly_payment=?,
                rent=?, loan_periods=?, company=?, yard=?, lease_bank_name=?, lease_bank_card_no=?,
                factory_guarantee_deposit=?, factory_repayment_months=?, factory_periods=?, deposit=?, down_payment=?,
                snapshot_guidance_price=?, snapshot_invoice_price=?,
                expected_profit_floor=?, expected_profit_ceiling=?,
                customer_plan_match_status='未上传',
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
            factory_plan_match_status=CASE WHEN ? THEN '未上传' ELSE factory_plan_match_status END,
            plan_compare_summary=NULL,
            contract_id=?
        WHERE id=?
    """, (1 if reset_factory else 0, contract_id, order_id))
    return contract_id


def safe_filename(text):
    cleaned = ''.join(ch if ch.isalnum() or ch in ('-', '_') else '_' for ch in str(text))
    while '__' in cleaned:
        cleaned = cleaned.replace('__', '_')
    return cleaned.strip('_') or 'contract'


def ensure_docx_support():
    global _DOCX_READY
    if _DOCX_READY is not None:
        return _DOCX_READY
    try:
        from docx import Document as _Document
        globals()['Document'] = _Document
        _DOCX_READY = True
    except Exception:
        _DOCX_READY = False
    return _DOCX_READY


def ensure_pdf_support():
    global _PDF_READY
    if _PDF_READY is not None:
        return _PDF_READY
    try:
        from reportlab.lib.pagesizes import A4 as _A4
        from reportlab.lib.utils import simpleSplit as _simpleSplit
        from reportlab.pdfbase import pdfmetrics as _pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont as _UnicodeCIDFont
        from reportlab.pdfgen import canvas as _canvas
        _pdfmetrics.registerFont(_UnicodeCIDFont('STSong-Light'))
        globals()['A4'] = _A4
        globals()['simpleSplit'] = _simpleSplit
        globals()['pdfmetrics'] = _pdfmetrics
        globals()['canvas'] = _canvas
        _PDF_READY = True
    except Exception:
        _PDF_READY = False
    return _PDF_READY


def format_currency(value):
    return f"{parse_money(value):,.2f}"


def format_date_parts(date_str):
    if not date_str:
        return ('', '', '')
    dt = datetime.strptime(date_str, '%Y-%m-%d')
    return (str(dt.year), str(dt.month), str(dt.day))


def contract_template_path(contract_type):
    filename = CONTRACT_TEMPLATE_FILES.get(contract_type)
    if not filename:
        raise ValueError(f'暂不支持合同类型: {contract_type}')
    path = os.path.join(CONTRACT_TEMPLATE_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f'合同模板不存在: {path}')
    return path


def fetch_contract_detail(contract_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT c.*, v.vin, v.plate_number, v.car_type, v.company as vehicle_company, v.engine_number,
               v.insurance_expiry_date, v.annual_review_date,
               cu.name as customer_name, cu.phone as customer_phone, cu.id_card, cu.address
        FROM contracts c
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE c.id=?
    """, (contract_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def fetch_contract_with_bank_info(contract_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT c.*, v.vin, v.plate_number, v.car_type, v.company as vehicle_company,
               cu.name as customer_name, cu.phone as customer_phone, cu.id_card, cu.address
        FROM contracts c
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        WHERE c.id=?
    """, (contract_id,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def build_contract_context(contract):
    contract_type = contract['contract_type']
    contract_no = f"JGY-{contract_type}-{contract['id']:06d}"
    start_date = contract.get('start_date') or datetime.now().strftime('%Y-%m-%d')
    end_date = contract.get('end_date') or start_date
    start_y, start_m, start_d = format_date_parts(start_date)
    end_y, end_m, end_d = format_date_parts(end_date)
    total_days = ''
    if start_date and end_date:
        total_days = str((datetime.strptime(end_date, '%Y-%m-%d') - datetime.strptime(start_date, '%Y-%m-%d')).days or 0)

    return {
        'contract_no': contract_no,
        'customer_name': contract.get('customer_name') or '',
        'customer_phone': contract.get('customer_phone') or '',
        'customer_id_card': contract.get('id_card') or '',
        'customer_address': contract.get('address') or '',
        'guarantor_name': ' / ',
        'guarantor_id_card': ' / ',
        'brand_name': '解放',
        'car_type': contract.get('car_type') or '',
        'plate_number': contract.get('plate_number') or '',
        'vin': contract.get('vin') or '',
        'company': contract.get('company') or contract.get('vehicle_company') or '陕西金聚源汽车服务有限公司',
        'yard': contract.get('yard') or '陕西金聚源',
        'lease_bank_name': contract.get('lease_bank_name') or '',
        'lease_bank_card_no': contract.get('lease_bank_card_no') or '',
        'start_date': start_date,
        'end_date': end_date,
        'start_year': start_y,
        'start_month': start_m,
        'start_day': start_d,
        'end_year': end_y,
        'end_month': end_m,
        'end_day': end_d,
        'loan_periods': str(contract.get('loan_periods') or ''),
        'loan_period_years': f"{round((contract.get('loan_periods') or 0) / 12, 2):g}" if contract.get('loan_periods') else '',
        'total_days': total_days,
        'total_price': format_currency(contract.get('total_price')),
        'rent': format_currency(contract.get('rent')),
        'deposit': format_currency(contract.get('deposit')),
        'down_payment': format_currency(contract.get('down_payment')),
        'loan_amount': format_currency(contract.get('loan_amount')),
        'repayment_day': str(contract.get('repayment_day') or ''),
        'monthly_payment': format_currency(contract.get('monthly_payment')),
        'factory_guarantee_deposit': format_currency(contract.get('factory_guarantee_deposit')),
        'contract_type': contract_type,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }


def replace_paragraph_text(paragraph, replacements):
    text = paragraph.text
    if not text:
        return
    new_text = text
    for old, new in replacements:
        if old in new_text:
            new_text = new_text.replace(old, new)
    if new_text != text:
        if paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ''
        else:
            paragraph.add_run(new_text)


def replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_paragraph_text(paragraph, replacements)


def build_replacements(contract):
    ctx = build_contract_context(contract)
    if contract['contract_type'] == '销售':
        return [
            ('合同编号：【  】', f"合同编号：【{ctx['contract_no']}】"),
            ('乙方（买受人）：                  身份证号：', f"乙方（买受人）：{ctx['customer_name']}    身份证号：{ctx['customer_id_card']}"),
            ('电话                  ', f"电话 {ctx['customer_phone']}"),
            ('车辆品牌：  ，车型： ，车架号：', f"车辆品牌：{ctx['brand_name']}，车型：{ctx['car_type']}，车架号：{ctx['vin']}"),
            ('2.1车辆含税金额为          元，税率为13%。', f"2.1车辆含税金额为 {ctx['total_price']} 元，税率为13%。"),
            ('按揭贷款方式付款：乙方应当于签订之日起     日内向甲方支付车辆首付款     元，余款     元乙方向相关贷款机构申请贷款支付。',
             f"按揭贷款方式付款：乙方应当于签订之日起 7 日内向甲方支付车辆首付款 {ctx['down_payment']} 元，余款 {ctx['loan_amount']} 元乙方向相关贷款机构申请贷款支付。"),
            ('分期付款：乙方应当于     年     月     日前分     期支付该车辆的全部价款，其中乙方应当于签订之日起     日内向甲方支付车辆首付款     元，剩余款项于每月     日前向甲方支付。',
             f"分期付款：乙方应当于 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日前分 {ctx['loan_periods']} 期支付该车辆的全部价款，其中乙方应当于签订之日起 7 日内向甲方支付车辆首付款 {ctx['down_payment']} 元，剩余款项于每月 {ctx['repayment_day']} 日前向甲方支付。"),
            ('4.1甲方应当于乙方支付全部车款后     日内向乙方交付车辆。', '4.1甲方应当于乙方支付全部车款后 15 日内向乙方交付车辆。'),
            ('（本页为合同编号为【         】车辆买卖合同的签章页）', f'（本页为合同编号为【{ctx["contract_no"]}】车辆买卖合同的签章页）'),
        ]
    if contract['contract_type'] == '租赁':
        return [
            ('合同编号：【2026032902】', f"合同编号：【{ctx['contract_no']}】"),
            ('乙方（承租人）：            身份证号：               电话', f"乙方（承租人）：{ctx['customer_name']}    身份证号：{ctx['customer_id_card']}    电话 {ctx['customer_phone']}"),
            ('丙方（保证人）：     /       身份证号：     /          电话    /', '丙方（保证人）： /       身份证号： /          电话 /'),
            ('车型：       ， 车牌号       ，车架号：                          。', f"车型：{ctx['car_type']}， 车牌号 {ctx['plate_number']}，车架号：{ctx['vin']}。"),
            ('2.1起算日：     年    月    日至   年   月   日（实际以租赁车辆实际交付乙方之日起算），共    年（    天）',
             f"2.1起算日：{ctx['start_year']} 年 {ctx['start_month']} 月 {ctx['start_day']} 日至 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日（实际以租赁车辆实际交付乙方之日起算），共 {ctx['loan_period_years']} 年（{ctx['total_days']} 天）"),
            ('3.1.1本合同租赁车辆的租金标准为      元/月。', f"3.1.1本合同租赁车辆的租金标准为 {ctx['rent']} 元/月。"),
            ('3.1.2乙方应当于甲方交付车辆前7日内向甲方交纳首期租金，剩余期限的租金乙方应当于每月【 】日前向甲方支付下一个支付周期的租金。',
             f"3.1.2乙方应当于甲方交付车辆前7日内向甲方交纳首期租金，剩余期限的租金乙方应当于每月【{ctx['repayment_day']}】日前向甲方支付下一个支付周期的租金。"),
            ('3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【    】元，作为乙方履行本合同的保证。',
             f"3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【{ctx['deposit']}】元，作为乙方履行本合同的保证。"),
        ]
    return [
        ('合同编号：【       】', f"合同编号：【{ctx['contract_no']}】"),
        ('乙方（承租人）：            ，身份证号：', f"乙方（承租人）：{ctx['customer_name']}，身份证号：{ctx['customer_id_card']}"),
        ('丙方（保证人）：            ，身份证号：', '丙方（保证人）： / ，身份证号： / '),
        ('车型：          ， 车牌号         ，车架号：               。', f"车型：{ctx['car_type']}， 车牌号 {ctx['plate_number']}，车架号：{ctx['vin']}。"),
        ('起算日：      年     月     日至      年    月    日（实际以租赁车辆实际交付乙方之日起算），共/年（    个月）',
         f"起算日：{ctx['start_year']} 年 {ctx['start_month']} 月 {ctx['start_day']} 日至 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日（实际以租赁车辆实际交付乙方之日起算），共 {ctx['loan_period_years']} 年（{ctx['loan_periods']} 个月）"),
        ('3.1.1本合同租赁车辆的租金标准为租金标准为每月     元/月，租期       个月，合计       元。',
         f"3.1.1本合同租赁车辆的租金标准为每月 {ctx['rent']} 元/月，租期 {ctx['loan_periods']} 个月，合计 {ctx['total_price']} 元。"),
        ('3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【 】元，作为乙方履行本合同的保证。',
         f"3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【{ctx['deposit']}】元，作为乙方履行本合同的保证。"),
        ('4.1甲方应当在签订本合同且收到乙方交付的首期租金及全额合同保证金后15日内在                         （车辆交付地点）将车辆交付给乙方，届时应当检测确认租赁车辆设备及租赁车辆状况，双方无异议后应签署《车辆交接单》（详见附件1）；乙方应签署《车辆交接单》，该清单签署后代表乙方已对车辆质量以及性能',
         f"4.1甲方应当在签订本合同且收到乙方交付的首期租金及全额合同保证金后15日内在 {ctx['yard']}（车辆交付地点）将车辆交付给乙方，届时应当检测确认租赁车辆设备及租赁车辆状况，双方无异议后应签署《车辆交接单》（详见附件1）；乙方应签署《车辆交接单》，该清单签署后代表乙方已对车辆质量以及性能"),
    ]


def fill_contract_docx(contract):
    if not ensure_docx_support():
        raise RuntimeError('当前 Python 环境缺少 python-docx，无法生成 Word 合同')
    template_path = contract_template_path(contract['contract_type'])
    doc = Document(template_path)
    replacements = build_replacements(contract)
    ctx = build_contract_context(contract)

    for paragraph in doc.paragraphs:
        replace_paragraph_text(paragraph, replacements)

    for table in doc.tables:
        replace_in_table(table, replacements)

    # 补充交接单表格常见字段
    for table in doc.tables:
        for row in table.rows:
            cell_text = [cell.text.strip() for cell in row.cells]
            joined = ' | '.join(cell_text)
            if '车型：' in joined and 'VIN：' in joined:
                if len(row.cells) >= 5:
                    row.cells[0].text = f"车型：{ctx['car_type']}"
                    row.cells[1].text = f"车牌：{ctx['plate_number']}"
                    row.cells[2].text = f"VIN：{ctx['vin']}"
                    row.cells[3].text = '颜色：'
                    row.cells[4].text = ''
            if '合同号：' in joined:
                row.cells[-1].text = f"合同号：{ctx['contract_no']}"

    base_name = safe_filename(f"{ctx['contract_no']}_{ctx['customer_name']}_{ctx['contract_type']}")
    docx_path = os.path.join(CONTRACT_OUTPUT_DIR, f"{base_name}.docx")
    doc.save(docx_path)
    return docx_path, ctx


def write_minimal_docx(path, ctx):
    """无 python-docx 时的兜底 Word 导出，保证现场仍能生成可打开的合同文件。"""
    lines = [
        f"{ctx['contract_type']}合同",
        f"合同编号：{ctx['contract_no']}",
        "甲方：陕西金聚源汽车服务有限公司",
        f"乙方：{ctx['customer_name']}",
        f"电话：{ctx['customer_phone']}",
        f"身份证号：{ctx['customer_id_card']}",
        f"车型：{ctx['car_type']}",
        f"车牌号：{ctx['plate_number']}",
        f"车架号/VIN：{ctx['vin']}",
        f"起租/开始日期：{ctx['start_date']}",
        f"结束日期：{ctx['end_date']}",
        f"总价：{ctx['total_price']} 元",
        f"月租：{ctx['rent']} 元",
        f"押金：{ctx['deposit']} 元",
        f"首付：{ctx['down_payment']} 元",
        f"贷款金额：{ctx['loan_amount']} 元",
        f"还款日：每月 {ctx['repayment_day']} 日",
        f"生成时间：{ctx['generated_at']}",
        "说明：当前运行环境缺少 python-docx，本文件为系统兜底生成版本；安装 requirements.txt 后可按原始合同模板填充导出。",
    ]
    paragraphs = ''.join(
        f"<w:p><w:r><w:t>{xml_escape(line)}</w:t></w:r></w:p>"
        for line in lines
    )
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {paragraphs}
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''
    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>{xml_escape(ctx['contract_no'])}</dc:title>
  <dc:creator>金聚源车辆管理系统</dc:creator>
</cp:coreProperties>'''
    app_props = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>金聚源车辆管理系统</Application>
</Properties>'''
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as package:
        package.writestr('[Content_Types].xml', content_types)
        package.writestr('_rels/.rels', rels)
        package.writestr('word/document.xml', document_xml)
        package.writestr('docProps/core.xml', core)
        package.writestr('docProps/app.xml', app_props)


def render_contract_pdf(contract, output_path, ctx):
    if not ensure_pdf_support():
        raise RuntimeError('当前 Python 环境缺少 reportlab，无法生成 PDF 合同')
    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4
    left = 50
    top = height - 50
    line_height = 18
    y = top
    title = f"{ctx['contract_type']}合同"
    c.setFont('STSong-Light', 16)
    c.drawString(left, y, title)
    y -= 30
    c.setFont('STSong-Light', 11)

    lines = [
        f"合同编号：{ctx['contract_no']}",
        f"甲方：陕西金聚源汽车服务有限公司",
        f"乙方：{ctx['customer_name']}    电话：{ctx['customer_phone']}    身份证号：{ctx['customer_id_card']}",
        f"车型：{ctx['car_type']}    车牌号：{ctx['plate_number']}    VIN：{ctx['vin']}",
        f"合同起止：{ctx['start_date']} 至 {ctx['end_date']}",
        f"总价：¥{ctx['total_price']}    月租：¥{ctx['rent']}    押金：¥{ctx['deposit']}",
        f"首付：¥{ctx['down_payment']}    贷款金额：¥{ctx['loan_amount']}    期数：{ctx['loan_periods']}",
        f"还款日：每月 {ctx['repayment_day']} 日    生成时间：{ctx['generated_at']}",
        '',
        '说明：本 PDF 为系统根据原始合同模板自动填充生成的便捷版本，正式签署请同时核对导出的 Word 原件。',
    ]

    for raw_line in lines:
        wrapped = simpleSplit(raw_line, 'STSong-Light', 11, width - 100) or ['']
        for line in wrapped:
            if y < 60:
                c.showPage()
                c.setFont('STSong-Light', 11)
                y = top
            c.drawString(left, y, line)
            y -= line_height

    c.save()


def pdf_text(value):
    raw = str(value).encode('utf-16-be')
    return '<FEFF' + raw.hex().upper() + '>'


def write_minimal_pdf(path, ctx):
    """无 reportlab 时的兜底 PDF 导出，使用内置 CJK 字体描述写入关键合同字段。"""
    lines = [
        f"{ctx['contract_type']}合同",
        f"合同编号：{ctx['contract_no']}",
        "甲方：陕西金聚源汽车服务有限公司",
        f"乙方：{ctx['customer_name']}    电话：{ctx['customer_phone']}",
        f"车型：{ctx['car_type']}    车牌号：{ctx['plate_number']}    VIN：{ctx['vin']}",
        f"合同起止：{ctx['start_date']} 至 {ctx['end_date']}",
        f"总价：{ctx['total_price']} 元    月租：{ctx['rent']} 元    押金：{ctx['deposit']} 元",
        f"首付：{ctx['down_payment']} 元    贷款金额：{ctx['loan_amount']} 元    期数：{ctx['loan_periods']}",
        f"还款日：每月 {ctx['repayment_day']} 日    生成时间：{ctx['generated_at']}",
        "说明：当前运行环境缺少 reportlab，本 PDF 为系统兜底生成版本。",
    ]
    text_ops = ["BT", "/F1 12 Tf", "50 790 Td", "18 TL"]
    for idx, line in enumerate(lines):
        if idx == 0:
            text_ops.extend(["/F1 16 Tf", f"{pdf_text(line)} Tj", "/F1 12 Tf", "T*"])
        else:
            text_ops.extend([f"{pdf_text(line)} Tj", "T*"])
    text_ops.append("ET")
    stream = '\n'.join(text_ops).encode('ascii')
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length " + str(len(stream)).encode('ascii') + b" >>\nstream\n" + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [6 0 R] >>",
        b"<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 2 >> /FontDescriptor 7 0 R >>",
        b"<< /Type /FontDescriptor /FontName /STSong-Light /Flags 4 /FontBBox [0 -120 1000 880] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 880 /StemV 80 >>",
    ]
    pdf = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = [0]
    for idx, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf.extend(f"{idx} 0 obj\n".encode('ascii'))
        pdf.extend(obj)
        pdf.extend(b"\nendobj\n")
    xref_pos = len(pdf)
    pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode('ascii'))
    pdf.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        pdf.extend(f"{offset:010d} 00000 n \n".encode('ascii'))
    pdf.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode('ascii'))
    with open(path, 'wb') as pdf_file:
        pdf_file.write(pdf)


def generate_contract_files(contract):
    ctx = build_contract_context(contract)
    base_name = safe_filename(f"{ctx['contract_no']}_{ctx['customer_name']}_{ctx['contract_type']}")
    docx_path = os.path.join(CONTRACT_OUTPUT_DIR, f"{base_name}.docx")
    try:
        docx_path, ctx = fill_contract_docx(contract)
    except RuntimeError:
        write_minimal_docx(docx_path, ctx)
    pdf_path = docx_path[:-5] + '.pdf'
    try:
        render_contract_pdf(contract, pdf_path, ctx)
    except RuntimeError:
        write_minimal_pdf(pdf_path, ctx)
    return docx_path, pdf_path


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
    stats = {'near': 0, 'due': 0, 'overdue': 0, 'late_fee': 0, 'urge_op': 0, 'urge_sales': 0}
    contract_late = {}  # contract_id -> 累计滞纳金
    c.execute("""
        SELECT r.id, r.contract_id, r.period, r.due_date, r.amount,
               COALESCE(r.paid_amount,0) AS paid_amount, r.status
        FROM repayments r
        WHERE (r.status IN ('待还款','临近还款','还款日','部分核销','预抵','逾期')
               OR r.status LIKE '逾期%日')
    """)
    for r in c.fetchall():
        _process_repayment_row(c, r, today, today_str, stats, contract_late)
    # 同步合同级 late_fee 应收
    for cid, cum in contract_late.items():
        _sync_late_fee_item(c, cid, cum)
    # 厂家还款仍按简单逾期标记
    c.execute("UPDATE factory_repayments SET status='逾期' WHERE status='待还款' AND due_date < ?", (today_str,))
    summary = (f"near={stats['near']} due={stats['due']} overdue={stats['overdue']} "
               f"late_fee_rows={stats['late_fee']} urge_op={stats['urge_op']} urge_sales={stats['urge_sales']}")
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
            if diff == 3:
                stats['urge_op'] += 1
            elif diff == 7:
                stats['urge_sales'] += 1

    if new_status != status:
        c.execute("UPDATE repayments SET status=? WHERE id=?", (new_status, rid))


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


def _sync_late_fee_item(c, cid, cumulative):
    """把合同累计滞纳金(扣除 J3 已生效减免)同步到 contract_fee_items(fee_type='late_fee')。
    应收 = 逐日计提毛额(cumulative) - 已减免总额(late_fee_ledger.waived_amount 之和)。
    日终重算会重复调用本函数，net 计算保证减免不被毛额覆盖（幂等）。"""
    c.execute("SELECT COALESCE(SUM(waived_amount),0) AS w FROM late_fee_ledger WHERE contract_id=? AND waived=1", (cid,))
    waived = parse_money(c.fetchone()['w'])
    net = round(max(0.0, parse_money(cumulative) - waived), 2)
    c.execute("SELECT id FROM contract_fee_items WHERE contract_id=? AND fee_type='late_fee' LIMIT 1", (cid,))
    row = c.fetchone()
    if row:
        c.execute("UPDATE contract_fee_items SET amount_due=? WHERE id=?", (net, row['id']))
    else:
        c.execute("INSERT INTO contract_fee_items (contract_id, fee_type, description, amount_due, amount_paid, status, created_by) "
                  "VALUES (?, 'late_fee', '逾期滞纳金(系统计提)', ?, 0, '待支付', '系统')", (cid, net))


# ======================== 页面路由 ========================
@app.route('/')
def index():
    return send_from_directory('templates', 'index.html')


# ======================== 仪表盘统计 ========================
@app.route('/api/dashboard/stats', methods=['GET'])
def get_stats():
    # 每次查看仪表盘时触发逾期检测
    check_overdue()

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) as cnt FROM vehicles")
    total_vehicles = c.fetchone()['cnt']

    c.execute("SELECT COUNT(*) as cnt FROM vehicles WHERE status NOT IN ('已结清')")
    active_vehicles = c.fetchone()['cnt']

    c.execute("SELECT COALESCE(SUM(invoice_price), 0) as val FROM vehicles")
    total_invoice = c.fetchone()['val']

    c.execute("SELECT COALESCE(SUM(estimated_residual_value), 0) as val FROM vehicles")
    total_residual = c.fetchone()['val']

    c.execute("SELECT COALESCE(SUM(loan_amount), 0) as total_loan, COALESCE(SUM(paid_principal), 0) as total_paid FROM contracts")
    row = c.fetchone()
    total_loan = row['total_loan']
    total_paid_principal = row['total_paid']

    c.execute("SELECT COALESCE(SUM(collected_rent), 0) as rent, COALESCE(SUM(collected_deposit), 0) as deposit FROM contracts")
    row = c.fetchone()
    total_rent = row['rent']
    total_deposit = row['deposit']

    # 客户逾期笔数（逐日 逾期N日 / 旧值 逾期 / 部分核销 均计入）
    c.execute("SELECT COUNT(*) as cnt FROM repayments WHERE status LIKE '逾期%' OR status='部分核销'")
    overdue_count = c.fetchone()['cnt']

    # 厂家逾期笔数
    c.execute("SELECT COUNT(*) as cnt FROM factory_repayments WHERE status = '逾期'")
    factory_overdue_count = c.fetchone()['cnt']

    # 本月客户应收
    c.execute("""SELECT COALESCE(SUM(amount), 0) as val FROM repayments
                 WHERE strftime('%Y-%m', due_date) = strftime('%Y-%m', 'now') AND status != '已还款'""")
    monthly_due = c.fetchone()['val']

    # 本月厂家应付
    c.execute("""SELECT COALESCE(SUM(amount), 0) as val FROM factory_repayments
                 WHERE strftime('%Y-%m', due_date) = strftime('%Y-%m', 'now') AND status != '已还款'""")
    monthly_factory_due = c.fetchone()['val']

    # 利润相关：已收客户租金总额 - 已付厂家月供总额
    c.execute("SELECT COALESCE(SUM(amount), 0) as val FROM repayments WHERE status = '已还款'")
    total_customer_received = c.fetchone()['val']
    c.execute("SELECT COALESCE(SUM(amount), 0) as val FROM factory_repayments WHERE status = '已还款'")
    total_factory_paid = c.fetchone()['val']
    gross_profit = round(total_customer_received - total_factory_paid, 2)

    c.execute("""
        SELECT COUNT(*) as cnt
        FROM vehicles
        WHERE insurance_expiry_date IS NOT NULL
          AND insurance_expiry_date != ''
          AND date(insurance_expiry_date) <= date('now', '+30 day')
    """)
    expiring_insurance_count = c.fetchone()['cnt']

    c.execute("""
        SELECT COUNT(*) as cnt
        FROM sales_orders
        WHERE order_status IN ('待价格特批', '待财务确认', '已激活')
    """)
    open_order_count = c.fetchone()['cnt']

    conn.close()
    return jsonify({
        'total_vehicles': total_vehicles,
        'active_vehicles': active_vehicles,
        'total_invoice': total_invoice,
        'total_residual': round(total_residual, 2),
        'total_loan': total_loan,
        'total_paid_principal': total_paid_principal,
        'total_rent': total_rent,
        'total_deposit': total_deposit,
        'overdue_count': overdue_count,
        'factory_overdue_count': factory_overdue_count,
        'monthly_due': monthly_due,
        'monthly_factory_due': monthly_factory_due,
        'total_customer_received': total_customer_received,
        'total_factory_paid': total_factory_paid,
        'gross_profit': gross_profit,
        'expiring_insurance_count': expiring_insurance_count,
        'open_order_count': open_order_count,
    })


# ======================== 车辆资产 CRUD ========================
@app.route('/api/vehicles', methods=['GET'])
@login_required
def get_vehicles():
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT v.*, c.rental_method, c.business_mode, c.loan_amount, c.monthly_payment, c.rent,
               c.loan_periods, c.deposit, c.paid_principal, c.loan_balance,
               c.collected_deposit, c.collected_rent, c.contract_status,
               cu.name as customer_name
        FROM vehicles v
        LEFT JOIN contracts c ON c.vehicle_id = v.id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        ORDER BY v.id ASC
    """)
    vehicles = [dict(row) for row in c.fetchall()]
    vehicles = redact_for_role(conn, user['role'], 'vehicles', vehicles)
    conn.close()
    return jsonify(vehicles)


# 车型预设列表（按能源类型分组）
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

@app.route('/api/car-types', methods=['GET'])
def get_car_types():
    return jsonify(CAR_TYPE_PRESETS)


@app.route('/api/sales-order-car-options', methods=['GET'])
@login_required
def get_sales_order_car_options():
    return jsonify({
        'categories': SALES_ORDER_CAR_OPTION_CONFIG,
        'common_options': SALES_ORDER_COMMON_OPTIONS,
    })


@app.route('/api/model-guidance-prices', methods=['GET'])
@login_required
def get_model_guidance_prices():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT car_type, guidance_price, lease_installment_price, sale_total_price,
               remark, updated_by, updated_at
        FROM model_guidance_prices
        ORDER BY car_type ASC
    """)
    configured = {row['car_type']: dict(row) for row in c.fetchall()}

    c.execute("""
        SELECT car_type,
               COUNT(*) as vehicle_count,
               AVG(COALESCE(guidance_price, 0)) as avg_vehicle_guidance_price
        FROM vehicles
        WHERE COALESCE(car_type, '') != ''
        GROUP BY car_type
        ORDER BY car_type ASC
    """)
    rows = []
    seen = set()
    for row in c.fetchall():
        item = configured.get(row['car_type'], {
            'car_type': row['car_type'],
            'guidance_price': 0,
            'lease_installment_price': 0,
            'sale_total_price': 0,
            'remark': '',
            'updated_by': '',
            'updated_at': '',
        })
        item['guidance_price'] = item.get('guidance_price') or 0
        item['lease_installment_price'] = item.get('lease_installment_price') or 0
        item['sale_total_price'] = item.get('sale_total_price') or 0
        item['vehicle_count'] = row['vehicle_count']
        item['avg_vehicle_guidance_price'] = row['avg_vehicle_guidance_price'] or 0
        rows.append(item)
        seen.add(row['car_type'])

    for car_type, item in configured.items():
        if car_type not in seen:
            item['guidance_price'] = item.get('guidance_price') or 0
            item['lease_installment_price'] = item.get('lease_installment_price') or 0
            item['sale_total_price'] = item.get('sale_total_price') or 0
            item['vehicle_count'] = 0
            item['avg_vehicle_guidance_price'] = 0
            rows.append(item)

    conn.close()
    return jsonify(rows)


@app.route('/api/guidance-price-alerts', methods=['GET'])
@require_role('老板')
def get_guidance_price_alerts():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT v.id, v.vin, v.plate_number, v.car_type, v.company, v.invoice_date,
               v.invoice_price, v.purchase_price, v.guidance_price, v.status, v.created_at,
               COALESCE(mgp.guidance_price, 0) as model_guidance_price,
               COALESCE(mgp.lease_installment_price, 0) as lease_installment_price,
               COALESCE(mgp.sale_total_price, 0) as sale_total_price
        FROM vehicles v
        LEFT JOIN model_guidance_prices mgp ON mgp.car_type = v.car_type
        WHERE COALESCE(v.guidance_price, 0) <= 0
          AND COALESCE(mgp.guidance_price, 0) <= 0
          AND (
              COALESCE(mgp.lease_installment_price, 0) <= 0
              OR COALESCE(mgp.sale_total_price, 0) <= 0
          )
          AND COALESCE(v.status, '') IN ('在库', '报单锁定中')
        ORDER BY v.created_at DESC, v.id DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify({
        'count': len(rows),
        'items': rows,
        'message': f'有 {len(rows)} 台新入库车辆未设置指导价' if rows else '',
    })


@app.route('/api/model-guidance-prices', methods=['POST'])
@require_role('老板')
def upsert_model_guidance_price():
    data = request.json or {}
    car_type = (data.get('car_type') or '').strip()
    legacy_input = data.get('guidance_price', data.get('price'))
    lease_price = parse_money(data.get('lease_installment_price'))
    sale_price = parse_money(data.get('sale_total_price'))
    new_price = sale_price or lease_price
    remark = (data.get('remark') or '').strip()
    if not car_type:
        return jsonify({'success': False, 'message': '请选择或填写车型'}), 400
    if lease_price <= 0 or sale_price <= 0:
        return jsonify({'success': False, 'message': '租赁每期指导价和以租代售整车指导价都必须大于0'}), 400

    user = request.current_user
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("""
            SELECT guidance_price, lease_installment_price, sale_total_price
            FROM model_guidance_prices
            WHERE car_type=?
        """, (car_type,))
        existing = c.fetchone()
        old_price = existing['guidance_price'] if existing else 0
        old_lease_price = existing['lease_installment_price'] if existing else 0
        old_sale_price = existing['sale_total_price'] if existing else 0

        c.execute("SELECT id, guidance_price FROM vehicles WHERE car_type=?", (car_type,))
        affected_vehicles = c.fetchall()
        affected_count = len(affected_vehicles)

        c.execute("""
            INSERT INTO model_guidance_prices
                (car_type, guidance_price, lease_installment_price, sale_total_price, remark, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(car_type) DO UPDATE SET
                guidance_price=excluded.guidance_price,
                lease_installment_price=excluded.lease_installment_price,
                sale_total_price=excluded.sale_total_price,
                remark=excluded.remark,
                updated_by=excluded.updated_by,
                updated_at=excluded.updated_at
        """, (car_type, new_price, lease_price, sale_price, remark, user['display_name'], now))
        history_rows = [
            ('lease_installment', old_lease_price or old_price or 0, lease_price),
            ('sale_total', old_sale_price or old_price or 0, sale_price),
        ]
        for price_kind, old_value, new_value in history_rows:
            c.execute("""
                INSERT INTO model_guidance_price_history
                    (car_type, price_kind, old_price, new_price, changed_by, effective_at, affected_vehicle_count, remark)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (car_type, price_kind, old_value, new_value, user['display_name'], now, affected_count, remark))

        for vehicle in affected_vehicles:
            c.execute("""
                INSERT INTO vehicle_guidance_price_history
                    (vehicle_id, old_price, new_price, changed_by, effective_at)
                VALUES (?, ?, ?, ?, ?)
            """, (vehicle['id'], vehicle['guidance_price'] or 0, new_price, user['display_name'], now))
        c.execute("UPDATE vehicles SET guidance_price=? WHERE car_type=?", (new_price, car_type))
        remaining_missing_count = unresolved_guidance_vehicle_count(conn)

        log_audit(conn, '更新车型指导价', 'model_guidance_price', None,
                  f'{user["display_name"]} 将车型「{car_type}」租赁每期价调为 ¥{lease_price}，以租代售整车价调为 ¥{sale_price}，同步车辆 {affected_count} 台',
                  user['display_name'])
        conn.commit()
        return jsonify({
            'success': True,
            'message': f'车型指导价已更新，同步 {affected_count} 台库存车辆',
            'affected_vehicle_count': affected_count,
            'remaining_missing_guidance_count': remaining_missing_count,
            'lease_installment_price': lease_price,
            'sale_total_price': sale_price,
        })
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


# 文件上传
@app.route('/api/upload', methods=['POST'])
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

        car_type = (data.get('car_type') or '').strip()
        guidance_price = parse_money(data.get('guidance_price', 0))
        if guidance_price <= 0 and car_type:
            model_price = c.execute(
                "SELECT guidance_price FROM model_guidance_prices WHERE car_type=?",
                (car_type,)
            ).fetchone()
            if model_price and parse_money(model_price['guidance_price']) > 0:
                guidance_price = parse_money(model_price['guidance_price'])

        c.execute('''
        INSERT INTO vehicles (vin, plate_number, company, car_type, is_new, invoice_date,
                              invoice_price, purchase_price, tax_rate, estimated_residual_value,
                              guidance_price, insurance_expiry_date, annual_review_date,
                              engine_number, vehicle_category, vehicle_cab, vehicle_engine_battery,
                              vehicle_power_battery, vehicle_gearbox, vehicle_color, vehicle_box_type,
                              box_type_remark, invoice_contract_file, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            vin, data.get('plate_number'), data.get('company', '陕西金聚源汽车服务有限公司'),
            car_type, data.get('is_new', '新车'), data.get('invoice_date'),
            data.get('invoice_price', 0), data.get('purchase_price', 0), data.get('tax_rate', 0.13),
            data.get('estimated_residual_value', 0), guidance_price,
            data.get('insurance_expiry_date', ''), data.get('annual_review_date', ''),
            data.get('engine_number', ''),
            data.get('vehicle_category', ''),
            data.get('vehicle_cab', ''),
            data.get('vehicle_engine_battery', ''),
            data.get('vehicle_power_battery', ''),
            data.get('vehicle_gearbox', ''),
            data.get('vehicle_color', ''),
            data.get('vehicle_box_type', ''),
            data.get('box_type_remark', ''),
            data.get('invoice_contract_file', ''),
            data.get('status', '在库')
        ))
        vehicle_id = c.lastrowid
        if guidance_price <= 0:
            log_audit(conn, '缺少指导价提醒', 'vehicle', vehicle_id,
                      f'新入库车辆 VIN:{vin} 车型:{car_type or "-"} 未设置指导价，请老板维护指导价')
        conn.commit()
        message = '车辆入库成功'
        if guidance_price <= 0:
            message = '车辆入库成功，该车未设置指导价，已提醒老板维护'
        return jsonify({'success': True, 'id': vehicle_id, 'message': message, 'missing_guidance_price': guidance_price <= 0})
    except Exception as e:
        conn.rollback()
        message = '车辆已在库' if 'UNIQUE constraint failed: vehicles.vin' in str(e) else str(e)
        return jsonify({'success': False, 'message': message}), 400
    finally:
        conn.close()

@app.route('/api/vehicles/<int:vid>', methods=['PUT'])
@require_role('车管')
def update_vehicle(vid):
    data = request.json
    conn = get_db()
    c = conn.cursor()
    fields = []
    values = []
    for key in ['plate_number', 'company', 'car_type', 'is_new', 'invoice_date',
                'invoice_price', 'purchase_price', 'tax_rate', 'guidance_price', 'invoice_contract_file', 'status',
                'insurance_expiry_date', 'annual_review_date', 'engine_number',
                'vehicle_category', 'vehicle_cab', 'vehicle_engine_battery', 'vehicle_power_battery',
                'vehicle_gearbox', 'vehicle_color', 'vehicle_box_type', 'box_type_remark']:
        if key in data:
            fields.append(f"{key} = ?")
            values.append(data[key])
    if fields:
        values.append(vid)
        c.execute(f"UPDATE vehicles SET {', '.join(fields)} WHERE id = ?", values)
        conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/vehicles/<int:vid>', methods=['DELETE'])
@require_role('老板')
def delete_vehicle(vid):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM repayments WHERE contract_id IN (SELECT id FROM contracts WHERE vehicle_id=?)", (vid,))
    c.execute("DELETE FROM factory_repayments WHERE contract_id IN (SELECT id FROM contracts WHERE vehicle_id=?)", (vid,))
    c.execute("DELETE FROM contracts WHERE vehicle_id=?", (vid,))
    c.execute("DELETE FROM vehicles WHERE id=?", (vid,))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/vehicles/<int:vid>/guidance_price', methods=['POST'])
@require_role('老板')
def update_guidance_price(vid):
    data = request.json or {}
    new_price = parse_money(data.get('price'))
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT guidance_price FROM vehicles WHERE id=?", (vid,))
    vehicle = c.fetchone()
    if not vehicle:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    old_price = vehicle['guidance_price'] or 0
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if new_price <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '指导价必须大于0'}), 400
    c.execute("UPDATE vehicles SET guidance_price = ? WHERE id = ?", (new_price, vid))
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

@app.route('/api/vehicles/<int:vid>/activate', methods=['POST'])
@require_role('车管')
def activate_vehicle(vid):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE vehicles SET status='在库', activated_at=? WHERE id=?",
              (datetime.now().strftime('%Y-%m-%d'), vid))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已激活'})


# ======================== 客户 CRUD ========================
@app.route('/api/customers', methods=['GET'])
def get_customers():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM customers ORDER BY id ASC")
    customers = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(customers)

@app.route('/api/customers', methods=['POST'])
def add_customer():
    data = request.json
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute("INSERT INTO customers (name, phone, id_card, address, remark) VALUES (?, ?, ?, ?, ?)",
                  (data.get('name'), data.get('phone'), data.get('id_card'), data.get('address'), data.get('remark')))
        conn.commit()
        return jsonify({'success': True, 'id': c.lastrowid})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
    finally:
        conn.close()


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
        VALUES (?, ?, ?, ?, ?, ?, ?, '待开票', ?, ?)
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
    log_audit(conn, '发起发票申请', 'invoice_request', invoice_id,
              f'合同{cid} 金额{amount} 抬头:{invoice_entity_name}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'id': invoice_id, 'message': '发票申请已提交'})


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
    if row['status'] not in ('待开票', '已开票'):
        conn.close()
        return jsonify({'success': False, 'message': f'当前状态为{row["status"]}，不能作废'}), 400

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
               c.contract_type, c.delivery_status
        FROM sales_orders so
        LEFT JOIN vehicles v ON v.id = so.vehicle_id
        LEFT JOIN contracts c ON c.id = so.contract_id
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
    if len(vin) != 17:
        return jsonify({'success': False, 'message': '请填写17位车架号(VIN)'}), 400
    is_draft = bool(data.get('save_as_draft') or data.get('draft') or data.get('action') == 'draft')

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, plate_number, car_type, status, guidance_price FROM vehicles WHERE vin=?", (vin,))
    vehicle = c.fetchone()
    if not vehicle:
        conn.close()
        return jsonify({'success': False, 'message': '未找到对应库存车辆'}), 404
    if not is_draft and vehicle['status'] != '在库':
        conn.close()
        return jsonify({'success': False, 'message': '仅在库车辆可以发起报单'}), 400

    if not is_draft:
        c.execute("""
            SELECT id FROM contracts
            WHERE vehicle_id=? AND contract_status!='已结清'
            ORDER BY id DESC LIMIT 1
        """, (vehicle['id'],))
        if c.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '该 VIN 存在未结清合同，不能再次报单'}), 400

        c.execute("""
            SELECT id FROM sales_orders
            WHERE vehicle_id=? AND order_status IN ('待价格特批', '待财务确认', '已激活')
            ORDER BY id DESC LIMIT 1
        """, (vehicle['id'],))
        if c.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '该车辆已有进行中的销售报单'}), 400

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

    sale_total_price = parse_money(data.get('sale_total_price'))
    sales_mode = normalize_sales_mode(data.get('sales_mode', '经营租赁'))
    model_dual_ready = has_model_dual_guidance(conn, vehicle['car_type'])
    guidance_prices = resolve_guidance_prices_for_vehicle(conn, vehicle)
    lease_guidance_price = guidance_prices['lease_installment_price']
    sale_guidance_price = guidance_prices['sale_total_price']
    if not is_draft and (not model_dual_ready or lease_guidance_price <= 0 or sale_guidance_price <= 0):
        conn.close()
        return jsonify({'success': False, 'message': '请联系老板维护车型指导价'}), 400

    lease_quote = parse_money(data.get('vehicle_rent_amount', data.get('rent')))
    if sales_mode == '租赁':
        quote_price = lease_quote or sale_total_price
        guidance_price = lease_guidance_price
        guidance_label = '租赁每期指导价'
    else:
        quote_price = sale_total_price
        guidance_price = sale_guidance_price
        guidance_label = '以租代售整车指导价' if sales_mode == '以租代售' else '整车指导价'
    needs_boss_price_approval = (not is_draft) and is_price_below_guidance(quote_price, guidance_price)
    order_status = '草稿' if is_draft else ('待价格特批' if needs_boss_price_approval else '待财务确认')
    price_check_status = '未提交' if is_draft else ('待老板审批' if needs_boss_price_approval else '无需审批')
    price_exception_reason = ''
    if needs_boss_price_approval:
        price_exception_reason = f'报价 ¥{quote_price} 低于{guidance_prices["source"]}{guidance_label} ¥{guidance_price}'

    now = datetime.now()
    saved_at = now.strftime('%Y-%m-%d %H:%M:%S') if is_draft else None
    expires_at = (now + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S') if is_draft else None

    c.execute("""
        INSERT INTO sales_orders
            (payment_date, customer_name, customer_phone, sales_mode, vehicle_id, vin,
             car_type, vehicle_color, plate_number, lease_term, cargo_length, sale_total_price,
             payment_category, car_purchase_amount, vehicle_rent_amount, receiving_company,
             wechat_interest, wechat_registration_fee, wechat_purchase_tax,
             full_package, wechat_private_fee, gifted_items, deposit_amount, order_status,
             snapshot_guidance_price, snapshot_lease_installment_price, snapshot_sale_total_price,
             price_check_status, price_exception_reason, customer_plan_match_status, factory_plan_match_status,
             saved_at, expires_at, sales_advisor, remark, created_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        data.get('payment_date') or datetime.now().strftime('%Y-%m-%d'),
        data.get('customer_name', '').strip() or ('草稿客户' if is_draft else ''),
        data.get('customer_phone', '').strip(),
        sales_mode,
        vehicle['id'],
        vin,
        data.get('car_type', '').strip() or (vehicle['car_type'] or ''),
        data.get('vehicle_color', '').strip(),
        data.get('plate_number', '').strip() or (vehicle['plate_number'] or ''),
        data.get('lease_term', '').strip(),
        data.get('cargo_length', '').strip(),
        sale_total_price,
        data.get('payment_category', '').strip(),
        parse_money(data.get('car_purchase_amount')),
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
        0 if is_draft else (guidance_price if sales_mode == '租赁' else lease_guidance_price),
        0 if is_draft else sale_guidance_price,
        price_check_status,
        price_exception_reason,
        '未生成' if is_draft else '已生成',
        '未上传',
        saved_at,
        expires_at,
        data.get('sales_advisor', '').strip() or user['display_name'],
        data.get('remark', '').strip(),
        user['display_name'],
    ))
    order_id = c.lastrowid
    if needs_boss_price_approval:
        create_approval_flow(conn, 'price_exception', order_id)
    elif not is_draft:
        # 价格无需特批，直接进入待财务确认：同步建财务确认报单待办流
        create_approval_flow(conn, 'sale_payment', order_id)
        ensure_sales_order_planning_contract(conn, order_id)
    if not is_draft:
        c.execute("UPDATE vehicles SET status='报单锁定中' WHERE id=?", (vehicle['id'],))
    log_audit(conn, '创建销售报单', 'sales_order', order_id,
              f"{user['display_name']} 报单 VIN:{vin} 模式:{sales_mode} 报价:{quote_price} {guidance_label}:{guidance_price} 状态:{order_status}",
              user['display_name'])
    conn.commit()
    conn.close()
    if is_draft:
        return jsonify({'success': True, 'id': order_id, 'message': '草稿已保存，7天内有效'})
    message = '销售报单已提交，车辆已锁定'
    if needs_boss_price_approval:
        message = '销售报单已提交，报价低于指导价，等待老板价格特批'
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
    if len(vin) != 17:
        conn.close()
        return jsonify({'success': False, 'message': '请填写17位车架号(VIN)'}), 400
    c.execute("SELECT id, plate_number, car_type, status, guidance_price FROM vehicles WHERE vin=?", (vin,))
    vehicle = c.fetchone()
    if not vehicle:
        conn.close()
        return jsonify({'success': False, 'message': '未找到对应库存车辆'}), 404

    sales_mode = normalize_sales_mode(data.get('sales_mode', order['sales_mode']))
    sale_total_price = parse_money(data.get('sale_total_price'), order['sale_total_price'])
    lease_quote = parse_money(data.get('vehicle_rent_amount', data.get('rent')), order['vehicle_rent_amount'])
    guidance_prices = resolve_guidance_prices_for_vehicle(conn, vehicle)
    lease_guidance_price = guidance_prices['lease_installment_price']
    sale_guidance_price = guidance_prices['sale_total_price']
    quote_price = lease_quote or sale_total_price if sales_mode == '租赁' else sale_total_price
    guidance_price = lease_guidance_price if sales_mode == '租赁' else sale_guidance_price
    guidance_label = '租赁每期指导价' if sales_mode == '租赁' else ('以租代售整车指导价' if sales_mode == '以租代售' else '整车指导价')

    order_status = '草稿'
    price_check_status = '未提交'
    price_exception_reason = ''
    saved_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    expires_at = (datetime.now() + timedelta(days=7)).strftime('%Y-%m-%d %H:%M:%S')
    snapshot_guidance_price = 0
    snapshot_lease_price = 0
    snapshot_sale_price = 0

    if submit_now:
        if vehicle['status'] != '在库':
            conn.close()
            return jsonify({'success': False, 'message': '该车辆已被其他报单或合同占用，请调整 VIN'}), 400
        if not has_model_dual_guidance(conn, vehicle['car_type']) or lease_guidance_price <= 0 or sale_guidance_price <= 0:
            conn.close()
            return jsonify({'success': False, 'message': '请联系老板维护车型指导价'}), 400
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
        c.execute("""
            SELECT id FROM sales_orders
            WHERE vehicle_id=? AND id!=? AND order_status IN ('待价格特批', '待财务确认', '已激活')
            LIMIT 1
        """, (vehicle['id'], order_id))
        if c.fetchone():
            conn.close()
            return jsonify({'success': False, 'message': '该车辆已有进行中的销售报单'}), 400
        needs_boss_price_approval = is_price_below_guidance(quote_price, guidance_price)
        order_status = '待价格特批' if needs_boss_price_approval else '待财务确认'
        price_check_status = '待老板审批' if needs_boss_price_approval else '无需审批'
        if needs_boss_price_approval:
            price_exception_reason = f'报价 ¥{quote_price} 低于{guidance_prices["source"]}{guidance_label} ¥{guidance_price}'
        saved_at = None
        expires_at = None
        snapshot_guidance_price = guidance_price
        snapshot_lease_price = guidance_price if sales_mode == '租赁' else lease_guidance_price
        snapshot_sale_price = sale_guidance_price

    c.execute("""
        UPDATE sales_orders
        SET payment_date=?, customer_name=?, customer_phone=?, sales_mode=?, vehicle_id=?, vin=?,
            car_type=?, vehicle_color=?, plate_number=?, lease_term=?, cargo_length=?, sale_total_price=?,
            payment_category=?, car_purchase_amount=?, vehicle_rent_amount=?, receiving_company=?,
            wechat_interest=?, wechat_registration_fee=?, wechat_purchase_tax=?,
            full_package=?, wechat_private_fee=?, gifted_items=?, deposit_amount=?, order_status=?,
            snapshot_guidance_price=?, snapshot_lease_installment_price=?, snapshot_sale_total_price=?,
            price_check_status=?, price_exception_reason=?, customer_plan_match_status=?, factory_plan_match_status=?,
            saved_at=?, expires_at=?, sales_advisor=?, remark=?
        WHERE id=?
    """, (
        data.get('payment_date') or order['payment_date'] or datetime.now().strftime('%Y-%m-%d'),
        (data.get('customer_name') or order['customer_name'] or '').strip(),
        (data.get('customer_phone') or order['customer_phone'] or '').strip(),
        sales_mode,
        vehicle['id'],
        vin,
        (data.get('car_type') or vehicle['car_type'] or order['car_type'] or '').strip(),
        (data.get('vehicle_color') or order['vehicle_color'] or '').strip(),
        (data.get('plate_number') or vehicle['plate_number'] or order['plate_number'] or '').strip(),
        (data.get('lease_term') or order['lease_term'] or '').strip(),
        (data.get('cargo_length') or order['cargo_length'] or '').strip(),
        sale_total_price,
        (data.get('payment_category') or order['payment_category'] or '').strip(),
        parse_money(data.get('car_purchase_amount'), order['car_purchase_amount']),
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
        '未生成' if order_status == '草稿' else '已生成',
        '未上传',
        saved_at,
        expires_at,
        (data.get('sales_advisor') or order['sales_advisor'] or user['display_name']).strip(),
        (data.get('remark') or order['remark'] or '').strip(),
        order_id,
    ))
    if submit_now:
        if order_status == '待价格特批':
            create_approval_flow(conn, 'price_exception', order_id)
        else:
            create_approval_flow(conn, 'sale_payment', order_id)
            ensure_sales_order_planning_contract(conn, order_id)
        c.execute("UPDATE vehicles SET status='报单锁定中' WHERE id=?", (vehicle['id'],))
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
            (payment_date, customer_name, customer_phone, sales_mode, vehicle_id, vin,
             car_type, vehicle_color, plate_number, lease_term, cargo_length, sale_total_price,
             payment_category, car_purchase_amount, vehicle_rent_amount, receiving_company,
             wechat_interest, wechat_registration_fee, wechat_purchase_tax,
             full_package, wechat_private_fee, gifted_items, deposit_amount, order_status,
             price_check_status, customer_plan_match_status, factory_plan_match_status,
             saved_at, expires_at, sales_advisor, remark, created_by)
        SELECT payment_date, customer_name, customer_phone, sales_mode, vehicle_id, vin,
               car_type, vehicle_color, plate_number, lease_term, cargo_length, sale_total_price,
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

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        UPDATE sales_orders
        SET order_status='已激活', finance_confirmed_by=?, finance_confirmed_at=?
        WHERE id=?
    """, (request.current_user['display_name'], now, order_id))
    c.execute("""
        UPDATE approval_flows
        SET status='已通过', operator_id=?, operator_name=?, comment=COALESCE(NULLIF(comment, ''), '财务确认报单'), acted_at=?
        WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
    """, (request.current_user['id'], request.current_user['display_name'], now, order_id))
    log_audit(conn, '确认销售报单', 'sales_order', order_id,
              f"财务确认报单意向 定金/意向金记录 ¥{parse_money(row['deposit_amount'])}", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '报单已确认，可进入合同生成'})


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
    if contract_type_from_sales_mode(order['sales_mode']) == '销售':
        conn.close()
        return jsonify({'success': False, 'message': '销售报单不需要分期计划'}), 400

    if request.method == 'GET':
        contract_id = order['contract_id'] or ensure_sales_order_planning_contract(conn, order_id)
        c.execute("SELECT * FROM contracts WHERE id=?", (contract_id,))
        contract = dict(c.fetchone())
        c.execute("SELECT * FROM repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
        repayments = [dict(row) for row in c.fetchall()]
        c.execute("SELECT * FROM factory_repayments WHERE contract_id=? ORDER BY period ASC, id ASC", (contract_id,))
        factory_rows = [dict(row) for row in c.fetchall()]
        conn.commit()
        conn.close()
        return jsonify({
            'success': True,
            'contract': contract,
            'repayments': repayments,
            'factory_repayments': factory_rows,
        })

    if order['order_status'] != '待财务确认':
        conn.close()
        return jsonify({'success': False, 'message': f'当前报单状态为{order["order_status"]}，不能调整 F1 计划'}), 400

    data = request.json or {}
    try:
        contract_id = ensure_sales_order_planning_contract(conn, order_id, overrides=data, reset_factory=True)
    except Exception as e:
        conn.rollback()
        conn.close()
        return jsonify({'success': False, 'message': str(e)}), 400
    log_audit(conn, '更新报单F1客户计划', 'sales_order', order_id,
              f'计划合同{contract_id}，厂家表已清空需重新上传', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'F1 客户计划已更新，请重新上传厂家计划表', 'contract_id': contract_id})


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
    c.execute("SELECT id, vehicle_id, order_status FROM sales_orders WHERE id=?", (order_id,))
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
        c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (row['vehicle_id'],))
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
        ORDER BY c.id ASC
    """)
    contracts = [dict(row) for row in c.fetchall()]
    user = get_current_user()
    if user:
        contracts = redact_for_role(conn, user['role'], 'contracts', contracts)
    conn.close()
    return jsonify(contracts)


@app.route('/api/contracts/<int:cid>/export', methods=['POST'])
@require_role('运营')
def export_contract(cid):
    contract = fetch_contract_detail(cid)
    if not contract:
        return jsonify({'success': False, 'message': '合同不存在'}), 404
    try:
        docx_path, pdf_path = generate_contract_files(contract)
    except Exception as e:
        return jsonify({'success': False, 'message': f'生成合同失败: {e}'}), 400

    docx_url = f"/api/contracts/exported/{os.path.basename(docx_path)}"
    pdf_url = f"/api/contracts/exported/{os.path.basename(pdf_path)}"
    conn = get_db()
    log_audit(conn, '导出合同', 'contract', cid,
              f"{request.current_user['display_name']} 导出 Word/PDF 合同", request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({
        'success': True,
        'message': '合同已生成',
        'docx_url': docx_url,
        'pdf_url': pdf_url,
    })


@app.route('/api/contracts/exported/<path:filename>')
def serve_exported_contract(filename):
    return send_from_directory(CONTRACT_OUTPUT_DIR, filename, as_attachment=True)

@app.route('/api/contracts', methods=['POST'])
@require_role('运营')
def add_contract():
    data = request.json
    user = request.current_user
    conn = get_db()
    c = conn.cursor()
    try:
        vehicle_id = data['vehicle_id']
        sales_order_id = data.get('sales_order_id')
        planning_contract_id = None

        if sales_order_id:
            c.execute("SELECT id, order_status, vehicle_id FROM sales_orders WHERE id=?", (sales_order_id,))
            order = c.fetchone()
            if not order:
                conn.close()
                return jsonify({'success': False, 'message': '关联报单不存在'}), 404
            if order['vehicle_id'] != vehicle_id:
                conn.close()
                return jsonify({'success': False, 'message': '报单车辆与合同车辆不一致'}), 400
            if order['order_status'] != '已激活':
                conn.close()
                if order['order_status'] == '待价格特批':
                    return jsonify({'success': False, 'message': '成交价低于指导价，请先由老板完成价格审批'}), 400
                return jsonify({'success': False, 'message': f'当前报单状态为{order["order_status"]}，请先由财务确认报单后再上传线下合同'}), 400
            c.execute("""
                SELECT id
                FROM contracts
                WHERE sales_order_id=? AND contract_status='报单计划中'
                ORDER BY id DESC LIMIT 1
            """, (sales_order_id,))
            planning = c.fetchone()
            planning_contract_id = planning['id'] if planning else None

        # ===== 校验：同一辆车不能重复签约；关联报单的 F1 计划合同可复用 =====
        if planning_contract_id:
            c.execute("""SELECT id, contract_type, contract_status, delivery_status
                         FROM contracts
                         WHERE vehicle_id=? AND contract_status != '已结清' AND id!=?""",
                      (vehicle_id, planning_contract_id))
        else:
            c.execute("""SELECT id, contract_type, contract_status, delivery_status
                         FROM contracts WHERE vehicle_id=? AND contract_status != '已结清'""", (vehicle_id,))
        existing = c.fetchone()
        if existing:
            conn.close()
            status_desc = existing['delivery_status'] or existing['contract_status']
            return jsonify({'success': False, 'message': f'该车辆已有未结清合同（状态: {status_desc}），不能重复签约'}), 400

        contract_type = data.get('contract_type', '租赁')
        rent = data.get('rent', 0)
        monthly_payment = data.get('monthly_payment', 0)
        loan_periods = data.get('loan_periods', 0)
        repayment_day = data.get('repayment_day', 1)
        down_payment = data.get('down_payment', 0)
        deposit = data.get('deposit', 0)
        customer_loan_amount = float(data.get('customer_loan_amount', 0) or 0)
        factory_guarantee_deposit = float(data.get('factory_guarantee_deposit', 0) or 0)
        factory_repayment_months = int(data.get('factory_repayment_months', 0) or 0)
        factory_periods = int(data.get('factory_periods', 0) or 0)
        start_date = data.get('start_date', '') or datetime.now().strftime('%Y-%m-%d')
        expected_profit_floor = parse_money(data.get('expected_profit_floor'), 0)
        expected_profit_ceiling = parse_money(data.get('expected_profit_ceiling'), 999999999)
        contract_file = (data.get('contract_file') or '').strip()
        if not contract_file:
            conn.close()
            return jsonify({'success': False, 'message': '请先上传线下签署的合同文档'}), 400

        start_dt = datetime.strptime(start_date, '%Y-%m-%d')
        end_dt = start_dt + timedelta(days=30 * loan_periods) if loan_periods > 0 else start_dt

        if contract_type != '销售':
            if loan_periods <= 0:
                conn.close()
                return jsonify({'success': False, 'message': '客户分期期数必须大于0'}), 400
            if customer_loan_amount > 0:
                min_customer_payment = customer_loan_amount / loan_periods
                if float(rent or 0) < min_customer_payment:
                    conn.close()
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
        if not customer_id and data.get('customer_name'):
            c.execute("INSERT INTO customers (name, phone) VALUES (?, ?)",
                      (data['customer_name'], data.get('customer_phone', '')))
            customer_id = c.lastrowid

        # PRD: 价格快照 — 成交时复制当前基准价至合同
        c.execute("SELECT car_type, guidance_price, invoice_price FROM vehicles WHERE id=?", (vehicle_id,))
        vrow = c.fetchone()
        snap_guidance = resolve_guidance_price_for_vehicle(conn, vrow)[0] if vrow else 0
        snap_invoice = vrow['invoice_price'] if vrow else 0

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
                    down_payment_status=?, deposit_status=?, delivery_status='待首付款',
                    contract_status=?, sales_order_id=?,
                    snapshot_guidance_price=?, snapshot_invoice_price=?,
                    expected_profit_floor=?, expected_profit_ceiling=?, contract_file=?, created_by=?
                WHERE id=?
            """, (
                vehicle_id, customer_id, contract_type, data.get('business_mode', '转租'),
                data.get('rental_method', '经营租赁'), repayment_day,
                start_date, end_dt.strftime('%Y-%m-%d'),
                data.get('total_price', 0), customer_loan_amount, data.get('loan_amount', 0),
                monthly_payment, rent, loan_periods, data.get('company', ''), data.get('yard', ''), data.get('lease_bank_name', ''), data.get('lease_bank_card_no', ''), factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                '免收' if down_payment == 0 else '待收',
                '免收' if deposit == 0 else '待收',
                contract_status, sales_order_id,
                snap_guidance, snap_invoice,
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
                                   expected_profit_floor, expected_profit_ceiling, contract_file, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                vehicle_id, customer_id, contract_type, data.get('business_mode', '转租'),
                data.get('rental_method', '经营租赁'), repayment_day,
                start_date, end_dt.strftime('%Y-%m-%d'),
                data.get('total_price', 0), customer_loan_amount, data.get('loan_amount', 0),
                monthly_payment, rent, loan_periods, data.get('company', ''), data.get('yard', ''), data.get('lease_bank_name', ''), data.get('lease_bank_card_no', ''), factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                '免收' if down_payment == 0 else '待收',
                '免收' if deposit == 0 else '待收',
                '待首付款',
                contract_status, sales_order_id,
                snap_guidance, snap_invoice,
                expected_profit_floor, expected_profit_ceiling,
                contract_file,
                user['display_name'],
            ))
            contract_id = c.lastrowid

        # F1: 销售合同不生成还款计划；租赁/以租代售先生成未激活客户计划，隔离出库前窗口。
        if contract_type != '销售' and not planning_contract_id:
            if contract_type == '租赁' and deposit > 0:
                c.execute("""
                    INSERT INTO repayments (contract_id, period, due_date, amount, status, remark)
                    VALUES (?, 0, NULL, ?, '未激活', '押金')
                """, (contract_id, deposit))
            if contract_type == '以租代售' and down_payment > 0:
                c.execute("""
                    INSERT INTO repayments (contract_id, period, due_date, amount, status, remark)
                    VALUES (?, 0, NULL, ?, '未激活', '首付款')
                """, (contract_id, down_payment))

            for p in range(1, loan_periods + 1):
                try:
                    if contract_type == '租赁':
                        due_dt = start_dt + relativedelta(months=p - 1)
                    else:
                        due_dt = start_dt + relativedelta(months=p)
                        due_dt = due_dt.replace(day=min(repayment_day, 28))
                except Exception:
                    due_dt = start_dt + timedelta(days=30 * (p - 1 if contract_type == '租赁' else p))
                c.execute("""
                    INSERT INTO repayments (contract_id, period, due_date, amount, status)
                    VALUES (?, ?, ?, ?, '未激活')
                """, (contract_id, p, due_dt.strftime('%Y-%m-%d'), rent))

            for p in range(1, factory_periods + 1):
                try:
                    due_dt = start_dt + relativedelta(months=p)
                    due_dt = due_dt.replace(day=min(repayment_day, 28))
                except Exception:
                    due_dt = start_dt + timedelta(days=30 * p)
                due_str = due_dt.strftime('%Y-%m-%d')
                if monthly_payment > 0 and p <= factory_periods:
                    c.execute("""
                        INSERT INTO factory_repayments (contract_id, period, due_date, amount, status)
                        VALUES (?, ?, ?, ?, '待还款')
                    """, (contract_id, p, due_str, monthly_payment))

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
                SET order_status='已激活', contract_id=?
                WHERE id=?
            """, (contract_id, sales_order_id))

        log_audit(conn, '上传线下合同', 'contract', contract_id,
                  f'类型{contract_type} 车辆{vehicle_id} 月租{rent} 月供{monthly_payment} 期数{loan_periods} 合同附件:{contract_file}')
        conn.commit()
        return jsonify({'success': True, 'id': contract_id, 'message': '线下合同已上传，请运营发起首次付款/押金审核'})
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
    c.execute("SELECT contract_id, amount, status FROM repayments WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '记录不存在'}), 404
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
    if not local_path.lower().endswith(('.xlsx', '.pdf')):
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

    fallback_amount = parse_money(contract['monthly_payment'])
    prepared_rows = []
    for row in rows:
        amount = parse_money(row['amount'], fallback_amount)
        if amount <= 0:
            amount = fallback_amount
        source_status = row.get('source_status')
        remark_parts = []
        if import_meta.get('source_format') == 'factory_plan_pdf':
            remark_parts.append(f"PDF状态:{source_status or '未知'}")
            if row.get('principal') is not None:
                remark_parts.append(f"本金:{row.get('principal')}")
            if row.get('interest') is not None:
                remark_parts.append(f"利息:{row.get('interest')}")
        prepared_rows.append((cid, row['period'], row['due_date'], amount, '待还款', None, ';'.join(remark_parts) or None))

    c.execute("DELETE FROM factory_repayments WHERE contract_id=?", (cid,))
    c.executemany("""
        INSERT INTO factory_repayments (contract_id, period, due_date, amount, status, paid_at, remark)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, prepared_rows)
    c.execute("""
        UPDATE contracts
        SET factory_periods=?, factory_repayment_months=?
        WHERE id=?
    """, (len(prepared_rows), len(prepared_rows), cid))
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
    if current == '已通过':
        conn.close()
        return jsonify({'success': False, 'message': '计划已通过，无需确认差异'}), 400

    summary = {}
    if contract['plan_compare_summary']:
        try:
            summary = json.loads(contract['plan_compare_summary'])
        except ValueError:
            summary = {}
    summary.update({
        'status': '差异已确认',
        'difference_confirmed_by': request.current_user['display_name'],
        'difference_confirmed_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
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
               v.invoice_price, v.purchase_price, v.tax_rate, v.estimated_residual_value,
               c.id AS contract_id, c.contract_type, c.business_mode, c.rent, c.monthly_payment,
               c.loan_periods, c.deposit, c.down_payment, c.contract_status,
               COALESCE((SELECT SUM(amount) FROM repayments WHERE contract_id=c.id AND period>=1), 0) AS customer_schedule_total,
               COALESCE((SELECT SUM(amount) FROM factory_repayments WHERE contract_id=c.id), 0) AS factory_schedule_total,
               COALESCE((SELECT SUM(COALESCE(paid_amount, 0)) FROM repayments WHERE contract_id=c.id), 0) AS customer_received,
               COALESCE((SELECT SUM(amount) FROM factory_repayments WHERE contract_id=c.id AND status='已还款'), 0) AS factory_paid,
               COALESCE((SELECT SUM(rebate_amount) FROM vehicle_rebates WHERE vehicle_id=v.id), 0) AS rebate_total
        FROM vehicles v
        JOIN contracts c ON c.vehicle_id = v.id
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
    scalar_queries = {
        'vehicle_count': "SELECT COUNT(*) FROM vehicles",
        'active_contract_count': "SELECT COUNT(*) FROM contracts WHERE contract_status='执行中'",
        'open_order_count': "SELECT COUNT(*) FROM sales_orders WHERE order_status IN ('待价格特批','待财务确认','已激活')",
        'overdue_repayment_count': "SELECT COUNT(*) FROM repayments WHERE status LIKE '逾期%'",
        'pending_invoice_count': "SELECT COUNT(*) FROM invoice_requests WHERE status='待开票'",
        'pending_waiver_count': "SELECT COUNT(*) FROM waivers WHERE status IN ('待审批','已通过')",
        'pending_return_count': "SELECT COUNT(*) FROM return_inspections WHERE status NOT IN ('已完成','已入库')",
        'pending_lock_count': "SELECT COUNT(*) FROM lock_requests WHERE status IN ('待运营审核','待老板审批','已批准待执行')",
    }
    overview = {}
    for key, sql in scalar_queries.items():
        c.execute(sql)
        overview[key] = c.fetchone()[0]
    overview['missing_guidance_count'] = unresolved_guidance_vehicle_count(conn)
    c.execute("SELECT COALESCE(SUM(rebate_amount),0) FROM vehicle_rebates")
    overview['rebate_total'] = round(parse_money(c.fetchone()[0]), 2)
    c.execute("SELECT COALESCE(SUM(amount),0) FROM repayments WHERE status='已还款'")
    overview['customer_received'] = round(parse_money(c.fetchone()[0]), 2)
    c.execute("SELECT COALESCE(SUM(amount),0) FROM factory_repayments WHERE status='已还款'")
    overview['factory_paid'] = round(parse_money(c.fetchone()[0]), 2)
    overview['realized_cash_profit'] = round(overview['customer_received'] - overview['factory_paid'] + overview['rebate_total'], 2)

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
        ORDER BY ot.id DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()
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
        WHERE contract_id=? AND period>=1
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
        'note': '以租代售无押金项；首付款已作为独立行核销，不重复扣减。',
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
    if not customer_screenshot_path:
        conn.close()
        return jsonify({'success': False, 'message': '请上传客户提前结清付款截图'}), 400
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
          AND period>=1
          AND NOT (
              status IN ('已还款', '预抵')
              OR COALESCE(paid_amount, 0) >= COALESCE(amount, 0)
          )
    """, (cid,))
    if c.fetchone()['cnt'] > 0:
        conn.close()
        return jsonify({'success': False, 'message': '客户分期尚未全部结清，不能发起过户'}), 400

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

    transfer_date = data.get('transfer_date') or datetime.now().strftime('%Y-%m-%d')
    c.execute("""
        UPDATE ownership_transfers
        SET status='已过户',
            transfer_date=?,
            new_owner_name=?,
            new_owner_id_card=?,
            transfer_doc_path=?
        WHERE id=?
    """, (
        transfer_date,
        (data.get('new_owner_name') or '').strip(),
        (data.get('new_owner_id_card') or '').strip(),
        (data.get('transfer_doc_path') or '').strip(),
        tid,
    ))
    c.execute("UPDATE contracts SET contract_status='已结清' WHERE id=?", (transfer['contract_id'],))
    if transfer['vehicle_id']:
        c.execute("UPDATE vehicles SET status='已售/已过户' WHERE id=?", (transfer['vehicle_id'],))
    log_audit(conn, '完成过户', 'ownership_transfer', tid,
              f'过户日期:{transfer_date}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '过户已完成'})


# ======================== 风控逾期概览 ========================
@app.route('/api/risk/overdue', methods=['GET'])
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
        WHERE r.status LIKE '逾期%' OR r.status='部分核销'
        ORDER BY r.due_date ASC
    """)
    overdue = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(overdue)

@app.route('/api/risk/factory-overdue', methods=['GET'])
def get_factory_overdue():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT fr.*, c.vehicle_id, v.vin, v.plate_number, v.car_type
        FROM factory_repayments fr
        JOIN contracts c ON c.id = fr.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        WHERE fr.status = '逾期'
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
        WHERE (insurance_expiry_date IS NOT NULL AND insurance_expiry_date != '')
           OR (annual_review_date IS NOT NULL AND annual_review_date != '')
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
        WHERE r.status='待还款' OR r.status LIKE '逾期%' OR r.status IN ('临近还款','还款日','部分核销')
        ORDER BY r.due_date ASC
    """)
    bills = [dict(row) for row in c.fetchall()]
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
               r.verified_by, r.verified_at, r.paid_amount, r.verified_amount, r.waterfall_summary,
               v.vin, v.plate_number, v.car_type,
               cu.name as customer_name, cu.phone as customer_phone
        FROM repayments r
        JOIN contracts c ON c.id = r.contract_id
        JOIN vehicles v ON v.id = c.vehicle_id
        LEFT JOIN customers cu ON cu.id = c.customer_id
        ORDER BY r.due_date DESC
    """)
    rows = [dict(row) for row in c.fetchall()]
    # 计算每条记录的核销步骤进度
    for row in rows:
        step = 0
        if row.get('screenshot_path'): step = 1
        if row.get('bank_receipt_path'): step = 2
        if row.get('bank_serial'): step = 3
        if row.get('status') == '已还款': step = 4
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
    """步骤1：运营发起对账并上传客户付款截图"""
    data = request.json
    screenshot_path = data.get('screenshot_path', '')
    if not screenshot_path:
        return jsonify({'success': False, 'message': '请上传付款截图'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE repayments SET screenshot_path=? WHERE id=?", (screenshot_path, rid))
    log_audit(conn, '上传付款截图', 'repayment', rid, f'截图: {screenshot_path}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '付款截图上传成功'})


@app.route('/api/reconciliation/<int:rid>/receipt', methods=['POST'])
@require_role('财务')
def upload_receipt(rid):
    """步骤2：财务上传银行回单"""
    data = request.json
    bank_receipt_path = data.get('bank_receipt_path', '')
    if not bank_receipt_path:
        return jsonify({'success': False, 'message': '请上传银行回单'}), 400
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE repayments SET bank_receipt_path=? WHERE id=?", (bank_receipt_path, rid))
    log_audit(conn, '上传银行回单', 'repayment', rid, f'回单: {bank_receipt_path}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '银行回单上传成功'})


@app.route('/api/reconciliation/<int:rid>/verify', methods=['POST'])
@require_role('财务')
def verify_reconciliation(rid):
    """步骤3：录入流水号并自动核销"""
    data = request.json
    bank_serial = data.get('bank_serial', '').strip()
    if not bank_serial:
        return jsonify({'success': False, 'message': '请输入银行流水号'}), 400
    conn = get_db()
    c = conn.cursor()
    # 检查前两步是否完成
    c.execute("SELECT screenshot_path, bank_receipt_path, contract_id, amount, status FROM repayments WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '记录不存在'}), 404
    if row['status'] == '已还款':
        conn.close()
        return jsonify({'success': False, 'message': '该笔账单已核销，请勿重复操作'}), 400
    if not row['screenshot_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先上传付款截图'}), 400
    if not row['bank_receipt_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先上传银行回单'}), 400
    if len(bank_serial) < 4:
        conn.close()
        return jsonify({'success': False, 'message': '银行流水号至少填写4位'}), 400

    user = request.current_user['display_name']
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    contract_id = row['contract_id']
    amount = row['amount']
    received_amount = parse_money(data.get('received_amount'), amount)
    if received_amount <= 0:
        conn.close()
        return jsonify({'success': False, 'message': '到账金额必须大于0'}), 400

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
        amount = data.get('amount', default_amount)
        try:
            amount = float(amount or 0)
        except (TypeError, ValueError):
            return jsonify({'success': False, 'message': '付款金额格式不正确'}), 400
        if amount <= 0:
            return jsonify({'success': False, 'message': '首次付款金额必须大于0'}), 400
        if amount < default_amount:
            return jsonify({'success': False, 'message': f'首次付款金额不足，应收 ¥{round(default_amount, 2)}'}), 400

        payment_type = data.get('payment_type') or initial_payment_label(contract)
        c.execute("""
            INSERT INTO contract_initial_payments
                (contract_id, payment_type, amount, customer_screenshot_path, status, requested_by, remark)
            VALUES (?, ?, ?, ?, '审批中', ?, ?)
        """, (cid, payment_type, amount, screenshot_path, user['display_name'], data.get('remark', '')))
        payment_id = c.lastrowid
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
    """财务上传公司到账/银行回单，之后才能审核通过首付款。"""
    data = request.json or {}
    receipt_path = data.get('bank_receipt_path') or data.get('receipt_path') or ''
    if not receipt_path:
        return jsonify({'success': False, 'message': '请上传公司收款回单'}), 400
    bank_serial = (data.get('bank_serial') or '').strip()
    if len(bank_serial) < 4:
        return jsonify({'success': False, 'message': '银行流水号至少填写4位'}), 400
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
    if received_amount < expected_amount:
        conn.close()
        return jsonify({'success': False, 'message': f'首次付款到账不足，应收 ¥{round(expected_amount, 2)}'}), 400

    c.execute("""
        UPDATE contract_initial_payments
        SET bank_receipt_path=?, bank_serial=?, received_amount=?
        WHERE id=?
    """, (receipt_path, bank_serial, received_amount, pid))
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
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, status, pre_repair_status FROM vehicles WHERE id=?", (vid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '车辆不存在'}), 404
    if row['status'] != '维修中':
        conn.close()
        return jsonify({'success': False, 'message': '车辆不在维修中'}), 400
    previous = row['pre_repair_status'] or '在库'
    next_status = '在库' if previous == '待维修' else previous
    c.execute("UPDATE vehicles SET status=?, pre_repair_status=NULL WHERE id=?", (next_status, vid))
    log_audit(conn, '车辆维修完成', 'vehicle', vid,
              f'恢复状态:{next_status}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '维修已完成', 'status': next_status})


# ======================== 租期结束车辆入库（车管操作）========================
@app.route('/api/vehicles/<int:vid>/return-stock', methods=['POST'])
@require_role('车管')
def return_stock(vid):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE vehicles SET status='在库' WHERE id=?", (vid,))
    c.execute("UPDATE contracts SET contract_status='已结清' WHERE vehicle_id=? AND contract_status='执行中'", (vid,))
    log_audit(conn, '租期结束入库', 'vehicle', vid, '车管确认车辆归还入库')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已入库'})


# ======================== 退还车辆验收单 ========================
@app.route('/api/return-inspections', methods=['GET'])
def get_return_inspections():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT ri.*, v.plate_number as vehicle_plate_number, v.vin as vehicle_vin, v.car_type as vehicle_car_type,
               v.company as vehicle_company, c.company as contract_company, c.yard as contract_yard,
               c.lease_bank_name, c.lease_bank_card_no
        FROM return_inspections ri
        LEFT JOIN vehicles v ON v.id = ri.vehicle_id
        LEFT JOIN contracts c ON c.id = ri.contract_id
        ORDER BY ri.id DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
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
            LEFT JOIN contracts c ON c.vehicle_id=v.id AND c.contract_status != '已结清'
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
def update_return_inspection(rid):
    data = request.json
    conn = get_db()
    c = conn.cursor()
    user = get_current_user()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404

    fields = []
    values = []
    for key in ['plate_number','customer_name','rental_period','vin','car_type','company','lease_bank_name','lease_bank_card_no',
                'refund_company_name','refund_bank_name','refund_bank_card_no',
                'return_reason','tool_triangle','tool_vest','tool_extinguisher','tool_wedge','tool_jack',
                'doc_license','doc_keys','mileage','body_tire_clean',
                'accident_info','insurance_surcharge','violation_info','etc_info','maintenance_info',
                'rent_late_fee','return_late_fee','deposit_rent_receivable','deposit_paid',
                'total_deduction','actual_refund','needs_repair','repair_reason','remark','sales_advisor']:
        if key in data:
            fields.append(f"{key}=?")
            values.append(data[key])

    if user and user['role'] == '销售':
        fields.append("sales_status=?")
        values.append('已登记')
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
        update_return_inspection_status(conn, rid)
        conn.commit()
    conn.close()
    return jsonify({'success': True})


@app.route('/api/return-inspections/<int:rid>/fleet', methods=['POST'])
@require_role('车管')
def update_return_fleet(rid):
    data = request.json or {}
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
            doc_license=?, doc_keys=?,
            mileage=?, body_tire_clean=?, accident_info=?, insurance_surcharge=?,
            violation_info=?, etc_info=?, maintenance_info=?,
            needs_repair=?, repair_reason=?,
            status='待运营填写'
        WHERE id=?
    """, (
        1 if data.get('tool_triangle') else 0,
        1 if data.get('tool_vest') else 0,
        1 if data.get('tool_extinguisher') else 0,
        1 if data.get('tool_wedge') else 0,
        1 if data.get('tool_jack') else 0,
        1 if data.get('doc_license') else 0,
        1 if data.get('doc_keys') else 0,
        data.get('mileage', ''),
        data.get('body_tire_clean', ''),
        data.get('accident_info', ''),
        data.get('insurance_surcharge', ''),
        data.get('violation_info', ''),
        data.get('etc_info', ''),
        data.get('maintenance_info', ''),
        1 if data.get('needs_repair') else 0,
        data.get('repair_reason', ''),
        rid,
    ))
    if data.get('needs_repair'):
        c.execute("""
            UPDATE vehicles
            SET status='待维修'
            WHERE id=(SELECT vehicle_id FROM return_inspections WHERE id=?)
        """, (rid,))
    update_return_inspection_status(conn, rid)
    log_audit(conn, '车管验车', 'return_inspection', rid, f"车管填写验车单 {data.get('plate_number', '')}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车管验车内容已保存'})


@app.route('/api/return-inspections/<int:rid>/operator', methods=['POST'])
@require_role('运营')
def update_return_operator(rid):
    data = request.json or {}
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
        data.get('rent_late_fee', 0),
        data.get('return_late_fee', 0),
        data.get('deposit_rent_receivable', 0),
        data.get('deposit_paid', 0),
        data.get('total_deduction', 0),
        data.get('actual_refund', 0),
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
        (request.json or {}).get('comment', ''),
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


@app.route('/api/return-inspections/<int:rid>/pay', methods=['POST'])
@require_role('财务')
def pay_return_refund(rid):
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM return_inspections WHERE id=?", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '退车单不存在'}), 404
    if row['status'] != '待出款':
        conn.close()
        return jsonify({'success': False, 'message': '当前退车单未到财务出款环节'}), 400
    if not row['boss_approved']:
        conn.close()
        return jsonify({'success': False, 'message': '请先完成领导审批'}), 400
    c.execute("""
        UPDATE return_inspections
        SET paid_out=1,
            paid_out_by=?,
            paid_out_at=?,
            status='已完成'
        WHERE id=?
    """, (
        request.current_user['display_name'],
        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        rid,
    ))
    if row['vehicle_id']:
        next_vehicle_status = '待维修' if row['needs_repair'] else '在库'
        c.execute("UPDATE vehicles SET status=?, is_new='二手车' WHERE id=?", (next_vehicle_status, row['vehicle_id']))
    if row['contract_id']:
        c.execute("UPDATE contracts SET contract_status='已结清', delivery_status='已完成' WHERE id=?", (row['contract_id'],))
    log_audit(conn, '退车财务出款', 'return_inspection', rid, f"财务出款完成 {row['plate_number']}")
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '退车已完成'})


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

        if base_ref_type in ('price_exception', 'sale_payment'):
            c.execute("""
                SELECT so.*, v.vin as vehicle_vin, v.plate_number as vehicle_plate_number,
                       v.car_type as vehicle_car_type, v.guidance_price as current_guidance_price
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
                'total_price': row.get('sale_total_price', 0),
                'snapshot_guidance_price': row.get('snapshot_guidance_price', 0),
                'snapshot_lease_installment_price': row.get('snapshot_lease_installment_price', 0),
                'snapshot_sale_total_price': row.get('snapshot_sale_total_price', 0),
                'price_check_status': row.get('price_check_status', ''),
                'price_exception_reason': row.get('price_exception_reason', ''),
                'business_mode': row.get('sales_mode', ''),
                'created_at': row.get('created_at', ''),
                'requested_by': row.get('created_by', ''),
                'follow_up_role': '财务' if row.get('order_status') == '待财务确认' else '',
            })
            return item

        if base_ref_type == 'initial_payment':
            c.execute("""
                SELECT ip.id as initial_payment_id, ip.payment_type, ip.amount as initial_payment_amount,
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

        return item

    result = []
    for (base_ref_type, base_ref_id), steps in grouped.items():
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
            c.execute("""
                SELECT id FROM approval_flows
                WHERE ref_type='sale_payment' AND ref_id=? AND status='待审批'
                LIMIT 1
            """, (ref_id,))
            if not c.fetchone():
                create_approval_flow(conn, 'sale_payment', ref_id)
            ensure_sales_order_planning_contract(conn, ref_id)
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
            c.execute("""
                UPDATE sales_orders
                SET order_status='已激活', finance_confirmed_by=?, finance_confirmed_at=?
                WHERE id=? AND order_status='待财务确认'
            """, (user['display_name'], now, ref_id))
            message = '报单已确认，可进入合同生成'
        elif ref_type == 'initial_payment':
            c.execute("""
                SELECT ip.*, c.contract_type, c.customer_id, c.deposit, c.down_payment, c.rent
                FROM contract_initial_payments ip
                JOIN contracts c ON c.id = ip.contract_id
                WHERE ip.id=?
            """, (ref_id,))
            payment = c.fetchone()
            if payment:
                contract_id = payment['contract_id']
                expected_amount = get_initial_payment_amount(payment)
                received_amount = parse_money(payment['received_amount'])
                if received_amount < expected_amount:
                    conn.rollback()
                    conn.close()
                    return jsonify({'success': False, 'message': f'首次付款到账不足，应收 ¥{round(expected_amount, 2)}'}), 400
                updates = ["delivery_status='待出库'"]
                params = []
                if (payment['deposit'] or 0) > 0:
                    updates.append("deposit_status='已收'")
                    updates.append("collected_deposit=?")
                    params.append(payment['deposit'])
                if (payment['down_payment'] or 0) > 0 or payment['contract_type'] == '销售':
                    updates.append("down_payment_status='已收'")
                if payment['contract_type'] == '租赁' and (payment['rent'] or 0) > 0:
                    updates.append("collected_rent=COALESCE(collected_rent,0)+?")
                    params.append(payment['rent'])
                params.append(contract_id)
                c.execute(f"UPDATE contracts SET {', '.join(updates)} WHERE id=?", params)
                if payment['contract_type'] != '销售':
                    c.execute("""
                        UPDATE repayments
                        SET status='已还款',
                            paid_amount=amount,
                            verified_amount=CASE WHEN COALESCE(verified_amount, 0)=0 THEN amount ELSE verified_amount END,
                            paid_at=?,
                            screenshot_path=COALESCE(screenshot_path, ?),
                            bank_receipt_path=COALESCE(bank_receipt_path, ?),
                            bank_serial=COALESCE(bank_serial, ?),
                            verified_by=COALESCE(verified_by, ?),
                            verified_at=COALESCE(verified_at, ?),
                            waterfall_summary=COALESCE(waterfall_summary, 'initial_payment 押金/首付款独立核销'),
                            remark=COALESCE(remark, CASE WHEN ?='以租代售' THEN '首付款' ELSE '押金' END)
                        WHERE contract_id=? AND period=0
                    """, (
                        datetime.now().strftime('%Y-%m-%d'),
                        payment['customer_screenshot_path'],
                        payment['bank_receipt_path'],
                        payment['bank_serial'] or f"INITIAL-{ref_id}",
                        user['display_name'],
                        now,
                        payment['contract_type'],
                        contract_id,
                    ))
                if payment['contract_type'] == '租赁' and (payment['rent'] or 0) > 0:
                    # 首次付款包含“首次支付金额/首期租金”时，直接核销第 1 期客户还款，
                    # 避免后续对账再核销同一笔租金造成重复入账。
                    c.execute("""
                        UPDATE repayments
                        SET status='已还款',
                            paid_amount=amount,
                            verified_amount=CASE WHEN COALESCE(verified_amount, 0)=0 THEN amount ELSE verified_amount END,
                            paid_at=?,
                            screenshot_path=COALESCE(screenshot_path, ?),
                            bank_receipt_path=COALESCE(bank_receipt_path, ?),
                            bank_serial=COALESCE(bank_serial, ?),
                            verified_by=COALESCE(verified_by, ?),
                            verified_at=COALESCE(verified_at, ?),
                            waterfall_summary=COALESCE(waterfall_summary, 'initial_payment 首期租金自动核销'),
                            remark=COALESCE(remark, '首次付款审核自动核销首期租金')
                        WHERE id=(
                            SELECT id FROM repayments
                            WHERE contract_id=? AND period>=1 AND status!='已还款'
                            ORDER BY period ASC
                            LIMIT 1
                        )
                    """, (
                        datetime.now().strftime('%Y-%m-%d'),
                        payment['customer_screenshot_path'],
                        payment['bank_receipt_path'],
                        payment['bank_serial'] or f"INITIAL-{ref_id}",
                        user['display_name'],
                        now,
                        contract_id,
                    ))
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
                        payment['bank_serial'] or f"INITIAL-{ref_id}",
                        extra_amount,
                        '首次付款多收款挂账，待客户确认抵充月份',
                    ))
                c.execute("""
                    UPDATE repayments
                    SET status='待还款'
                    WHERE contract_id=? AND status='未激活'
                """, (contract_id,))
                c.execute("UPDATE contract_initial_payments SET status='已通过', approved_by=?, approved_at=? WHERE id=?",
                          (user['display_name'], now, ref_id))
            message = '首次付款审核完成，等待车管出库'
        elif ref_type == 'return_stock':
            # 退车领导审批通过 → 待财务出款
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
            message = '领导审批通过，等待财务出款'
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

    log_audit(conn, '审批通过', ref_type, ref_id,
              f'{user["display_name"]}({user["role"]}) 通过 {flow["step_label"]}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': message})


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
    if ref_type == 'price_exception':
        c.execute("SELECT vehicle_id FROM sales_orders WHERE id=?", (ref_id,))
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
        if order and order['vehicle_id']:
            c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (order['vehicle_id'],))
    elif ref_type == 'sale_payment':
        c.execute("SELECT vehicle_id FROM sales_orders WHERE id=?", (ref_id,))
        order = c.fetchone()
        c.execute("""
            UPDATE sales_orders
            SET order_status='已作废',
                voided_at=?,
                void_reason=?,
                voided_by=?
            WHERE id=?
        """, (now, comment, user['display_name'], ref_id))
        if order and order['vehicle_id']:
            c.execute("UPDATE vehicles SET status='在库' WHERE id=? AND status='报单锁定中'", (order['vehicle_id'],))
    elif ref_type == 'contract_delivery':
        c.execute("UPDATE contracts SET delivery_status='已驳回' WHERE id=?", (ref_id,))
    elif ref_type == 'initial_payment':
        c.execute("UPDATE contract_initial_payments SET status='已驳回' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE contracts
            SET delivery_status='首付已驳回'
            WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
        """, (ref_id,))
    elif ref_type == 'return_stock':
        c.execute("UPDATE return_inspections SET boss_approved=0, status='待领导审批' WHERE id=?", (ref_id,))

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
        c.execute("UPDATE return_inspections SET status='待领导审批' WHERE id=?", (ref_id,))

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
    if not ct or ct['delivery_status'] != '待出库':
        conn.close()
        return jsonify({'success': False, 'message': '合同未通过审批或尚未到出库步骤'}), 400
    if not ct['delivery_photo_path'] or not ct['delivery_document_path']:
        conn.close()
        return jsonify({'success': False, 'message': '请先上传出库照片和出库单照片，再执行出库'}), 400

    c.execute("UPDATE contracts SET delivery_status='已出库', delivery_date=? WHERE id=?",
              (datetime.now().strftime('%Y-%m-%d'), ct['id']))

    # 更新车辆状态
    contract_type = ct['contract_type']
    status_map = {'销售': '已售', '以租代售': '以租代售', '租赁': '租赁中'}
    new_status = status_map.get(contract_type, '租赁中')
    c.execute("UPDATE vehicles SET status=? WHERE id=?", (new_status, vid))

    # 卖车合同出库后直接标记已结清
    if contract_type == '销售':
        c.execute("UPDATE contracts SET contract_status='已结清' WHERE id=?", (ct['id'],))

    log_audit(conn, '车辆出库', 'vehicle', vid, f'{user["display_name"]}({user["role"]}) 确认出库 合同类型:{contract_type}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '车辆已出库'})


# ======================== 逾期催促 T+N ========================
@app.route('/api/repayments/<int:rid>/urge', methods=['POST'])
@login_required
def urge_repayment(rid):
    """H2 催款执行（新口径）：diff>=3 运营催款 / diff>=7 销售催款。
    催款仅推送任务，不动押金/首付款，与滞纳金计提是两套独立节奏。"""
    user = request.current_user
    conn = get_db()
    c = conn.cursor()

    c.execute("""SELECT r.*, c.vehicle_id, c.id as cid FROM repayments r
                 JOIN contracts c ON c.id = r.contract_id WHERE r.id=?""", (rid,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'success': False, 'message': '还款记录不存在'}), 404

    overdue_days = max(0, (datetime.now() - datetime.strptime(row['due_date'], '%Y-%m-%d')).days)

    # 确定催款类型（H2 新口径：diff=3 运营催款 / diff=7 销售催款）
    if overdue_days >= 7 and user['role'] in ('销售', '老板'):
        urge_type = '销售催款'
        urge_day = 7
    elif overdue_days >= 3 and user['role'] in ('运营', '老板'):
        urge_type = '运营催款'
        urge_day = 3
    else:
        conn.close()
        return jsonify({'success': False, 'message': f'当前逾期{overdue_days}天，不满足催款条件或角色不匹配'}), 400

    c.execute("""INSERT INTO urge_records (repayment_id, contract_id, vehicle_id, urge_type, urge_day, operator_id, operator_name)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
              (rid, row['cid'], row['vehicle_id'], urge_type, urge_day, user['id'], user['display_name']))

    log_audit(conn, '催款', 'repayment', rid,
              f'{user["display_name"]}({urge_type}) 逾期{overdue_days}天')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'{urge_type} 已记录'})


@app.route('/api/repayments/<int:rid>/urge-records', methods=['GET'])
def get_urge_records(rid):
    """获取某笔还款的催促记录"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM urge_records WHERE repayment_id=? ORDER BY created_at DESC", (rid,))
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

    # 前置校验：对应期数已存在 H2 催款记录
    c.execute("SELECT id FROM urge_records WHERE repayment_id=? LIMIT 1", (repayment_id,))
    if not c.fetchone() and user['role'] != '老板':
        conn.close()
        return jsonify({'success': False, 'message': '请先完成该期 H2 催款，再发起锁车申请'}), 400

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


# ======================== 旧车入库审批（退车后：销售发起→车管验车→运营查数据→财务复核→车管入库）========================
@app.route('/api/return-inspections/<int:rid>/submit-approval', methods=['POST'])
@require_role('销售', '车管')
def submit_return_approval(rid):
    """验收单提交审批"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT status FROM return_inspections WHERE id=?", (rid,))
    ri = c.fetchone()
    if not ri:
        conn.close()
        return jsonify({'success': False, 'message': '验收单不存在'}), 404
    if ri['status'] == '待车管验车':
        conn.close()
        return jsonify({'success': False, 'message': '请先由车管完成验车填单'}), 400

    c.execute("UPDATE return_inspections SET status='待审批' WHERE id=?", (rid,))
    create_approval_flow(conn, 'return_stock', rid)

    log_audit(conn, '退车审批提交', 'return_inspection', rid, '验收单提交入库审批')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已提交入库审批（运营查车辆数据→财务复核）'})


@app.route('/api/return-inspections/<int:rid>/execute-stock', methods=['POST'])
@require_role('车管')
def execute_return_stock(rid):
    """车管执行旧车入库（审批通过后）"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT vehicle_id, contract_id, status FROM return_inspections WHERE id=?", (rid,))
    ri = c.fetchone()
    if not ri or ri['status'] != '待入库':
        conn.close()
        return jsonify({'success': False, 'message': '审批未通过或状态不正确'}), 400

    vehicle_id = ri['vehicle_id']
    contract_id = ri['contract_id']

    c.execute("UPDATE return_inspections SET status='已入库' WHERE id=?", (rid,))
    if vehicle_id:
        c.execute("UPDATE vehicles SET status='在库', is_new='二手车' WHERE id=?", (vehicle_id,))
    if contract_id:
        c.execute("UPDATE contracts SET contract_status='已结清' WHERE id=?", (contract_id,))

    log_audit(conn, '旧车入库', 'vehicle', vehicle_id, '车管执行旧车入库')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '旧车已入库'})


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
    # 启动补跑一次（漏跑兜底），并启动每日 00:30 调度线程
    try:
        run_daily_collect(force=False)
    except Exception as e:
        print(f'[daily-collect] startup catch-up error: {e}')
    start_scheduler()
    app.run(port=49165, debug=True, use_reloader=False)
