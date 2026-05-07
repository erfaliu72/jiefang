from flask import Flask, jsonify, request, send_from_directory, session
from flask_cors import CORS
from database import get_db, init_db, seed_data
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from functools import wraps
from werkzeug.utils import secure_filename
import os, uuid

app = Flask(__name__)
app.secret_key = 'jinjuyuan-secret-2024'
CORS(app, supports_credentials=True)

# 上传文件目录
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ======================== PRD 角色权限矩阵 ========================
# 每个角色可访问的页面 — 车辆列表全员可见
ROLE_PAGES = {
    '老板': ['dashboard', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return', 'profit', 'settings'],
    '运营': ['dashboard', 'assets', 'approvals', 'risk', 'return'],
    '法务': ['dashboard', 'assets', 'approvals'],
    '财务': ['dashboard', 'assets', 'approvals', 'bills', 'reconciliation', 'profit', 'return'],
    '车管': ['dashboard', 'assets', 'approvals', 'return'],
    '销售': ['dashboard', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return'],
}

# 每个角色可执行的操作
ROLE_ACTIONS = {
    '老板': ['*'],
    '运营': ['view_contracts', 'view_overdue', 'lock_vehicle', 'execute_lock', 'confirm_repayment', 'initiate_return'],
    '法务': ['view_contracts', 'approve_contract'],
    '财务': ['view_contracts', 'confirm_repayment', 'confirm_factory', 'view_bills', 'view_profit', 'upload_receipt', 'collect_payment', 'verify_return', 'upload_initial_receipt'],
    '车管': ['add_vehicle', 'update_vehicle', 'activate_vehicle', 'return_inspect', 'deliver_vehicle', 'return_stock'],
    '销售': ['create_contract', 'view_contracts', 'upload_screenshot', 'view_overdue', 'initiate_return', 'request_lock', 'initiate_initial_payment'],
}

# ======================== 通用审批流程配置 ========================
APPROVAL_CONFIGS = {
    'contract_delivery': [
        {'step': 1, 'role': '运营', 'label': '运营审核'},
        {'step': 2, 'role': '财务', 'label': '财务审核'},
        {'step': 3, 'role': '法务', 'label': '法务审核'},
        {'step': 4, 'role': '老板', 'label': '老板审批'},
    ],
    'sale_payment': [
        {'step': 1, 'role': '运营', 'label': '运营发起支付核对'},
        {'step': 2, 'role': '财务', 'label': '财务核对平账'},
    ],
    'initial_payment': [
        {'step': 1, 'role': '财务', 'label': '财务审核收款'},
        {'step': 2, 'role': '老板', 'label': '老板确认'},
    ],
    'lock_request': [
        {'step': 1, 'role': '财务', 'label': '财务审核历史数据'},
        {'step': 2, 'role': '老板', 'label': '老板确认锁车'},
    ],
    'return_stock': [
        {'step': 1, 'role': '运营', 'label': '运营审批'},
        {'step': 2, 'role': '财务', 'label': '财务审批'},
        {'step': 3, 'role': '老板', 'label': '老板确认'},
    ],
}


def normalize_approval_step_label(ref_type, step_order, required_role, step_label):
    """兼容历史审批流文案，确保页面展示使用当前标准名称。"""
    if ref_type == 'initial_payment' and step_order == 2 and required_role == '老板':
        return '老板确认'
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
    if contract_type == '销售':
        return float(contract['total_price'] or contract['down_payment'] or 0)
    return float(contract['deposit'] or 0) + float(contract['down_payment'] or 0) + float(contract['rent'] or 0)


def initial_payment_label(contract):
    if not contract:
        return '首付款审核'
    return '卖车付款审核' if contract['contract_type'] == '销售' else '押金及首次支付审核'


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
    conn.close()
    if not user:
        return jsonify({'success': False, 'message': '用户名或密码错误'}), 401
    session['user_id'] = user['id']
    return jsonify({
        'success': True,
        'user': {
            'id': user['id'],
            'username': user['username'],
            'display_name': user['display_name'],
            'role': user['role'],
            'pages': ROLE_PAGES.get(user['role'], []),
            'actions': ROLE_ACTIONS.get(user['role'], []),
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
    return jsonify({
        'success': True,
        'user': {
            **user,
            'pages': ROLE_PAGES.get(user['role'], []),
            'actions': ROLE_ACTIONS.get(user['role'], []),
        }
    })



# ======================== 审计日志 ========================
def log_audit(conn, action, target_type, target_id, detail, operator='系统'):
    """PRD NFR: 所有财务状态变更必须记录操作人+原始金额+变更后金额"""
    ip = request.remote_addr if request else '127.0.0.1'
    conn.execute(
        "INSERT INTO audit_logs (action, target_type, target_id, detail, operator, ip_address) VALUES (?,?,?,?,?,?)",
        (action, target_type, target_id, detail, operator, ip)
    )


# ======================== 自动逾期检测 + T+7锁车 ========================
def check_overdue():
    """扫描所有到期未还的记录，自动标记为逾期。"""
    conn = get_db()
    c = conn.cursor()
    today = datetime.now().strftime('%Y-%m-%d')
    c.execute("UPDATE repayments SET status='逾期' WHERE status='待还款' AND due_date < ?", (today,))
    c.execute("UPDATE factory_repayments SET status='逾期' WHERE status='待还款' AND due_date < ?", (today,))

    conn.commit()
    conn.close()


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

    # 客户逾期笔数
    c.execute("SELECT COUNT(*) as cnt FROM repayments WHERE status = '逾期'")
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
    })


# ======================== 车辆资产 CRUD ========================
@app.route('/api/vehicles', methods=['GET'])
def get_vehicles():
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

@app.route('/api/car-types', methods=['GET'])
def get_car_types():
    return jsonify(CAR_TYPE_PRESETS)


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
    data = request.json
    conn = get_db()
    c = conn.cursor()
    try:
        c.execute('''
        INSERT INTO vehicles (vin, plate_number, company, car_type, is_new, invoice_date,
                              invoice_price, purchase_price, tax_rate, estimated_residual_value,
                              guidance_price, invoice_contract_file, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data.get('vin'), data.get('plate_number'), data.get('company', '陕西金聚源汽车服务有限公司'),
            data.get('car_type'), data.get('is_new', '新车'), data.get('invoice_date'),
            data.get('invoice_price', 0), data.get('purchase_price', 0), data.get('tax_rate', 0.13),
            data.get('estimated_residual_value', 0), data.get('guidance_price', 0),
            data.get('invoice_contract_file', ''),
            data.get('status', '在库')
        ))
        vehicle_id = c.lastrowid
        conn.commit()
        return jsonify({'success': True, 'id': vehicle_id, 'message': '车辆入库成功'})
    except Exception as e:
        conn.rollback()
        return jsonify({'success': False, 'message': str(e)}), 400
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
                'invoice_price', 'purchase_price', 'tax_rate', 'guidance_price', 'invoice_contract_file', 'status']:
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
    data = request.json
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE vehicles SET guidance_price = ? WHERE id = ?", (data.get('price'), vid))
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '指导价更新成功'})

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
    conn.close()
    return jsonify(contracts)

@app.route('/api/contracts', methods=['POST'])
def add_contract():
    data = request.json
    conn = get_db()
    c = conn.cursor()
    try:
        vehicle_id = data['vehicle_id']

        # ===== 校验：同一辆车不能重复签约 =====
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
        c.execute("SELECT guidance_price, invoice_price FROM vehicles WHERE id=?", (vehicle_id,))
        vrow = c.fetchone()
        snap_guidance = vrow['guidance_price'] if vrow else 0
        snap_invoice = vrow['invoice_price'] if vrow else 0

        # 所有合同进入审批流程，不再直接结清/出库
        contract_status = '执行中'

        c.execute('''
        INSERT INTO contracts (vehicle_id, customer_id, contract_type, business_mode, rental_method, repayment_day,
                               start_date, end_date, total_price, customer_loan_amount, loan_amount, monthly_payment,
                               rent, loan_periods, company, yard, factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
                               down_payment_status, deposit_status, delivery_status,
                               contract_status,
                               snapshot_guidance_price, snapshot_invoice_price, contract_file)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            vehicle_id, customer_id, contract_type, data.get('business_mode', '转租'),
            data.get('rental_method', '经营租赁'), repayment_day,
            start_date, end_dt.strftime('%Y-%m-%d'),
            data.get('total_price', 0), customer_loan_amount, data.get('loan_amount', 0),
            monthly_payment, rent, loan_periods, data.get('company', ''), data.get('yard', ''), factory_guarantee_deposit, factory_repayment_months, factory_periods, deposit, down_payment,
            '免收' if down_payment == 0 else '待收',
            '免收' if deposit == 0 else '待收',
            '待审批',
            contract_status,
            snap_guidance, snap_invoice,
            data.get('contract_file', '')
        ))
        contract_id = c.lastrowid

        # 销售合同不生成还款计划
        if contract_type != '销售':
            for p in range(1, max(loan_periods, factory_periods) + 1):
                try:
                    due_dt = start_dt + relativedelta(months=p)
                    due_dt = due_dt.replace(day=min(repayment_day, 28))
                except Exception:
                    due_dt = start_dt + timedelta(days=30 * p)
                due_str = due_dt.strftime('%Y-%m-%d')

                if rent > 0 and p <= loan_periods:
                    c.execute("INSERT INTO repayments (contract_id, period, due_date, amount) VALUES (?, ?, ?, ?)",
                              (contract_id, p, due_str, rent))
                if monthly_payment > 0 and p <= factory_periods:
                    c.execute("INSERT INTO factory_repayments (contract_id, period, due_date, amount) VALUES (?, ?, ?, ?)",
                              (contract_id, p, due_str, monthly_payment))

        # 车辆状态暂不改变，待出库审批通过后再更新
        # 创建出库审批流程（运营→财务→老板）
        create_approval_flow(conn, 'contract_delivery', contract_id)

        log_audit(conn, '创建合同', 'contract', contract_id,
                  f'类型{contract_type} 车辆{vehicle_id} 月租{rent} 月供{monthly_payment} 期数{loan_periods} 进入审批流程')
        conn.commit()
        return jsonify({'success': True, 'id': contract_id, 'message': '合同创建成功，已进入审批流程'})
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
    conn.close()
    return jsonify(repayments)

@app.route('/api/repayments/<int:rid>/confirm', methods=['POST'])
@require_role('财务')
def confirm_repayment(rid):
    """PRD: 财务确认后，系统自动核减余额"""
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

    c.execute("UPDATE repayments SET status='已还款', paid_at=? WHERE id=?",
              (datetime.now().strftime('%Y-%m-%d'), rid))
    # 联动更新合同已收租金
    c.execute("UPDATE contracts SET collected_rent = collected_rent + ? WHERE id=?",
              (amount, contract_id))
    log_audit(conn, '客户还款核销', 'repayment', rid,
             f'合同{contract_id} 金额{amount} 原状态{old_status}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '核销成功'})


# ======================== 厂家还款（公司 → 一汽解放）========================
@app.route('/api/contracts/<int:cid>/factory-repayments', methods=['GET'])
def get_factory_repayments(cid):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM factory_repayments WHERE contract_id = ? ORDER BY period ASC", (cid,))
    repayments = [dict(row) for row in c.fetchall()]
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


# ======================== 利润核算 ========================
@app.route('/api/profit/by-vehicle', methods=['GET'])
def get_profit_by_vehicle():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT v.id, v.vin, v.plate_number, v.car_type, v.invoice_price,
               c.rent, c.monthly_payment, c.loan_periods, c.deposit,
               c.collected_rent, c.collected_deposit, c.paid_principal,
               c.contract_status, c.business_mode,
               COALESCE((SELECT SUM(amount) FROM repayments WHERE contract_id=c.id AND status='已还款'), 0) as customer_received,
               COALESCE((SELECT SUM(amount) FROM factory_repayments WHERE contract_id=c.id AND status='已还款'), 0) as factory_paid
        FROM vehicles v
        JOIN contracts c ON c.vehicle_id = v.id
        ORDER BY v.id ASC
    """)
    rows = []
    for row in c.fetchall():
        d = dict(row)
        d['profit'] = round(d['customer_received'] - d['factory_paid'], 2)
        d['monthly_spread'] = round((d['rent'] or 0) - (d['monthly_payment'] or 0), 2)
        rows.append(d)
    conn.close()
    return jsonify(rows)


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
        WHERE r.status = '逾期'
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
        WHERE r.status IN ('待还款', '逾期')
        ORDER BY r.due_date ASC
    """)
    bills = [dict(row) for row in c.fetchall()]
    conn.close()
    return jsonify(bills)


# ======================== 对账核销流程（三位一体）========================
@app.route('/api/reconciliation/list', methods=['GET'])
def get_reconciliation_list():
    """获取所有需要对账的还款记录（含已核销的），用于对账单页面"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT r.id, r.contract_id, r.period, r.due_date, r.amount, r.status,
               r.paid_at, r.screenshot_path, r.bank_receipt_path, r.bank_serial,
               r.verified_by, r.verified_at,
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


@app.route('/api/reconciliation/<int:rid>/screenshot', methods=['POST'])
@require_role('销售')
def upload_screenshot(rid):
    """步骤1：销售上传客户付款截图"""
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

    # 更新流水号并自动核销
    c.execute("""UPDATE repayments SET bank_serial=?, verified_by=?, verified_at=?,
                 status='已还款', paid_at=? WHERE id=?""",
              (bank_serial, user, now, datetime.now().strftime('%Y-%m-%d'), rid))
    # 联动更新合同已收租金
    c.execute("UPDATE contracts SET collected_rent = collected_rent + ? WHERE id=?",
              (amount, contract_id))
    log_audit(conn, '对账核销', 'repayment', rid,
             f'合同{contract_id} 金额{amount} 流水号{bank_serial} 核销人{user}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'核销成功，流水号: {bank_serial}'})


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
    conn = get_db()
    c = conn.cursor()
    if pay_type == 'down_payment':
        c.execute("UPDATE contracts SET down_payment_status='已收' WHERE id=?", (cid,))
        log_audit(conn, '收取首付', 'contract', cid, '财务确认收取首付')
    else:
        c.execute("SELECT deposit FROM contracts WHERE id=?", (cid,))
        row = c.fetchone()
        amount = row['deposit'] if row else 0
        c.execute("UPDATE contracts SET deposit_status='已收', collected_deposit=? WHERE id=?", (amount, cid))
        log_audit(conn, '收取押金', 'contract', cid, f'财务确认收取押金 {amount}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '收款确认成功'})


# ======================== 合同首次付款（销售发起 → 财务收款 → 老板确认）========================
@app.route('/api/contracts/<int:cid>/initial-payment', methods=['POST'])
@require_role('销售')
def create_initial_payment(cid):
    """合同审批完成后，销售发起首次付款审核。"""
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
                  f'合同{cid} {payment_type} 金额{amount} 销售{user["display_name"]}', user['display_name'])
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

    c.execute("UPDATE contract_initial_payments SET bank_receipt_path=? WHERE id=?", (receipt_path, pid))
    log_audit(conn, '上传首次付款回单', 'initial_payment', pid,
              f'财务上传公司收款回单 {receipt_path}', request.current_user['display_name'])
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '公司收款回单已上传'})


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
    c.execute("SELECT * FROM return_inspections ORDER BY id DESC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)


@app.route('/api/return-inspections', methods=['POST'])
@require_role('销售', '车管')
def create_return_inspection():
    data = request.json
    conn = get_db()
    c = conn.cursor()
    user = get_current_user()

    # 自动填充车辆信息
    vehicle_id = data.get('vehicle_id')
    if vehicle_id:
        c.execute("SELECT v.*, c.id as cid, c.deposit, c.start_date, c.end_date, cu.name as cust_name FROM vehicles v LEFT JOIN contracts c ON c.vehicle_id=v.id LEFT JOIN customers cu ON cu.id=c.customer_id WHERE v.id=?", (vehicle_id,))
        vrow = c.fetchone()
        if vrow:
            if not data.get('plate_number'): data['plate_number'] = vrow['plate_number']
            if not data.get('vin'): data['vin'] = vrow['vin']
            if not data.get('car_type'): data['car_type'] = vrow['car_type']
            if not data.get('company'): data['company'] = vrow['company']
            if not data.get('customer_name'): data['customer_name'] = vrow['cust_name'] or ''
            if not data.get('contract_id'): data['contract_id'] = vrow['cid']
            if not data.get('rental_period') and vrow['start_date'] and vrow['end_date']:
                data['rental_period'] = f"{vrow['start_date']} ~ {vrow['end_date']}"

    c.execute("""INSERT INTO return_inspections
        (vehicle_id, contract_id, plate_number, customer_name, rental_period, vin, car_type, company,
         return_reason, tool_triangle, tool_vest, tool_extinguisher, tool_wedge, tool_jack,
         doc_license, doc_keys, mileage, body_tire_clean,
         accident_info, insurance_surcharge, violation_info, etc_info, maintenance_info,
         rent_late_fee, return_late_fee, deposit_rent_receivable, deposit_paid,
         total_deduction, actual_refund, remark, status, sales_advisor, created_by)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (vehicle_id, data.get('contract_id'),
         data.get('plate_number',''), data.get('customer_name',''), data.get('rental_period',''),
         data.get('vin',''), data.get('car_type',''), data.get('company',''),
         data.get('return_reason','到期退车'),
         1 if data.get('tool_triangle') else 0, 1 if data.get('tool_vest') else 0,
         1 if data.get('tool_extinguisher') else 0, 1 if data.get('tool_wedge') else 0,
         1 if data.get('tool_jack') else 0,
         1 if data.get('doc_license') else 0, 1 if data.get('doc_keys') else 0,
         data.get('mileage',''), data.get('body_tire_clean',''),
         data.get('accident_info',''), data.get('insurance_surcharge',''),
         data.get('violation_info',''), data.get('etc_info',''), data.get('maintenance_info',''),
         data.get('rent_late_fee',0), data.get('return_late_fee',0),
         data.get('deposit_rent_receivable',0), data.get('deposit_paid',0),
         data.get('total_deduction',0), data.get('actual_refund',0),
         data.get('remark',''), '已登记', data.get('sales_advisor',''),
         user['display_name'] if user else ''))

    rid = c.lastrowid

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
    fields = []
    values = []
    for key in ['plate_number','customer_name','rental_period','vin','car_type','company',
                'return_reason','tool_triangle','tool_vest','tool_extinguisher','tool_wedge','tool_jack',
                'doc_license','doc_keys','mileage','body_tire_clean',
                'accident_info','insurance_surcharge','violation_info','etc_info','maintenance_info',
                'rent_late_fee','return_late_fee','deposit_rent_receivable','deposit_paid',
                'total_deduction','actual_refund','remark','status','sales_advisor','sales_manager','general_manager']:
        if key in data:
            fields.append(f"{key}=?")
            values.append(data[key])
    if fields:
        values.append(rid)
        c.execute(f"UPDATE return_inspections SET {','.join(fields)} WHERE id=?", values)
        conn.commit()
    conn.close()
    return jsonify({'success': True})


# ======================== 通用审批流程 API ========================
@app.route('/api/approvals', methods=['GET'])
@login_required
def get_approvals():
    """获取审批列表，支持按 ref_type 和 role 过滤"""
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
            'steps': [],
            'current_step': 0,
            'overall_status': '已完成',
        }

        if base_ref_type in ('contract_delivery', 'sale_payment'):
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
                'delivery_photo_path': row.get('delivery_photo_path', ''),
                'delivery_document_path': row.get('delivery_document_path', ''),
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

    if user['role'] != flow['required_role'] and user['role'] != '老板':
        conn.close()
        return jsonify({'success': False, 'message': f'需要{flow["required_role"]}角色审批'}), 403

    # 首次付款的财务审核必须先上传公司到账/银行回单。
    if flow['ref_type'] == 'initial_payment' and flow['required_role'] == '财务':
        c.execute("SELECT bank_receipt_path FROM contract_initial_payments WHERE id=?", (flow['ref_id'],))
        payment = c.fetchone()
        if not payment or not payment['bank_receipt_path']:
            conn.close()
            return jsonify({'success': False, 'message': '请先上传公司收款回单，再进行财务审核'}), 400

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
        if ref_type == 'contract_delivery':
            # 流程图要求：销售发起的合同（租赁/销售/以租代售）审批全走完后，
            # 不能直接出库，必须先由销售发起首次付款。
            c.execute("UPDATE contracts SET delivery_status='待首付款' WHERE id=?", (ref_id,))
            message = '合同审批通过，请销售发起首次付款'
        elif ref_type == 'sale_payment':
            # 支付核对完成 → 待出库
            c.execute("UPDATE contracts SET delivery_status='待出库' WHERE id=?", (ref_id,))
            message = '支付核对完成，等待车管出库'
        elif ref_type == 'initial_payment':
            c.execute("""
                SELECT ip.*, c.contract_type, c.deposit, c.down_payment, c.rent
                FROM contract_initial_payments ip
                JOIN contracts c ON c.id = ip.contract_id
                WHERE ip.id=?
            """, (ref_id,))
            payment = c.fetchone()
            if payment:
                contract_id = payment['contract_id']
                updates = ["delivery_status='待出库'"]
                params = []
                if (payment['deposit'] or 0) > 0:
                    updates.append("deposit_status='已收'")
                    updates.append("collected_deposit=?")
                    params.append(payment['deposit'])
                if (payment['down_payment'] or 0) > 0 or payment['contract_type'] == '销售':
                    updates.append("down_payment_status='已收'")
                if payment['contract_type'] != '销售' and (payment['rent'] or 0) > 0:
                    updates.append("collected_rent=COALESCE(collected_rent,0)+?")
                    params.append(payment['rent'])
                params.append(contract_id)
                c.execute(f"UPDATE contracts SET {', '.join(updates)} WHERE id=?", params)
                if payment['contract_type'] != '销售' and (payment['rent'] or 0) > 0:
                    # 首次付款包含“首次支付金额/首期租金”时，直接核销第 1 期客户还款，
                    # 避免后续对账再核销同一笔租金造成重复入账。
                    c.execute("""
                        UPDATE repayments
                        SET status='已还款',
                            paid_at=?,
                            screenshot_path=COALESCE(screenshot_path, ?),
                            bank_receipt_path=COALESCE(bank_receipt_path, ?),
                            bank_serial=COALESCE(bank_serial, ?),
                            verified_by=COALESCE(verified_by, ?),
                            verified_at=COALESCE(verified_at, ?),
                            remark=COALESCE(remark, '首次付款审核自动核销首期租金')
                        WHERE id=(
                            SELECT id FROM repayments
                            WHERE contract_id=? AND status!='已还款'
                            ORDER BY period ASC
                            LIMIT 1
                        )
                    """, (
                        datetime.now().strftime('%Y-%m-%d'),
                        payment['customer_screenshot_path'],
                        payment['bank_receipt_path'],
                        f"INITIAL-{ref_id}",
                        user['display_name'],
                        now,
                        contract_id,
                    ))
                c.execute("UPDATE contract_initial_payments SET status='已通过', approved_by=?, approved_at=? WHERE id=?",
                          (user['display_name'], now, ref_id))
            message = '首次付款审核完成，等待车管出库'
        elif ref_type == 'lock_request':
            c.execute("UPDATE lock_requests SET status='已通过', approved_by=?, approved_at=? WHERE id=?",
                      (user['display_name'], now, ref_id))
            c.execute("SELECT vehicle_id FROM lock_requests WHERE id=?", (ref_id,))
            lr = c.fetchone()
            if lr:
                c.execute("UPDATE vehicles SET status='已锁车' WHERE id=?", (lr['vehicle_id'],))
            message = '锁车审批通过，车辆已锁'
        elif ref_type == 'return_stock':
            # 旧车入库审批通过 → 待车管入库
            c.execute("SELECT vehicle_id FROM return_inspections WHERE id=?", (ref_id,))
            ri = c.fetchone()
            if ri:
                c.execute("UPDATE return_inspections SET status='待入库' WHERE id=?", (ref_id,))
            message = '退车审批通过，等待车管入库'
    else:
        # 更新合同状态显示当前审批进度
        if ref_type == 'contract_delivery':
            c.execute("UPDATE contracts SET delivery_status='审批中' WHERE id=?", (ref_id,))
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
    if user['role'] != flow['required_role'] and user['role'] != '老板':
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
    if ref_type in ('contract_delivery', 'sale_payment'):
        c.execute("UPDATE contracts SET delivery_status='已驳回' WHERE id=?", (ref_id,))
    elif ref_type == 'initial_payment':
        c.execute("UPDATE contract_initial_payments SET status='已驳回' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE contracts
            SET delivery_status='首付已驳回'
            WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
        """, (ref_id,))
    elif ref_type == 'lock_request':
        c.execute("UPDATE lock_requests SET status='已驳回' WHERE id=?", (ref_id,))
    elif ref_type == 'return_stock':
        c.execute("UPDATE return_inspections SET status='已驳回' WHERE id=?", (ref_id,))

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

    if ref_type in ('contract_delivery', 'sale_payment'):
        c.execute("UPDATE contracts SET delivery_status='待审批' WHERE id=?", (ref_id,))
    elif ref_type == 'initial_payment':
        c.execute("UPDATE contract_initial_payments SET status='审批中' WHERE id=?", (ref_id,))
        c.execute("""
            UPDATE contracts
            SET delivery_status='首付审批中'
            WHERE id=(SELECT contract_id FROM contract_initial_payments WHERE id=?)
        """, (ref_id,))
    elif ref_type == 'return_stock':
        c.execute("UPDATE return_inspections SET status='待审批' WHERE id=?", (ref_id,))

    create_approval_flow(conn, ref_type, ref_id)
    log_audit(conn, '重新提交审批', ref_type, ref_id, '驳回后重新提交')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已重新提交审批'})


@app.route('/api/contracts/<int:cid>/approval-status', methods=['GET'])
def get_contract_approval_status(cid):
    """获取合同的审批状态（合同审批+首次付款/兼容旧支付核对）"""
    conn = get_db()
    delivery = get_approval_status(conn, 'contract_delivery', cid)
    payment = get_approval_status(conn, 'sale_payment', cid)
    c = conn.cursor()
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
    """T+1 运营催促 / T+3 销售催促"""
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

    # 确定催促类型
    if overdue_days >= 1 and user['role'] in ('运营', '老板'):
        urge_type = 'T+1_运营'
        urge_day = 1
    elif overdue_days >= 3 and user['role'] in ('销售', '老板'):
        urge_type = 'T+3_销售'
        urge_day = 3
    else:
        conn.close()
        return jsonify({'success': False, 'message': f'当前逾期{overdue_days}天，不满足催促条件或角色不匹配'}), 400

    c.execute("""INSERT INTO urge_records (repayment_id, contract_id, vehicle_id, urge_type, urge_day, operator_id, operator_name)
                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
              (rid, row['cid'], row['vehicle_id'], urge_type, urge_day, user['id'], user['display_name']))

    log_audit(conn, '催促还款', 'repayment', rid,
              f'{user["display_name"]}({urge_type}) 逾期{overdue_days}天')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': f'{urge_type} 催促已记录'})


@app.route('/api/repayments/<int:rid>/urge-records', methods=['GET'])
def get_urge_records(rid):
    """获取某笔还款的催促记录"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM urge_records WHERE repayment_id=? ORDER BY created_at DESC", (rid,))
    records = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(records)


# ======================== 锁车审批链（销售T+5发起→财务→老板）========================
@app.route('/api/lock-requests', methods=['GET'])
@login_required
def get_lock_requests():
    """获取锁车申请列表"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""SELECT lr.*, v.plate_number, v.car_type, cu.name as customer_name
                 FROM lock_requests lr
                 JOIN vehicles v ON v.id = lr.vehicle_id
                 LEFT JOIN contracts c ON c.id = lr.contract_id
                 LEFT JOIN customers cu ON cu.id = c.customer_id
                 ORDER BY lr.id DESC""")
    rows = [dict(r) for r in c.fetchall()]
    # 附加审批状态
    for r in rows:
        status = get_approval_status(conn, 'lock_request', r['id'])
        r['approval_steps'] = status['steps']
        r['approval_status'] = status['status']
        r['approval_current_step'] = status['current_step']
    conn.close()
    return jsonify(rows)


@app.route('/api/lock-requests', methods=['POST'])
@require_role('销售')
def create_lock_request():
    """销售 T+5 发起锁车申请"""
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
    if overdue_days < 5 and user['role'] != '老板':
        conn.close()
        return jsonify({'success': False, 'message': 'T+5 之后才能发起锁车申请'}), 400

    c.execute("""
        SELECT id FROM lock_requests
        WHERE repayment_id=? AND status IN ('待审批', '已通过')
        ORDER BY id DESC LIMIT 1
    """, (repayment_id,))
    if c.fetchone():
        conn.close()
        return jsonify({'success': False, 'message': '该笔逾期已存在有效锁车申请'}), 400

    c.execute("""INSERT INTO lock_requests (vehicle_id, contract_id, repayment_id, reason, overdue_days, requested_by)
                 VALUES (?, ?, ?, ?, ?, ?)""",
              (repayment['vehicle_id'], repayment['contract_id'], repayment_id,
               data.get('reason', '逾期锁车'), overdue_days, user['display_name']))
    lr_id = c.lastrowid

    # 创建锁车审批流（财务→老板）
    create_approval_flow(conn, 'lock_request', lr_id)

    log_audit(conn, '锁车申请', 'vehicle', repayment['vehicle_id'],
              f'逾期{overdue_days}天 发起人:{user["display_name"]}')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '锁车申请已提交，等待财务→老板审批'})


# ======================== 旧车入库审批（退车后：运营→财务→老板→车管）========================
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

    c.execute("UPDATE return_inspections SET status='待审批' WHERE id=?", (rid,))
    create_approval_flow(conn, 'return_stock', rid)

    log_audit(conn, '退车审批提交', 'return_inspection', rid, '验收单提交入库审批')
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': '已提交入库审批（运营→财务→老板）'})


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


# ======================== 启动 ========================
if __name__ == '__main__':
    init_db()
    seed_data()
    app.run(port=49165, debug=True)
