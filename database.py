import sqlite3
import os
from datetime import datetime, timedelta

DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jinjuyuan.db')

def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

# ================================================================
#  安全建表 — 仅在表不存在时创建，不会删除已有数据
# ================================================================
def init_db():
    conn = get_db()
    c = conn.cursor()

    # ====== 车辆资产表 ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS vehicles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vin TEXT UNIQUE NOT NULL,
        plate_number TEXT,
        company TEXT,
        car_type TEXT,
        is_new TEXT DEFAULT '新车',
        invoice_date TEXT,
        invoice_price REAL DEFAULT 0,
        purchase_price REAL DEFAULT 0,
        tax_rate REAL DEFAULT 0.13,
        estimated_residual_value REAL DEFAULT 0,
        guidance_price REAL DEFAULT 0,
        depreciation_months INTEGER DEFAULT 60,
        invoice_contract_file TEXT,
        status TEXT DEFAULT '在库',
        activated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # ====== 客户表 ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        id_card TEXT,
        address TEXT,
        remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # ====== 合同表（增加价格快照字段，PRD要求冗余存储计价因子） ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS contracts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        customer_id INTEGER,
        contract_type TEXT DEFAULT '租赁',
        business_mode TEXT DEFAULT '转租',
        rental_method TEXT,
        repayment_day INTEGER DEFAULT 1,
        start_date TEXT,
        end_date TEXT,
        total_price REAL DEFAULT 0,
        loan_amount REAL DEFAULT 0,
        monthly_payment REAL DEFAULT 0,
        rent REAL DEFAULT 0,
        loan_periods INTEGER DEFAULT 0,
        deposit REAL DEFAULT 0,
        down_payment REAL DEFAULT 0,
        down_payment_status TEXT DEFAULT '待收',
        deposit_status TEXT DEFAULT '待收',
        delivery_status TEXT DEFAULT '待出库',
        delivery_date TEXT,
        paid_principal REAL DEFAULT 0,
        loan_balance TEXT DEFAULT '0',
        collected_deposit REAL DEFAULT 0,
        collected_rent REAL DEFAULT 0,
        contract_status TEXT DEFAULT '执行中',
        -- PRD 价格快照：成交时锁定的参考价格，后续调价不影响
        snapshot_guidance_price REAL DEFAULT 0,
        snapshot_invoice_price REAL DEFAULT 0,
        contract_file TEXT,
        remark TEXT,
        loan_remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id),
        FOREIGN KEY (customer_id) REFERENCES customers (id)
    )
    ''')

    # ====== 客户还款计划（客户 → 公司）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS repayments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        period INTEGER NOT NULL,
        due_date TEXT,
        amount REAL DEFAULT 0,
        status TEXT DEFAULT '待还款',
        paid_at TEXT,
        screenshot_path TEXT,
        bank_receipt_path TEXT,
        bank_serial TEXT,
        verified_by TEXT,
        verified_at TEXT,
        remark TEXT,
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 厂家还款计划（公司 → 一汽解放）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS factory_repayments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        period INTEGER NOT NULL,
        due_date TEXT,
        amount REAL DEFAULT 0,
        status TEXT DEFAULT '待还款',
        paid_at TEXT,
        remark TEXT,
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 附件表（通用文件上传）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS attachments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ref_type TEXT NOT NULL,
        ref_id INTEGER NOT NULL,
        file_type TEXT,
        file_path TEXT NOT NULL,
        uploaded_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # ====== 审计日志表（PRD NFR: 所有财务状态变更必须记录）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS audit_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT NOT NULL,
        target_type TEXT,
        target_id INTEGER,
        detail TEXT,
        operator TEXT DEFAULT '系统',
        ip_address TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # ====== 锁车申请表（PRD 锁车审批流程）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS lock_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        contract_id INTEGER,
        repayment_id INTEGER,
        reason TEXT,
        overdue_days INTEGER DEFAULT 0,
        status TEXT DEFAULT '待审批',
        requested_by TEXT,
        approved_by TEXT,
        approved_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id)
    )
    ''')

    # ====== 通用审批流程表（出库审批 / 卖车支付核对 / 锁车审批 / 旧车入库审批）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS approval_flows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ref_type TEXT NOT NULL,
        ref_id INTEGER NOT NULL,
        batch_no TEXT NOT NULL,
        step_order INTEGER NOT NULL,
        required_role TEXT NOT NULL,
        step_label TEXT,
        status TEXT DEFAULT '待审批',
        operator_id INTEGER,
        operator_name TEXT,
        comment TEXT,
        acted_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # ====== 催促记录表（T+1运营催促 / T+3销售催促）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS urge_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        repayment_id INTEGER NOT NULL,
        contract_id INTEGER NOT NULL,
        vehicle_id INTEGER NOT NULL,
        urge_type TEXT NOT NULL,
        urge_day INTEGER NOT NULL,
        status TEXT DEFAULT '已催促',
        result TEXT DEFAULT '待跟进',
        operator_id INTEGER,
        operator_name TEXT,
        remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (repayment_id) REFERENCES repayments (id)
    )
    ''')

    # ====== 退还车辆验收单 ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS return_inspections (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER,
        contract_id INTEGER,
        -- 表头信息
        plate_number TEXT,
        customer_name TEXT,
        rental_period TEXT,
        vin TEXT,
        car_type TEXT,
        company TEXT,
        -- 退租原因: 到期退车/提前退车/租赁转购车/临时用车
        return_reason TEXT DEFAULT '到期退车',
        -- 随车工具检查 (1=有 0=无)
        tool_triangle INTEGER DEFAULT 0,
        tool_vest INTEGER DEFAULT 0,
        tool_extinguisher INTEGER DEFAULT 0,
        tool_wedge INTEGER DEFAULT 0,
        tool_jack INTEGER DEFAULT 0,
        -- 证件资料
        doc_license INTEGER DEFAULT 0,
        doc_keys INTEGER DEFAULT 0,
        -- 公里数
        mileage TEXT,
        body_tire_clean TEXT,
        -- 事故
        accident_info TEXT,
        insurance_surcharge TEXT,
        -- 违章
        violation_info TEXT,
        -- ETC
        etc_info TEXT,
        -- 维修保养
        maintenance_info TEXT,
        -- 押金支付情况
        rent_late_fee REAL DEFAULT 0,
        return_late_fee REAL DEFAULT 0,
        deposit_rent_receivable REAL DEFAULT 0,
        deposit_paid REAL DEFAULT 0,
        total_deduction REAL DEFAULT 0,
        actual_refund REAL DEFAULT 0,
        -- 备注
        remark TEXT,
        -- 流程状态
        status TEXT DEFAULT '待登记',
        sales_advisor TEXT,
        sales_manager TEXT,
        general_manager TEXT,
        created_by TEXT,
        inspected_by TEXT,
        inspected_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 用户表（角色鉴权）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        display_name TEXT NOT NULL,
        role TEXT NOT NULL,
        is_active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    # === 安全添加新列（如果表已存在但缺少新字段）===
    safe_alter_columns = [
        ("contracts", "snapshot_guidance_price", "REAL DEFAULT 0"),
        ("contracts", "snapshot_invoice_price", "REAL DEFAULT 0"),
        ("contracts", "contract_type", "TEXT DEFAULT '租赁'"),
        ("contracts", "down_payment", "REAL DEFAULT 0"),
        ("contracts", "down_payment_status", "TEXT DEFAULT '待收'"),
        ("contracts", "deposit_status", "TEXT DEFAULT '待收'"),
        ("contracts", "delivery_status", "TEXT DEFAULT '待出库'"),
        ("contracts", "delivery_date", "TEXT"),
        ("vehicles", "customer_name", "TEXT"),
        ("vehicles", "invoice_contract_file", "TEXT"),
        ("vehicles", "purchase_price", "REAL DEFAULT 0"),
        ("vehicles", "tax_rate", "REAL DEFAULT 0.13"),
        ("vehicles", "estimated_residual_value", "REAL DEFAULT 0"),
        ("vehicles", "guidance_price", "REAL DEFAULT 0"),
        ("vehicles", "depreciation_months", "INTEGER DEFAULT 60"),
        ("contracts", "contract_file", "TEXT"),
        ("repayments", "screenshot_path", "TEXT"),
        ("repayments", "bank_receipt_path", "TEXT"),
        ("repayments", "bank_serial", "TEXT"),
        ("repayments", "verified_by", "TEXT"),
        ("repayments", "verified_at", "TEXT"),
    ]
    for table, col, col_type in safe_alter_columns:
        try:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
        except Exception:
            pass

    conn.commit()
    conn.close()
    print(f"Database initialized: {DATABASE}")


# ================================================================
#  种子数据 — 仅在表为空时插入，幂等安全
# ================================================================
def seed_data():
    conn = get_db()
    c = conn.cursor()

    # ====== 默认用户（幂等）======
    c.execute("SELECT COUNT(*) as cnt FROM users")
    if c.fetchone()['cnt'] == 0:
        default_users = [
            ('boss',   '123456', '王老板', '老板'),
            ('ops',    '123456', '李运营', '运营'),
            ('fin',    '123456', '张财务', '财务'),
            ('fleet',  '123456', '赵车管', '车管'),
            ('sales',  '123456', '周销售', '销售'),
        ]
        for u in default_users:
            c.execute("INSERT INTO users (username, password, display_name, role) VALUES (?,?,?,?)", u)
        conn.commit()
        print("Default users created.")

    c.execute("SELECT COUNT(*) as cnt FROM vehicles")
    if c.fetchone()['cnt'] > 0:
        conn.close()
        print("Seed data already exists, skipping.")
        return

    # ====== 9 台真实车辆 ======
    real_vehicles = [
        ('LFNA4LDA1NAE08565', '陕ADU8101', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2022-12-28', 99900),
        ('LFNA4LDA7NAE08442', '陕AAY1263', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2022-12-28', 99900),
        ('LFNA4LDA8NAE08563', '陕AA06286', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-01-04', 99900),
        ('LFNA4LDA9NAE08443', '陕ADY4890', '金聚源挂靠玛特汇',           'J6F 81度电厢货', '二手车', '2023-01-04', 99900),
        ('LFNA4LDA4NAE08978', '陕AA85625', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-01-17', 157000),
        ('LFNA4LDA3NAE08972', '陕AA00662', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-01-17', 157000),
        ('LFNA4LDA7NAE08974', '陕AA10855', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-01-17', 157000),
        ('LFNA4LDA8PAE19629', '陕AA28012', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-06-05', 139000),
        ('LFNA4LDA4PAE19630', '陕AA25685', '陕西金聚源汽车服务有限公司', 'J6F 81度电厢货', '新车', '2023-06-06', 139000),
    ]
    statuses = ['已结清', '已结清', '经营租赁', '经营租赁', '经营租赁', '经营租赁', '经营租赁', '经营租赁', '经营租赁']

    # 客户种子数据
    customers = [
        ('张三丰', '13800000001', '610102199001010011', '西安市未央区'),
        ('李四光', '13800000002', '610102199002020022', '西安市雁塔区'),
        ('王五常', '13800000003', '610102199003030033', '西安市碑林区'),
        ('赵六合', '13800000004', '610102199004040044', '西安市新城区'),
        ('钱七星', '13800000005', '610102199005050055', '西安市莲湖区'),
        ('孙八斗', '13800000006', '610102199006060066', '西安市灞桥区'),
        ('周九天', '13800000007', '610102199007070077', '西安市长安区'),
    ]
    for cu in customers:
        c.execute("INSERT INTO customers (name, phone, id_card, address) VALUES (?, ?, ?, ?)", cu)

    for i, v in enumerate(real_vehicles):
        invoice_date = datetime.strptime(v[5], '%Y-%m-%d')
        months_passed = (datetime.now() - invoice_date).days // 30
        residual_rate = max(0.3, 1.0 - months_passed * 0.01)
        residual = round(v[6] * residual_rate, 2)
        guidance = round(residual * 1.05, 2)

        c.execute('''
        INSERT INTO vehicles (vin, plate_number, company, car_type, is_new, invoice_date,
                              invoice_price, purchase_price, estimated_residual_value, guidance_price, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[6], residual, guidance, statuses[i]))

    # ====== 9 份真实合同（增加 customer_id 关联）======
    # (vehicle_id, rental_method, total_price, loan_amount, monthly_payment, rent, loan_periods, deposit,
    #  paid_principal, loan_balance, collected_deposit, collected_rent, customer_id)
    real_contracts = [
        (1, '经营租赁', 99900, 150000, 6648.09, 4000, 24, 1000, 150000, '已结清', 11000, 67000, 1),
        (2, '经营租赁', 99900, 150000, 6648.09, 3600, 24, 1000, 150000, '已结清', 7000, 85980, 2),
        (3, '经营租赁', 99900, 142999, 2610.12, 3200, 60, 7149.95, 73031.53, '62817.52', 6000, 81000, 3),
        (4, '经营租赁', 99900, 125000, 3871.08, 3800, 36, 6250, 80309.19, '38440.81', 8000, 69167, 4),
        (5, '经营租赁', 157000, 116142.35, 3533.28, 3500, 36, 0, 112626.83, '3515.52', 0, 76115, 5),
        (6, '经营租赁', 157000, 116142.35, 3533.28, 3500, 36, 0, 109128.62, '7013.73', 10000, 82666.67, 5),
        (7, '经营租赁', 157000, 148268.70, 4510.62, 3800, 36, 0, 139314.65, '8954.05', 5000, 72926.40, 6),
        (8, '经营租赁', 139000, 179281, 3281.06, 3500, 60, 8964.05, 91440.83, '78876.12', 12260, 99165, 7),
        (9, '经营租赁', 139000, 179281, 3281.06, 4000, 60, 8964.05, 91440.83, '78876.12', 0, 106147, 7),
    ]

    for ct in real_contracts:
        contract_status = '已结清' if ct[9] == '已结清' else '执行中'
        # 获取车辆开票日期作为合同起始日
        c.execute("SELECT invoice_date, guidance_price, invoice_price FROM vehicles WHERE id = ?", (ct[0],))
        vrow = c.fetchone()
        start_date_str = vrow['invoice_date']
        start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_dt = start_dt + timedelta(days=30 * ct[6])

        c.execute('''
        INSERT INTO contracts (vehicle_id, customer_id, business_mode, rental_method, repayment_day,
                               start_date, end_date, total_price, loan_amount, monthly_payment,
                               rent, loan_periods, deposit, paid_principal, loan_balance,
                               collected_deposit, collected_rent, contract_status,
                               snapshot_guidance_price, snapshot_invoice_price)
        VALUES (?, ?, '转租', ?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                ?, ?)
        ''', (ct[0], ct[12], ct[1], start_date_str, end_dt.strftime('%Y-%m-%d'),
              ct[2], ct[3], ct[4], ct[5], ct[6], ct[7], ct[8], ct[9], ct[10], ct[11], contract_status,
              vrow['guidance_price'], vrow['invoice_price']))

    # ====== 为活跃合同生成双向还款计划 ======
    for ct in real_contracts:
        if ct[9] == '已结清':
            continue

        vehicle_id = ct[0]
        contract_id = vehicle_id  # seed 数据中 1:1 对应
        periods = ct[6]
        factory_monthly = ct[4]   # 厂家月供
        customer_rent = ct[5]     # 客户月租

        c.execute("SELECT invoice_date FROM vehicles WHERE id = ?", (vehicle_id,))
        start_date = datetime.strptime(c.fetchone()['invoice_date'], '%Y-%m-%d')

        # 计算已还期数
        factory_paid_count = int(ct[8] / factory_monthly) if factory_monthly > 0 else 0
        customer_paid_count = int(ct[11] / customer_rent) if customer_rent > 0 else 0

        for p in range(1, periods + 1):
            due = start_date + timedelta(days=30 * p)
            due_str = due.strftime('%Y-%m-%d')
            is_past_due = due < datetime.now()

            # --- 厂家还款（公司 → 一汽解放）---
            if p <= factory_paid_count:
                f_status = '已还款'
                f_paid = (due + timedelta(days=p % 5)).strftime('%Y-%m-%d')
            elif p == factory_paid_count + 1 and is_past_due:
                f_status = '逾期'
                f_paid = None
            else:
                f_status = '待还款'
                f_paid = None

            c.execute('''
            INSERT INTO factory_repayments (contract_id, period, due_date, amount, status, paid_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', (contract_id, p, due_str, factory_monthly, f_status, f_paid))

            # --- 客户还款（客户 → 公司）---
            if p <= customer_paid_count:
                c_status = '已还款'
                c_paid = (due + timedelta(days=p % 3)).strftime('%Y-%m-%d')
            elif p == customer_paid_count + 1 and is_past_due:
                c_status = '逾期'
                c_paid = None
            else:
                c_status = '待还款'
                c_paid = None

            c.execute('''
            INSERT INTO repayments (contract_id, period, due_date, amount, status, paid_at)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', (contract_id, p, due_str, customer_rent, c_status, c_paid))

    conn.commit()
    conn.close()
    print("Seed data inserted successfully.")


# ================================================================
#  开发工具 — 删库重建（生产环境禁用）
# ================================================================
def reset_db():
    if os.path.exists(DATABASE):
        os.remove(DATABASE)
    init_db()
    seed_data()
    print("Database has been reset.")


if __name__ == '__main__':
    reset_db()
