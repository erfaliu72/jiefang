import sqlite3
import os
import re
from datetime import datetime, timedelta

DATABASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'jinjuyuan.db')

# ================================================================
#  数据库后端选择：设置环境变量 JJY_DB_HOST 则用 MySQL(RDS)，否则用本地 SQLite
# ================================================================
USE_MYSQL = bool(os.environ.get('JJY_DB_HOST'))

if USE_MYSQL:
    import pymysql
    from pymysql.cursors import DictCursor

    _MYSQL_CONF = {
        'host': os.environ.get('JJY_DB_HOST'),
        'port': int(os.environ.get('JJY_DB_PORT', '3306')),
        'user': os.environ.get('JJY_DB_USER', 'root'),
        'password': os.environ.get('JJY_DB_PASSWORD', ''),
        'database': os.environ.get('JJY_DB_NAME', 'jinjuyuan'),
        'charset': 'utf8mb4',
    }

    # --- SQL 方言翻译：把业务代码里的 SQLite 写法转成 MySQL ---
    _RE_DATETIME_NOW = re.compile(r"datetime\(\s*'now'\s*,\s*'localtime'\s*\)", re.IGNORECASE)
    _RE_DATETIME_NOW2 = re.compile(r"datetime\(\s*'now'\s*\)", re.IGNORECASE)
    _RE_DATE_MODIFIER = re.compile(
        r"date\(([^,)]+),\s*'([+-])\s*(\d+)\s*(day|month|year)'\s*\)",
        re.IGNORECASE
    )

    def _translate_sql(sql, has_params):
        # datetime('now','localtime') -> NOW()
        sql = _RE_DATETIME_NOW.sub('NOW()', sql)
        sql = _RE_DATETIME_NOW2.sub('NOW()', sql)
        # date(col, '+N day') -> DATE_ADD(col, INTERVAL N DAY) / date(col, '-N day') -> DATE_SUB(col, INTERVAL N DAY)
        sql = _RE_DATE_MODIFIER.sub(
            lambda m: f"DATE_ADD({m.group(1)}, INTERVAL {m.group(3)} {m.group(4).upper()})"
            if m.group(2) == '+' else
            f"DATE_SUB({m.group(1)}, INTERVAL {m.group(3)} {m.group(4).upper()})",
            sql
        )
        # INSERT OR IGNORE / OR REPLACE
        sql = re.sub(r'INSERT\s+OR\s+IGNORE', 'INSERT IGNORE', sql, flags=re.IGNORECASE)
        sql = re.sub(r'INSERT\s+OR\s+REPLACE', 'REPLACE', sql, flags=re.IGNORECASE)
        # ON CONFLICT(x) DO UPDATE SET a=excluded.a,... -> ON DUPLICATE KEY UPDATE a=VALUES(a),...
        m = re.search(r'ON\s+CONFLICT\s*\([^)]*\)\s+DO\s+UPDATE\s+SET\s+(.*)$', sql, flags=re.IGNORECASE | re.DOTALL)
        if m:
            set_clause = m.group(1)
            set_clause = re.sub(r'excluded\.(\w+)', r'VALUES(\1)', set_clause, flags=re.IGNORECASE)
            sql = sql[:m.start()] + 'ON DUPLICATE KEY UPDATE ' + set_clause
        # 占位符 ? -> %s（pymysql 用 %s）。先把已有的 % 转义成 %%，避免与 paramstyle 冲突
        if has_params:
            sql = sql.replace('%', '%%')
            sql = sql.replace('?', '%s')
        else:
            # 无参数时 pymysql 不做格式化，% 原样保留
            sql = sql.replace('?', '%s')
        # SQLite 标量 MIN(a,b) -> MySQL LEAST(a,b)。匹配两参数模式避免误伤聚合 MIN(col)。
        if has_params:
            sql = re.sub(r'\bMIN\s*\(([^,)]+)\s*,\s*([^,)]+)\s*\)', r'LEAST(\1, \2)', sql, flags=re.IGNORECASE)
        return sql

    class _Cursor:
        """包装 pymysql DictCursor，提供与 sqlite3 一致的接口。"""
        def __init__(self, raw):
            self._raw = raw

        def execute(self, sql, params=None):
            sql2 = _translate_sql(sql, params is not None and (not hasattr(params, '__len__') or len(params) > 0))
            if params is None or (hasattr(params, '__len__') and len(params) == 0):
                self._raw.execute(sql2)
            else:
                self._raw.execute(sql2, params)
            return self  # 返回 self 以支持链式调用：c.execute(sql).fetchone()

        def fetchone(self):
            return self._raw.fetchone()

        def fetchall(self):
            return self._raw.fetchall()

        @property
        def lastrowid(self):
            return self._raw.lastrowid

        @property
        def rowcount(self):
            return self._raw.rowcount

        def close(self):
            self._raw.close()

    class _Conn:
        """包装 pymysql 连接，提供与 sqlite3.Connection 一致的接口。"""
        def __init__(self, raw):
            self._raw = raw

        def cursor(self):
            return _Cursor(self._raw.cursor())

        def execute(self, sql, params=None):
            cur = self.cursor()
            cur.execute(sql, params)
            return cur

        def commit(self):
            self._raw.commit()

        def rollback(self):
            self._raw.rollback()

        def close(self):
            self._raw.close()

    def get_db():
        raw = pymysql.connect(cursorclass=DictCursor, autocommit=False, **_MYSQL_CONF)
        # 关闭 only_full_group_by / strict 模式，兼容 SQLite 宽松行为
        with raw.cursor() as c:
            c.execute("SET SESSION sql_mode=''")
        return _Conn(raw)

else:
    def get_db():
        conn = sqlite3.connect(DATABASE)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn


def _ddl(sql):
    """建表 DDL：SQLite 原样执行；MySQL 时翻译方言。"""
    if not USE_MYSQL:
        return sql
    s = sql
    # 主键自增
    s = re.sub(r'INTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT', 'BIGINT AUTO_INCREMENT PRIMARY KEY', s, flags=re.IGNORECASE)
    # 唯一文本列需指定长度（TEXT 不能做唯一键）
    s = re.sub(r'\bTEXT\s+UNIQUE\b', 'VARCHAR(191) UNIQUE', s, flags=re.IGNORECASE)
    # created_at 这类：TEXT DEFAULT (datetime('now','localtime')) -> DATETIME DEFAULT CURRENT_TIMESTAMP
    s = re.sub(r"\bTEXT\s+DEFAULT\s*\(\s*datetime\(\s*'now'\s*,\s*'localtime'\s*\)\s*\)",
               'DATETIME DEFAULT CURRENT_TIMESTAMP', s, flags=re.IGNORECASE)
    # 其余裸的 DEFAULT (datetime(...)) -> DEFAULT CURRENT_TIMESTAMP
    s = re.sub(r"DEFAULT\s*\(\s*datetime\(\s*'now'\s*,\s*'localtime'\s*\)\s*\)", 'DEFAULT CURRENT_TIMESTAMP', s, flags=re.IGNORECASE)
    # MySQL 的 TEXT 列不能有默认值：TEXT DEFAULT 'x' -> VARCHAR(255) DEFAULT 'x'
    s = re.sub(r"\bTEXT\s+DEFAULT\b", 'VARCHAR(255) DEFAULT', s, flags=re.IGNORECASE)
    # 进入索引/唯一约束的 TEXT 列必须有长度 -> VARCHAR(191)
    _indexed_cols = ['role', 'page_key', 'action_key', 'resource_key', 'field_key',
                     'job', 'run_date', 'status', 'customer_phone', 'customer_name',
                     'waiver_kind', 'accrued_date', 'car_type', 'is_new']
    for col in _indexed_cols:
        s = re.sub(r'\b(' + col + r')\s+TEXT\b', r'\1 VARCHAR(191)', s, flags=re.IGNORECASE)
    # 类型映射：REAL -> DOUBLE；剩余 INTEGER -> BIGINT；TEXT 保留
    s = re.sub(r'\bREAL\b', 'DOUBLE', s, flags=re.IGNORECASE)
    s = re.sub(r'\bINTEGER\b', 'BIGINT', s, flags=re.IGNORECASE)
    return s

# ================================================================
#  安全建表 — 仅在表不存在时创建，不会删除已有数据
# ================================================================
def init_db():
    conn = get_db()
    _raw_c = conn.cursor()
    if USE_MYSQL:
        # MySQL 外键严格要求被引用表先建；建表期间关闭外键检查，避免表顺序问题
        _raw_c.execute("SET FOREIGN_KEY_CHECKS=0")

    class _DDLCursor:
        """init_db 专用：自动把建表 DDL 翻译成目标方言；
        MySQL 不支持 CREATE INDEX IF NOT EXISTS，吞掉重复建索引/列的报错。"""
        def __init__(self, cur):
            self._c = cur
        def execute(self, sql, params=None):
            s = sql
            if 'CREATE TABLE' in s.upper() or 'ALTER TABLE' in s.upper():
                s = _ddl(s)
            if USE_MYSQL and re.search(r'CREATE\s+(?:UNIQUE\s+)?INDEX\s+IF\s+NOT\s+EXISTS', s, re.IGNORECASE):
                s = re.sub(
                    r'CREATE\s+(UNIQUE\s+)?INDEX\s+IF\s+NOT\s+EXISTS',
                    lambda m: f"CREATE {m.group(1) or ''}INDEX",
                    s,
                    flags=re.IGNORECASE,
                )
            try:
                if params is None:
                    return self._c.execute(s)
                return self._c.execute(s, params)
            except Exception as e:
                # MySQL 重复建索引/加列会报错，建表用 IF NOT EXISTS 不会
                if USE_MYSQL and ('CREATE INDEX' in s.upper() or 'ALTER TABLE' in s.upper()):
                    return None
                raise
        def fetchone(self): return self._c.fetchone()
        def fetchall(self): return self._c.fetchall()
        @property
        def lastrowid(self): return self._c.lastrowid
        @property
        def rowcount(self): return self._c.rowcount
    c = _DDLCursor(_raw_c)

    # ====== 车辆资产表 ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS vehicles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vin TEXT UNIQUE NOT NULL,
        plate_number TEXT,
        company TEXT,
        car_type TEXT,
        vehicle_category TEXT,
        vehicle_engine_battery TEXT,
        vehicle_power_battery TEXT,
        vehicle_color TEXT,
        vehicle_box_type TEXT,
        box_type_remark TEXT,
        purchase_price REAL DEFAULT 0,
        tax_rate REAL DEFAULT 0.13,
        estimated_residual_value REAL DEFAULT 0,
        depreciation_months INTEGER DEFAULT 60,
        insurance_expiry_date TEXT,
        annual_review_date TEXT,
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

    # ====== 销售报单表（PRD 系统入口）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS sales_orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        payment_date TEXT,
        customer_name TEXT NOT NULL,
        customer_phone TEXT,
        customer_id_card TEXT,
        sales_mode TEXT NOT NULL,
        vehicle_id INTEGER,
        vin TEXT NOT NULL,
        car_type TEXT,
        vehicle_color TEXT,
        plate_number TEXT,
        tail_plate TEXT,
        lease_term TEXT,
        cargo_length TEXT,
        sale_total_price REAL DEFAULT 0,
        payment_category TEXT,
        car_purchase_amount REAL DEFAULT 0,
        vehicle_rent_amount REAL DEFAULT 0,
        receiving_company TEXT,
        wechat_interest REAL DEFAULT 0,
        wechat_registration_fee REAL DEFAULT 0,
        wechat_purchase_tax REAL DEFAULT 0,
        sales_advisor TEXT,
        full_package INTEGER DEFAULT 0,
        wechat_private_fee REAL DEFAULT 0,
        gifted_items TEXT,
        deposit_amount REAL DEFAULT 0,
        order_status TEXT DEFAULT '待财务确认',
        snapshot_guidance_price REAL DEFAULT 0,
        snapshot_lease_installment_price REAL DEFAULT 0,
        snapshot_sale_total_price REAL DEFAULT 0,
        price_check_status TEXT DEFAULT '无需审批',
        price_exception_reason TEXT,
        boss_price_approved_by TEXT,
        boss_price_approved_at TEXT,
        finance_confirmed_by TEXT,
        finance_confirmed_at TEXT,
        finance_bank_serial TEXT,
        finance_bank_receipt_path TEXT,
        customer_plan_match_status TEXT DEFAULT '未生成',
        factory_plan_match_status TEXT DEFAULT '未上传',
        plan_compare_summary TEXT,
        saved_at TEXT,
        expires_at TEXT,
        voided_at TEXT,
        void_reason TEXT,
        voided_by TEXT,
        contract_id INTEGER,
        remark TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 车辆指导价历史（老板调价留痕；后续成交用新价，历史报单/合同快照不变）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS vehicle_guidance_price_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        old_price REAL DEFAULT 0,
        new_price REAL DEFAULT 0,
        changed_by TEXT,
        effective_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id)
    )
    ''')

    # ====== 车型指导价主数据（老板按车型维护，后续报单优先使用）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS model_guidance_prices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        car_type TEXT NOT NULL,
        is_new TEXT DEFAULT '新车',
        guidance_price REAL DEFAULT 0,
        lease_installment_price REAL DEFAULT 0,
        sale_total_price REAL DEFAULT 0,
        lease_deposit_ratio REAL DEFAULT 0,
        lease_repayment_ratio REAL DEFAULT 0,
        sale_down_payment_ratio REAL DEFAULT 0,
        sale_repayment_ratio REAL DEFAULT 0,
        remark TEXT,
        updated_by TEXT,
        updated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')

    c.execute('''
    CREATE TABLE IF NOT EXISTS model_guidance_price_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        car_type TEXT NOT NULL,
        is_new TEXT DEFAULT '新车',
        price_kind TEXT DEFAULT 'legacy',
        old_price REAL DEFAULT 0,
        new_price REAL DEFAULT 0,
        changed_by TEXT,
        effective_at TEXT,
        affected_vehicle_count INTEGER DEFAULT 0,
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
        sales_order_id INTEGER,
        contract_type TEXT DEFAULT '租赁',
        business_mode TEXT DEFAULT '转租',
        rental_method TEXT,
        repayment_day INTEGER DEFAULT 1,
        start_date TEXT,
        end_date TEXT,
        total_price REAL DEFAULT 0,
        customer_loan_amount REAL DEFAULT 0,
        loan_amount REAL DEFAULT 0,
        monthly_payment REAL DEFAULT 0,
        rent REAL DEFAULT 0,
        loan_periods INTEGER DEFAULT 0,
        company TEXT,
        yard TEXT,
        lease_bank_name TEXT,
        lease_bank_card_no TEXT,
        factory_guarantee_deposit REAL DEFAULT 0,
        factory_repayment_months INTEGER DEFAULT 0,
        factory_periods INTEGER DEFAULT 0,
        deposit REAL DEFAULT 0,
        down_payment REAL DEFAULT 0,
        down_payment_status TEXT DEFAULT '待收',
        deposit_status TEXT DEFAULT '待收',
        delivery_status TEXT DEFAULT '待出库',
        delivery_date TEXT,
        delivery_photo_path TEXT,
        delivery_document_path TEXT,
        paid_principal REAL DEFAULT 0,
        loan_balance TEXT DEFAULT '0',
        collected_deposit REAL DEFAULT 0,
        collected_rent REAL DEFAULT 0,
        contract_status TEXT DEFAULT '执行中',
        -- PRD 价格快照：成交时锁定的参考价格，后续调价不影响
        snapshot_guidance_price REAL DEFAULT 0,
        snapshot_invoice_price REAL DEFAULT 0,
        customer_plan_match_status TEXT DEFAULT '未比对',
        plan_compare_summary TEXT,
        expected_profit_floor REAL DEFAULT 0,
        expected_profit_ceiling REAL,
        contract_file TEXT,
        remark TEXT,
        loan_remark TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id),
        FOREIGN KEY (customer_id) REFERENCES customers (id),
        FOREIGN KEY (sales_order_id) REFERENCES sales_orders (id)
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
        paid_amount REAL DEFAULT 0,
        verified_amount REAL DEFAULT 0,
        reported_amount REAL DEFAULT 0,
        status TEXT DEFAULT '待还款',
        paid_at TEXT,
        screenshot_path TEXT,
        bank_receipt_path TEXT,
        bank_serial TEXT,
        verified_by TEXT,
        verified_at TEXT,
        waterfall_summary TEXT,
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

    # ====== 合同费用项（瀑布式核销优先级）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS contract_fee_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        fee_type TEXT NOT NULL,
        description TEXT,
        amount_due REAL DEFAULT 0,
        amount_paid REAL DEFAULT 0,
        due_date TEXT,
        status TEXT DEFAULT '待支付',
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 核销分配明细（审计瀑布式分配结果）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS reconciliation_allocations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        repayment_id INTEGER NOT NULL,
        contract_id INTEGER NOT NULL,
        fee_item_id INTEGER,
        allocation_type TEXT NOT NULL,
        allocated_amount REAL DEFAULT 0,
        note TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (repayment_id) REFERENCES repayments (id),
        FOREIGN KEY (contract_id) REFERENCES contracts (id),
        FOREIGN KEY (fee_item_id) REFERENCES contract_fee_items (id)
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

    # ====== 合同首付款/首次支付审核（运营发起 → 财务）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS contract_initial_payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        payment_type TEXT DEFAULT '首付款',
        amount REAL DEFAULT 0,
        received_amount REAL DEFAULT 0,
        shortage_amount REAL DEFAULT 0,
        shortage_reason TEXT,
        promised_repay_date TEXT,
        shortage_status TEXT DEFAULT '无欠款',
        bank_serial TEXT,
        customer_screenshot_path TEXT,
        bank_receipt_path TEXT,
        status TEXT DEFAULT '待审批',
        requested_by TEXT,
        approved_by TEXT,
        approved_at TEXT,
        remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')

    # ====== 通用审批流程表（出库审批 / 首付款审核 / 锁车审批 / 旧车入库审批）======
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
        evidence_path TEXT,
        promised_repay_date TEXT,
        completed_at TEXT,
        closed_by TEXT,
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
        sales_status TEXT DEFAULT '待登记',
        fleet_status TEXT DEFAULT '待填写',
        operator_status TEXT DEFAULT '待填写',
        finance_status TEXT DEFAULT '待填写',
        boss_approved INTEGER DEFAULT 0,
        boss_approved_by TEXT,
        boss_approved_at TEXT,
        rejected_by TEXT,
        rejected_at TEXT,
        reject_reason TEXT,
        resubmitted_by TEXT,
        resubmitted_at TEXT,
        resubmit_note TEXT,
        finance_approved INTEGER DEFAULT 0,
        finance_approved_by TEXT,
        finance_approved_at TEXT,
        paid_out INTEGER DEFAULT 0,
        paid_out_by TEXT,
        paid_out_at TEXT,
        refund_company_name TEXT,
        refund_bank_name TEXT,
        refund_bank_card_no TEXT,
        lease_bank_name TEXT,
        lease_bank_card_no TEXT,
        leader_remark TEXT,
        -- 随车工具检查 (1=有 0=无)
        tool_triangle INTEGER DEFAULT 0,
        tool_vest INTEGER DEFAULT 0,
        tool_extinguisher INTEGER DEFAULT 0,
        tool_wedge INTEGER DEFAULT 0,
        tool_jack INTEGER DEFAULT 0,
        tool_kit INTEGER DEFAULT 0,
        tent_pole INTEGER DEFAULT 0,
        car_wash_fee INTEGER DEFAULT 0,
        body_ad_clean INTEGER DEFAULT 0,
        other_info TEXT,
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

    # ====== 角色页面/动作映射（A1/U3，可由静态矩阵种子化）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS role_pages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        role TEXT NOT NULL,
        page_key TEXT NOT NULL,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(role, page_key)
    )
    ''')
    c.execute('''
    CREATE TABLE IF NOT EXISTS role_actions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        role TEXT NOT NULL,
        action_key TEXT NOT NULL,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(role, action_key)
    )
    ''')
    c.execute('''
    CREATE TABLE IF NOT EXISTS role_field_permissions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        role TEXT NOT NULL,
        resource_key TEXT NOT NULL,
        field_key TEXT NOT NULL,
        can_view INTEGER DEFAULT 1,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(role, resource_key, field_key)
    )
    ''')

    # ====== 客户黑名单（S，报单提交硬阻塞）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS customer_blacklist (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_name TEXT,
        customer_phone TEXT,
        id_card TEXT,
        level TEXT DEFAULT '禁止报单',
        reason TEXT,
        status TEXT DEFAULT '生效',
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_blacklist_phone ON customer_blacklist(customer_phone, status)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_blacklist_name ON customer_blacklist(customer_name, status)')

    # ====== 车辆返利（O，visibility 由权限矩阵控制，不在表内冗余）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS vehicle_rebates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vehicle_id INTEGER NOT NULL,
        contract_id INTEGER,
        rebate_amount REAL DEFAULT 0,
        rebate_date TEXT,
        remark TEXT,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (vehicle_id) REFERENCES vehicles (id),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_vehicle_rebates_vehicle ON vehicle_rebates(vehicle_id)')

    # ====== 减免审批表 (J3/K, 决策项30, 统一表名 waivers) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS waivers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        waiver_kind TEXT NOT NULL DEFAULT 'rent',
        target_period_list TEXT,
        waive_amount REAL,
        reason TEXT,
        attachment_path TEXT,
        status TEXT DEFAULT '待审批',
        sales_applied_by TEXT,
        sales_applied_at TEXT,
        boss_approved_by TEXT,
        boss_approved_at TEXT,
        finance_reviewed_by TEXT,
        finance_reviewed_at TEXT,
        revoke_reason TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_waivers_contract ON waivers(contract_id)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_waivers_status ON waivers(status)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_waivers_contract_kind_status ON waivers(contract_id, waiver_kind, status)')

    # ====== 滞纳金每日计提台账 (J1) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS late_fee_ledger (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        repayment_id INTEGER NOT NULL,
        contract_id INTEGER,
        period INTEGER,
        accrued_date TEXT NOT NULL,
        outstanding REAL DEFAULT 0,
        daily_amount REAL DEFAULT 0,
        cumulative_amount REAL DEFAULT 0,
        waived INTEGER DEFAULT 0,
        waived_amount REAL DEFAULT 0,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(repayment_id, accrued_date),
        FOREIGN KEY (repayment_id) REFERENCES repayments (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_late_fee_repayment ON late_fee_ledger(repayment_id)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_late_fee_contract ON late_fee_ledger(contract_id)')

    # ====== 客户预收余额 (I2 多还抵不完的部分) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS customer_prepayments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        customer_id INTEGER,
        amount REAL DEFAULT 0,
        source_repayment_id INTEGER,
        source_bank_serial TEXT,
        balance REAL DEFAULT 0,
        remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_prepay_contract ON customer_prepayments(contract_id)')

    # ====== 过户单 (N1/N2 写入) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS ownership_transfers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        vehicle_id INTEGER,
        settle_type TEXT DEFAULT 'natural_settle',
        status TEXT DEFAULT '待过户',
        transfer_date TEXT,
        new_owner_name TEXT,
        new_owner_id_card TEXT,
        transfer_doc_path TEXT,
        idempotency_key TEXT UNIQUE,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        completed_by TEXT,
        completed_at TEXT,
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_transfer_contract ON ownership_transfers(contract_id)')

    # ====== 收款公司主数据 (R) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS receiving_companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT UNIQUE NOT NULL,
        bank_name TEXT,
        bank_account_no TEXT,
        tax_no TEXT,
        status TEXT DEFAULT '启用',
        remark TEXT,
        updated_by TEXT,
        updated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_receiving_companies_status ON receiving_companies(status)')

    # ====== 发票申请/作废/红冲 (L1/L2/L3) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS invoice_requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        period INTEGER,
        amount REAL DEFAULT 0,
        receiving_company TEXT,
        title_type TEXT,
        invoice_entity_name TEXT,
        invoice_entity_tax_no TEXT,
        invoice_no TEXT,
        invoiced_at TEXT,
        invoice_file_path TEXT,
        status TEXT DEFAULT '待开票',
        applied_by TEXT,
        applied_at TEXT,
        boss_approved_by TEXT,
        boss_approved_at TEXT,
        processed_by TEXT,
        voided_by TEXT,
        voided_at TEXT,
        void_reason TEXT,
        previous_invoice_id INTEGER,
        red_invoice_no TEXT,
        red_invoiced_at TEXT,
        red_reason TEXT,
        red_certificate_path TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        FOREIGN KEY (contract_id) REFERENCES contracts (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_invoice_contract ON invoice_requests(contract_id)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_invoice_status ON invoice_requests(status)')

    # ====== 日终批处理运行记录 (H1 (job,run_date) 幂等 + 补跑) ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS daily_job_runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job TEXT NOT NULL,
        run_date TEXT NOT NULL,
        ran_at TEXT DEFAULT (datetime('now','localtime')),
        summary TEXT,
        UNIQUE(job, run_date)
    )
    ''')

    # ====== 挂账应收：首次付款不足、每期少还均进入此表 ======
    c.execute('''
    CREATE TABLE IF NOT EXISTS receivables (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        repayment_id INTEGER,
        initial_payment_id INTEGER,
        receivable_type TEXT NOT NULL,
        source_period INTEGER,
        amount REAL DEFAULT 0,
        paid_amount REAL DEFAULT 0,
        due_date TEXT,
        promised_repay_date TEXT,
        reason TEXT,
        status TEXT DEFAULT '待归还',
        late_fee_accrued REAL DEFAULT 0,
        created_by TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime')),
        settled_at TEXT,
        FOREIGN KEY (contract_id) REFERENCES contracts (id),
        FOREIGN KEY (repayment_id) REFERENCES repayments (id),
        FOREIGN KEY (initial_payment_id) REFERENCES contract_initial_payments (id)
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_receivables_contract ON receivables(contract_id)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_receivables_status ON receivables(status)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_receivables_repayment ON receivables(repayment_id)')

    # ====== 以租代售金融方案（老板按车型维护 n 种：首付 + 每期价格 + 期数）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS finance_plans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        car_type TEXT NOT NULL,
        plan_name TEXT,
        down_payment REAL DEFAULT 0,
        period_price REAL DEFAULT 0,
        periods INTEGER DEFAULT 0,
        tail_plate_price REAL DEFAULT 0,
        condition TEXT DEFAULT '新车',
        box_type TEXT DEFAULT '',
        status TEXT DEFAULT '启用',
        sort_order INTEGER DEFAULT 0,
        remark TEXT,
        created_by TEXT,
        updated_by TEXT,
        updated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_finance_plans_car_type ON finance_plans(car_type, status)')

    # ====== 报单驳回退款（老板驳回后发起，财务执行打款）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS order_refunds (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sales_order_id INTEGER NOT NULL,
        vehicle_id INTEGER,
        contract_id INTEGER,
        refund_amount REAL DEFAULT 0,
        refund_serial TEXT,
        refund_paid_amount REAL DEFAULT 0,
        status TEXT DEFAULT '待退款',
        initiated_by TEXT,
        initiated_at TEXT,
        executed_by TEXT,
        executed_at TEXT,
        customer_name TEXT,
        customer_phone TEXT,
        bank_name TEXT,
        bank_card_no TEXT,
        remark TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_order_refunds_order ON order_refunds(sales_order_id, status)')

    # ====== 数据字典（SKU 改造：品牌/品系/厢型/电池度数/马力/档位等维度动态化）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS data_dictionaries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT NOT NULL,          -- 当前：brand/product_series/box_type/battery_capacity/horsepower/gear_position/battery_brand（历史兼容值可能仍存在）
        value TEXT NOT NULL,             -- 选项值（如 解放 / J6F / 厢货 / 100度 / 锡柴180 / 六档）
        energy_type TEXT,                -- 适用能源类型：纯电/混动/燃油车；空=通用维度
        sort_order INTEGER DEFAULT 0,
        status TEXT DEFAULT '启用',      -- 启用/停用
        remark TEXT,
        created_by TEXT,
        updated_by TEXT,
        updated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_data_dictionaries_cat ON data_dictionaries(category, status)')

    # ====== SKU 主表（SKU 改造：基准车型+成色+厢型+尾板 组合索引，用于库存聚合与销售筛选）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS skus (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        car_type TEXT NOT NULL,          -- 基准车型（normalize_base_car_type 后）
        condition TEXT DEFAULT '新车',    -- 成色 新车/二手车
        box_type TEXT,                    -- 厢型（底盘车不入 SKU）
        tailgate TEXT DEFAULT '无',       -- 尾板 有/无
        status TEXT DEFAULT '生效',       -- 生效/失效（老板可停用某组合）
        created_by TEXT,
        updated_by TEXT,
        updated_at TEXT,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_skus_comb ON skus(car_type, condition, box_type, tailgate)')

    # ====== 合同多车桥表（合同 → 多辆 VIN，contracts.vehicle_id 保留主车）======
    c.execute('''
    CREATE TABLE IF NOT EXISTS contract_vehicles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        contract_id INTEGER NOT NULL,
        vehicle_id INTEGER NOT NULL,
        is_primary INTEGER DEFAULT 0,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    )
    ''')
    c.execute('CREATE INDEX IF NOT EXISTS idx_contract_vehicles_contract ON contract_vehicles(contract_id)')
    c.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_contract_vehicles_uniq ON contract_vehicles(contract_id, vehicle_id)')

    # === 安全添加新列（如果表已存在但缺少新字段）===
    safe_alter_columns = [
        ("contracts", "snapshot_guidance_price", "REAL DEFAULT 0"),
        ("contracts", "snapshot_invoice_price", "REAL DEFAULT 0"),
        ("contracts", "customer_plan_match_status", "TEXT DEFAULT '未比对'"),
        ("contracts", "plan_compare_summary", "TEXT"),
        ("contracts", "expected_profit_floor", "REAL DEFAULT 0"),
        ("contracts", "expected_profit_ceiling", "REAL"),
        ("contracts", "contract_type", "TEXT DEFAULT '租赁'"),
        ("contracts", "customer_loan_amount", "REAL DEFAULT 0"),
        ("contracts", "down_payment", "REAL DEFAULT 0"),
        ("contracts", "company", "TEXT"),
        ("contracts", "yard", "TEXT"),
        ("contracts", "lease_bank_name", "TEXT"),
        ("contracts", "lease_bank_card_no", "TEXT"),
        ("contracts", "contract_status", "TEXT DEFAULT '执行中'"),
        ("contracts", "factory_guarantee_deposit", "REAL DEFAULT 0"),
        ("contracts", "factory_repayment_months", "INTEGER DEFAULT 0"),
        ("contracts", "factory_periods", "INTEGER DEFAULT 0"),
        ("contracts", "sales_order_id", "INTEGER"),
        ("contracts", "down_payment_status", "TEXT DEFAULT '待收'"),
        ("contracts", "deposit_status", "TEXT DEFAULT '待收'"),
        ("contracts", "delivery_status", "TEXT DEFAULT '待出库'"),
        ("contracts", "delivery_date", "TEXT"),
        ("contracts", "delivery_photo_path", "TEXT"),
        ("contracts", "delivery_document_path", "TEXT"),
        ("contracts", "early_settlement_amount", "REAL DEFAULT 0"),
        ("contracts", "early_settlement_serial", "TEXT"),
        ("contracts", "early_settled_at", "TEXT"),
        ("contracts", "early_settled_by", "TEXT"),
        ("vehicles", "customer_name", "TEXT"),
        ("vehicles", "vehicle_category", "TEXT"),
        ("vehicles", "vehicle_engine_battery", "TEXT"),
        ("vehicles", "vehicle_power_battery", "TEXT"),
        ("vehicles", "vehicle_color", "TEXT"),
        ("vehicles", "vehicle_box_type", "TEXT"),
        ("vehicles", "box_type_remark", "TEXT"),
        ("vehicles", "invoice_contract_file", "TEXT"),
        ("vehicles", "purchase_price", "REAL DEFAULT 0"),
        ("vehicles", "tax_rate", "REAL DEFAULT 0.13"),
        ("vehicles", "estimated_residual_value", "REAL DEFAULT 0"),
        ("vehicles", "depreciation_months", "INTEGER DEFAULT 60"),
        ("vehicles", "insurance_expiry_date", "TEXT"),
        ("vehicles", "annual_review_date", "TEXT"),
        ("contracts", "contract_file", "TEXT"),
        ("contracts", "lease_bank_name", "TEXT"),
        ("contracts", "lease_bank_card_no", "TEXT"),
        ("contracts", "created_by", "TEXT"),
        ("return_inspections", "sales_status", "TEXT DEFAULT '待登记'"),
        ("return_inspections", "fleet_status", "TEXT DEFAULT '待填写'"),
        ("return_inspections", "operator_status", "TEXT DEFAULT '待填写'"),
        ("return_inspections", "finance_status", "TEXT DEFAULT '待填写'"),
        ("return_inspections", "boss_approved", "INTEGER DEFAULT 0"),
        ("return_inspections", "boss_approved_by", "TEXT"),
        ("return_inspections", "boss_approved_at", "TEXT"),
        ("return_inspections", "finance_approved", "INTEGER DEFAULT 0"),
        ("return_inspections", "finance_approved_by", "TEXT"),
        ("return_inspections", "finance_approved_at", "TEXT"),
        ("return_inspections", "paid_out", "INTEGER DEFAULT 0"),
        ("return_inspections", "paid_out_by", "TEXT"),
        ("return_inspections", "paid_out_at", "TEXT"),
        ("return_inspections", "refund_serial", "TEXT"),
        ("return_inspections", "refund_paid_amount", "REAL"),
        ("return_inspections", "refund_company_name", "TEXT"),
        ("return_inspections", "refund_bank_name", "TEXT"),
        ("return_inspections", "refund_bank_card_no", "TEXT"),
        ("return_inspections", "lease_bank_name", "TEXT"),
        ("return_inspections", "lease_bank_card_no", "TEXT"),
        ("return_inspections", "leader_remark", "TEXT"),
        ("sales_orders", "is_new", "TEXT DEFAULT '新车'"),
        ("sales_orders", "vehicle_brand", "TEXT"),
        ("sales_orders", "lease_start_date", "TEXT"),
        ("sales_orders", "vehicle_category", "TEXT"),
        ("sales_orders", "vehicle_cab", "TEXT"),
        ("sales_orders", "vehicle_engine_battery", "TEXT"),
        ("sales_orders", "vehicle_power_battery", "TEXT"),
        ("sales_orders", "vehicle_gearbox", "TEXT"),
        ("sales_orders", "vehicle_box_type", "TEXT"),
        ("sales_orders", "car_type", "TEXT"),
        ("sales_orders", "vehicle_color", "TEXT"),
        ("sales_orders", "plate_number", "TEXT"),
        ("sales_orders", "lease_term", "TEXT"),
        ("sales_orders", "cargo_length", "TEXT"),
        ("sales_orders", "sale_total_price", "REAL DEFAULT 0"),
        ("sales_orders", "payment_category", "TEXT"),
        ("sales_orders", "car_purchase_amount", "REAL DEFAULT 0"),
        ("sales_orders", "vehicle_rent_amount", "REAL DEFAULT 0"),
        ("sales_orders", "receiving_company", "TEXT"),
        ("sales_orders", "wechat_interest", "REAL DEFAULT 0"),
        ("sales_orders", "wechat_registration_fee", "REAL DEFAULT 0"),
        ("sales_orders", "wechat_purchase_tax", "REAL DEFAULT 0"),
        ("sales_orders", "sales_advisor", "TEXT"),
        ("sales_orders", "snapshot_guidance_price", "REAL DEFAULT 0"),
        ("sales_orders", "snapshot_lease_installment_price", "REAL DEFAULT 0"),
        ("sales_orders", "snapshot_sale_total_price", "REAL DEFAULT 0"),
        ("sales_orders", "price_check_status", "TEXT DEFAULT '无需审批'"),
        ("sales_orders", "price_exception_reason", "TEXT"),
        ("sales_orders", "boss_price_approved_by", "TEXT"),
        ("sales_orders", "boss_price_approved_at", "TEXT"),
        ("sales_orders", "finance_confirmed_by", "TEXT"),
        ("sales_orders", "finance_confirmed_at", "TEXT"),
        ("sales_orders", "customer_plan_match_status", "TEXT DEFAULT '未生成'"),
        ("sales_orders", "factory_plan_match_status", "TEXT DEFAULT '未上传'"),
        ("sales_orders", "plan_compare_summary", "TEXT"),
        ("sales_orders", "saved_at", "TEXT"),
        ("sales_orders", "expires_at", "TEXT"),
        ("sales_orders", "voided_at", "TEXT"),
        ("sales_orders", "void_reason", "TEXT"),
        ("sales_orders", "voided_by", "TEXT"),
        ("model_guidance_prices", "lease_installment_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "sale_total_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "lease_deposit_ratio", "REAL DEFAULT 0"),
        ("model_guidance_prices", "lease_repayment_ratio", "REAL DEFAULT 0"),
        ("model_guidance_prices", "sale_down_payment_ratio", "REAL DEFAULT 0"),
        ("model_guidance_prices", "sale_repayment_ratio", "REAL DEFAULT 0"),
        ("model_guidance_prices", "product_code", "TEXT"),
        ("model_guidance_prices", "fuel_type", "TEXT"),
        ("model_guidance_prices", "chassis_base_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "landing_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "interest_free_plan", "TEXT"),
        ("model_guidance_prices", "rent_to_buy_plan", "TEXT"),
        ("model_guidance_prices", "min_loan_plan", "TEXT"),
        ("model_guidance_prices", "lease_plan", "TEXT"),
        ("model_guidance_prices", "is_new", "TEXT DEFAULT '新车'"),
        ("model_guidance_price_history", "price_kind", "TEXT DEFAULT 'legacy'"),
        ("model_guidance_price_history", "is_new", "TEXT DEFAULT '新车'"),
        ("contract_initial_payments", "received_amount", "REAL DEFAULT 0"),
        ("contract_initial_payments", "shortage_amount", "REAL DEFAULT 0"),
        ("contract_initial_payments", "shortage_reason", "TEXT"),
        ("contract_initial_payments", "promised_repay_date", "TEXT"),
        ("contract_initial_payments", "shortage_status", "TEXT DEFAULT '无欠款'"),
        ("contract_initial_payments", "bank_serial", "TEXT"),
        ("repayments", "paid_amount", "REAL DEFAULT 0"),
        ("repayments", "verified_amount", "REAL DEFAULT 0"),
        ("repayments", "screenshot_path", "TEXT"),
        ("repayments", "bank_receipt_path", "TEXT"),
        ("repayments", "bank_serial", "TEXT"),
        ("repayments", "verified_by", "TEXT"),
        ("repayments", "verified_at", "TEXT"),
        ("repayments", "waterfall_summary", "TEXT"),
        ("repayments", "reported_amount", "REAL DEFAULT 0"),
        # === W2: repayments 追加列 (v3 附录 W) ===
        ("repayments", "extra_alloc_confirmed_by", "TEXT"),
        ("repayments", "extra_alloc_confirmed_at", "TEXT"),
        ("repayments", "extra_alloc_periods", "TEXT"),
        ("repayments", "idempotency_key", "TEXT"),
        # === vehicles 锁车状态 (H4, 与 status 解耦) ===
        ("vehicles", "lock_status", "TEXT DEFAULT '未锁'"),
        ("vehicles", "pre_repair_status", "TEXT"),
        # === lock_requests H4 标记式审批流追加列 ===
        ("lock_requests", "action", "TEXT DEFAULT 'lock'"),
        ("lock_requests", "ops_approved_by", "TEXT"),
        ("lock_requests", "ops_approved_at", "TEXT"),
        ("lock_requests", "boss_approved_by", "TEXT"),
        ("lock_requests", "boss_approved_at", "TEXT"),
        ("lock_requests", "reject_reason", "TEXT"),
        ("lock_requests", "locked_by", "TEXT"),
        ("lock_requests", "locked_at", "TEXT"),
        ("lock_requests", "unlocked_by", "TEXT"),
        ("lock_requests", "unlocked_at", "TEXT"),
        ("return_inspections", "needs_repair", "INTEGER DEFAULT 0"),
        ("return_inspections", "repair_reason", "TEXT"),
        # === 新车 Excel 批量上传入库：经销商买断库存表的扩展维度列 ===
        ("vehicles", "tech_route", "TEXT"),                      # 技术路线
        ("vehicles", "dealer_code", "TEXT"),                     # 经销商代码
        ("vehicles", "import_raw", "TEXT"),                      # 原始Excel全部52列(JSON)
        # === 20260804 最新模板（最终定稿）：车辆入库导入-月份发票 25列维度 ===
        ("vehicles", "condition", "TEXT"),                       # 成色
        ("vehicles", "brand", "TEXT"),                           # 品牌
        ("vehicles", "product_series", "TEXT"),                  # 品系
        ("vehicles", "battery_capacity", "TEXT"),                # 电池度数
        ("vehicles", "horsepower", "TEXT"),                      # 马力
        ("vehicles", "box_type", "TEXT"),                        # 厢型
        ("vehicles", "box_dimension", "TEXT"),                   # 厢尺寸
        ("vehicles", "gear_position", "TEXT"),                   # 档位
        ("vehicles", "tailgate", "TEXT"),                        # 尾板
        ("vehicles", "battery_brand", "TEXT"),                   # 电池品牌
        ("vehicles", "product_code", "TEXT"),                    # 产品代码
        ("vehicles", "dealer_price", "REAL DEFAULT 0"),          # 网员价
        ("vehicles", "cab_type", "TEXT"),                        # 驾驶室
        ("vehicles", "other_config", "TEXT"),                    # 其他
        ("vehicles", "fuel_form", "TEXT"),                       # 燃料形式
        ("vehicles", "cab_style", "TEXT"),                       # 驾驶室类型
        ("vehicles", "engine_spec", "TEXT"),                     # 发动机厂家及功率
        ("vehicles", "gearbox_spec", "TEXT"),                    # 变速箱厂家及型号
        ("vehicles", "drive_motor_model", "TEXT"),               # 驱动电机型号
        ("vehicles", "battery_model", "TEXT"),                   # 新能源动力电池型号
        ("vehicles", "suspension_model", "TEXT"),                # 悬架型号
        # === 新车管入库信息.xlsx 维度的扩展列（79列模板）===
        # === 销售报单：尾板（有/无）===
        ("sales_orders", "tail_plate", "TEXT"),
        # === 销售报单：客户身份证号 ===
        ("sales_orders", "customer_id_card", "TEXT"),
        # === 车辆字典校验状态（车型等字段 vs 数据字典）===
        ("vehicles", "validation_status", "TEXT DEFAULT 'valid'"),   # valid / invalid / warning
        ("vehicles", "validation_message", "TEXT DEFAULT ''"),
        # === 软删除标记（老板端删除车辆后数据保留，前端不再展现）===
        ("vehicles", "is_deleted", "INTEGER DEFAULT 0"),
        # === 退车验车：新增随车工具/棚杆/洗车费/车体广告清洗/其他 ===
        ("return_inspections", "tool_kit", "INTEGER DEFAULT 0"),       # 随车工具
        ("return_inspections", "tent_pole", "INTEGER DEFAULT 0"),       # 棚杆
        ("return_inspections", "car_wash_fee", "INTEGER DEFAULT 0"),    # 洗车费
        ("return_inspections", "body_ad_clean", "INTEGER DEFAULT 0"),   # 车体广告清洗
        ("return_inspections", "other_info", "TEXT"),                  # 其他
        ("receivables", "screenshot_path", "TEXT"),
        ("receivables", "bank_serial", "TEXT"),
        ("receivables", "verified_by", "TEXT"),
        ("receivables", "verified_at", "TEXT"),
        # === 20260804 大改版：首付前移 / 合并审批 / 金融方案 / 驳回退款 ===
        ("sales_orders", "customer_screenshot_path", "TEXT"),
        ("sales_orders", "first_payment_received_amount", "REAL DEFAULT 0"),
        ("sales_orders", "first_payment_shortage_amount", "REAL DEFAULT 0"),
        ("sales_orders", "first_payment_shortage_reason", "TEXT"),
        ("sales_orders", "first_payment_promised_date", "TEXT"),
        ("sales_orders", "first_payment_check_status", "TEXT DEFAULT '未校验'"),
        ("sales_orders", "order_exception_reason", "TEXT"),
        ("sales_orders", "finance_plan_id", "INTEGER"),
        ("sales_orders", "snapshot_finance_plan", "TEXT"),
        ("sales_orders", "snapshot_lease_deposit_guidance", "REAL DEFAULT 0"),
        ("sales_orders", "snapshot_box_monthly_guidance", "REAL DEFAULT 0"),
        ("sales_orders", "refund_id", "INTEGER"),
        ("model_guidance_prices", "lease_deposit_guidance", "REAL DEFAULT 0"),
        ("model_guidance_prices", "box_standard_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "box_wide_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "box_high_rail_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "box_refrigerated_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "box_flatbed_price", "REAL DEFAULT 0"),
        ("model_guidance_prices", "tail_plate_price", "REAL DEFAULT 0"),
        ("contracts", "snapshot_finance_plan", "TEXT"),
        ("contracts", "snapshot_lease_deposit_guidance", "REAL DEFAULT 0"),
        ("contracts", "snapshot_box_monthly_guidance", "REAL DEFAULT 0"),
        ("receivables", "sales_order_id", "INTEGER"),
        # === SKU 改造：金融方案生效/失效日期（未来生效时间 + 有效期控制）===
        ("finance_plans", "effective_date", "TEXT"),
        ("finance_plans", "expiry_date", "TEXT"),
        # === 以租代售方案按新车子型号（基准车型 + 成色 + 厢型）维护 ===
        ("finance_plans", "condition", "TEXT DEFAULT '新车'"),
        ("finance_plans", "box_type", "TEXT DEFAULT ''"),
        # === SKU 改造：报单多车（JSON 数组 vehicle_ids，主车仍写 vehicle_id）===
        ("sales_orders", "vehicle_ids", "TEXT"),
        # === 财务确认报单凭证 ===
        ("sales_orders", "finance_bank_serial", "TEXT"),
        ("sales_orders", "finance_bank_receipt_path", "TEXT"),
        # === 催收闭环 ===
        ("urge_records", "evidence_path", "TEXT"),
        ("urge_records", "promised_repay_date", "TEXT"),
        ("urge_records", "completed_at", "TEXT"),
        ("urge_records", "closed_by", "TEXT"),
        # === 退车驳回/重提闭环 ===
        ("return_inspections", "rejected_by", "TEXT"),
        ("return_inspections", "rejected_at", "TEXT"),
        ("return_inspections", "reject_reason", "TEXT"),
        ("return_inspections", "resubmitted_by", "TEXT"),
        ("return_inspections", "resubmitted_at", "TEXT"),
        ("return_inspections", "resubmit_note", "TEXT"),
        # === 退车维修完工留痕 ===
        ("return_inspections", "repair_completed_by", "TEXT"),
        ("return_inspections", "repair_completed_at", "TEXT"),
        ("return_inspections", "repair_completion_note", "TEXT"),
        ("return_inspections", "repair_cost", "REAL DEFAULT 0"),
        ("ownership_transfers", "completed_by", "TEXT"),
        ("ownership_transfers", "completed_at", "TEXT"),
    ]
    for table, col, col_type in safe_alter_columns:
        try:
            c.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}")
        except Exception:
            pass

    # 历史退车单此前只把车辆状态改回在库/待维修，遗漏了成色变更。
    # 租赁车辆完成退车后再次入库一律按二手车管理；兼容旧库中的“已入库”状态。
    c.execute("""
        UPDATE vehicles
        SET condition='二手车'
        WHERE id IN (
            SELECT DISTINCT vehicle_id
            FROM return_inspections
            WHERE vehicle_id IS NOT NULL
              AND status IN ('已完成', '已入库')
        )
          AND COALESCE(condition, '') <> '二手车'
    """)

    # 旧流程在车管验车阶段就把需维修车辆改为“待维修”，导致退车尚未结算
    # 就从资产视图中脱离退车流程。未完结退车单应保持“退车中”。
    c.execute("""
        UPDATE vehicles
        SET status='退车中'
        WHERE status='待维修'
          AND id IN (
              SELECT DISTINCT vehicle_id
              FROM return_inspections
              WHERE vehicle_id IS NOT NULL
                AND needs_repair=1
                AND status NOT IN ('已完成', '已入库')
          )
    """)

    # 指导价 upsert 以车型和成色为自然键。旧库可能有重复/空成色数据，先归一化并去重。
    c.execute("""
        UPDATE model_guidance_prices
        SET is_new='新车'
        WHERE is_new IS NULL OR TRIM(is_new)=''
    """)
    c.execute("""
        DELETE FROM model_guidance_prices
        WHERE id NOT IN (
            SELECT keep_id
            FROM (
                SELECT MAX(id) AS keep_id
                FROM model_guidance_prices
                GROUP BY car_type, COALESCE(is_new, '新车')
            ) AS guidance_price_keep
        )
    """)
    c.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_model_guidance_prices_car_type_is_new
        ON model_guidance_prices(car_type, is_new)
    """)
    c.execute("""
        CREATE INDEX IF NOT EXISTS idx_finance_plans_scope
        ON finance_plans(car_type, condition, box_type, status)
    """)
    c.execute("""
        DELETE FROM urge_records
        WHERE id NOT IN (
            SELECT keep_id
            FROM (
                SELECT MAX(id) AS keep_id
                FROM urge_records
                GROUP BY repayment_id, urge_day
            ) AS urge_record_keep
        )
    """)
    c.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_urge_records_repayment_day
        ON urge_records(repayment_id, urge_day)
    """)

    # 数据字典以“维度 + 值 + 适用能源类型”为自然键。早期表缺少唯一约束，
    # 每次启动灌入种子数据都会重复新增，先规范旧数据再建立约束。
    c.execute("""
        UPDATE data_dictionaries
        SET category=TRIM(category),
            value=TRIM(value),
            energy_type=TRIM(COALESCE(energy_type, ''))
    """)
    c.execute("""
        UPDATE data_dictionaries AS keep
        SET status=CASE
            WHEN EXISTS (
                SELECT 1
                FROM data_dictionaries AS duplicate
                WHERE duplicate.category=keep.category
                  AND duplicate.value=keep.value
                  AND duplicate.energy_type=keep.energy_type
                  AND duplicate.status='启用'
            ) THEN '启用'
            ELSE keep.status
        END
        WHERE keep.id IN (
            SELECT keep_id
            FROM (
                SELECT MIN(id) AS keep_id
                FROM data_dictionaries
                GROUP BY category, value, energy_type
            ) AS dictionary_status_keep
        )
    """)
    c.execute("""
        DELETE FROM data_dictionaries
        WHERE id NOT IN (
            SELECT keep_id
            FROM (
                SELECT MIN(id) AS keep_id
                FROM data_dictionaries
                GROUP BY category, value, energy_type
            ) AS dictionary_keep
        )
    """)
    c.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_data_dictionaries_natural_key
        ON data_dictionaries(category, value, energy_type)
    """)

    if USE_MYSQL:
        _raw_c.execute("SET FOREIGN_KEY_CHECKS=1")
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
    default_users = [
        ('boss',   '123456', '王老板', '老板'),
        ('ops',    '123456', '李运营', '运营'),
        ('fin',    '123456', '张财务', '财务'),
        ('fleet',  '123456', '赵车管', '车管'),
        ('sales',  '123456', '周销售', '销售'),
    ]
    created_users = 0
    for u in default_users:
        c.execute("INSERT OR IGNORE INTO users (username, password, display_name, role) VALUES (?,?,?,?)", u)
        created_users += c.rowcount
    if created_users:
        conn.commit()
        print("Default users ensured.")
    c.execute("UPDATE users SET is_active=0 WHERE username='legal'")
    conn.commit()

    role_pages = {
        '老板': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'bills', 'reconciliation', 'risk', 'profit', 'settings'],
        '运营': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'reconciliation', 'risk', 'invoice'],
        '财务': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'bills', 'reconciliation', 'profit'],
        '车管': ['dashboard', 'assets', 'completion_history', 'approvals'],
        '销售': ['dashboard', 'orders', 'assets', 'completion_history', 'approvals', 'risk'],
    }
    role_actions = {
        '老板': ['*'],
        '运营': ['view_contracts', 'view_overdue', 'lock_vehicle', 'execute_lock', 'confirm_repayment', 'initiate_return', 'view_orders'],
        '财务': ['view_contracts', 'confirm_repayment', 'confirm_factory', 'view_bills', 'view_profit', 'upload_receipt', 'collect_payment', 'verify_return', 'upload_initial_receipt', 'activate_order'],
        '车管': ['add_vehicle', 'update_vehicle', 'return_inspect', 'deliver_vehicle'],
        '销售': ['create_contract', 'view_contracts', 'upload_screenshot', 'view_overdue', 'initiate_return', 'request_lock', 'initiate_initial_payment', 'create_order'],
    }
    hidden_fields = {
        '销售': {
            'vehicles': [
                'purchase_price', 'tax_rate', 'estimated_residual_value', 'guidance_price',

            ],
            'contracts': [
                'loan_amount', 'monthly_payment', 'factory_guarantee_deposit', 'paid_principal',
                'loan_balance', 'collected_deposit', 'collected_rent', 'expected_profit_floor',
                'expected_profit_ceiling', 'snapshot_guidance_price', 'snapshot_invoice_price',
            ],
            'sales_orders': [],
            'factory_repayments': ['amount'],
            'vehicle_rebates': ['*'],
            'profit': ['*'],
            'customer_blacklist': ['*'],
        },
        '运营': {
            'vehicles': [
                'purchase_price', 'tax_rate', 'guidance_price',

            ],
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
            'sales_orders': ['snapshot_guidance_price', 'snapshot_lease_installment_price'],
            'factory_repayments': ['amount'],
            'vehicle_rebates': ['*'],
            'profit': ['*'],
            'customer_blacklist': ['*'],
        },
    }
    for role, pages in role_pages.items():
        for page in pages:
            c.execute("INSERT OR IGNORE INTO role_pages (role, page_key) VALUES (?, ?)", (role, page))
    for role, actions in role_actions.items():
        for action in actions:
            c.execute("INSERT OR IGNORE INTO role_actions (role, action_key) VALUES (?, ?)", (role, action))
    for role, resources in hidden_fields.items():
        for resource, fields in resources.items():
            for field in fields:
                c.execute("""
                    INSERT OR IGNORE INTO role_field_permissions
                        (role, resource_key, field_key, can_view)
                    VALUES (?, ?, ?, 0)
                """, (role, resource, field))
    conn.commit()

    # (removed) # 只在 model_guidance_prices 为空时插入默认指导价
    # (removed) c.execute("SELECT COUNT(*) AS cnt FROM model_guidance_prices")
    # (removed) if c.fetchone()['cnt'] == 0:
        # (removed) default_model_guidance = [
            # (removed) ('解放轻卡4米2-虎6G140度纯电-宁德电池', 98000, 3500, 128000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡4米2-虎6G120度纯电-宁德电池', 98000, 3200, 120000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡-虎VR纯电-轻盈版', 90000, 3000, 115000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡4米2-虎6G 180混动-盟固利电池', 98000, 3300, 125000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡4米2-领途190马力', 98000, 3200, 125000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡4米2-领途150马力', 98000, 3000, 98000, 0.10, 0.025, 0.15, 0.025),
            # (removed) ('解放轻卡3米8-云内150排半', 90000, 2800, 90000, 0.10, 0.025, 0.15, 0.025),
        # (removed) ]
        # (removed) for car_type, legacy_price, lease_price, sale_price, lease_deposit_ratio, lease_repayment_ratio, sale_down_payment_ratio, sale_repayment_ratio in default_model_guidance:
            # (removed) c.execute("""
                # (removed) INSERT OR IGNORE INTO model_guidance_prices
                    # (removed) (car_type, guidance_price, lease_installment_price, sale_total_price,
                     # (removed) lease_deposit_ratio, lease_repayment_ratio, sale_down_payment_ratio, sale_repayment_ratio,
                     # (removed) remark, updated_by, updated_at)
                # (removed) VALUES (?, ?, ?, ?, ?, ?, ?, ?, '系统默认车型指导口径', '系统', datetime('now','localtime'))
            # (removed) """, (
                # (removed) car_type, legacy_price, lease_price, sale_price,
                # (removed) lease_deposit_ratio, lease_repayment_ratio, sale_down_payment_ratio, sale_repayment_ratio,
            # (removed) ))
        # (removed) conn.commit()

    c.execute("""
        INSERT OR IGNORE INTO receiving_companies
            (company_name, bank_name, bank_account_no, tax_no, status, remark, updated_by, updated_at)
        VALUES (?, ?, ?, ?, '启用', '系统默认收款主体', '系统', datetime('now','localtime'))
    """, (
        '陕西金聚源汽车服务有限公司',
        '中国建设银行西安分行',
        '6227000012345678901',
        '',
    ))
    conn.commit()

    # 不再插入测试车辆/合同/还款数据 —— 全新系统
    conn.close()
    print("Seed data inserted successfully (users & permissions only).")


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
