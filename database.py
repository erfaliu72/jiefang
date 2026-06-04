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
        engine_number TEXT,
        company TEXT,
        car_type TEXT,
        vehicle_category TEXT,
        vehicle_cab TEXT,
        vehicle_engine_battery TEXT,
        vehicle_power_battery TEXT,
        vehicle_gearbox TEXT,
        vehicle_color TEXT,
        vehicle_box_type TEXT,
        box_type_remark TEXT,
        is_new TEXT DEFAULT '新车',
        invoice_date TEXT,
        invoice_price REAL DEFAULT 0,
        purchase_price REAL DEFAULT 0,
        tax_rate REAL DEFAULT 0.13,
        estimated_residual_value REAL DEFAULT 0,
        guidance_price REAL DEFAULT 0,
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
        sales_mode TEXT NOT NULL,
        vehicle_id INTEGER,
        vin TEXT NOT NULL,
        car_type TEXT,
        vehicle_color TEXT,
        plate_number TEXT,
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
        car_type TEXT UNIQUE NOT NULL,
        guidance_price REAL DEFAULT 0,
        lease_installment_price REAL DEFAULT 0,
        sale_total_price REAL DEFAULT 0,
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
        expected_profit_ceiling REAL DEFAULT 999999999,
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

    # === 安全添加新列（如果表已存在但缺少新字段）===
    safe_alter_columns = [
        ("contracts", "snapshot_guidance_price", "REAL DEFAULT 0"),
        ("contracts", "snapshot_invoice_price", "REAL DEFAULT 0"),
        ("contracts", "customer_plan_match_status", "TEXT DEFAULT '未比对'"),
        ("contracts", "plan_compare_summary", "TEXT"),
        ("contracts", "expected_profit_floor", "REAL DEFAULT 0"),
        ("contracts", "expected_profit_ceiling", "REAL DEFAULT 999999999"),
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
        ("vehicles", "engine_number", "TEXT"),
        ("vehicles", "vehicle_category", "TEXT"),
        ("vehicles", "vehicle_cab", "TEXT"),
        ("vehicles", "vehicle_engine_battery", "TEXT"),
        ("vehicles", "vehicle_power_battery", "TEXT"),
        ("vehicles", "vehicle_gearbox", "TEXT"),
        ("vehicles", "vehicle_color", "TEXT"),
        ("vehicles", "vehicle_box_type", "TEXT"),
        ("vehicles", "box_type_remark", "TEXT"),
        ("vehicles", "invoice_contract_file", "TEXT"),
        ("vehicles", "purchase_price", "REAL DEFAULT 0"),
        ("vehicles", "tax_rate", "REAL DEFAULT 0.13"),
        ("vehicles", "estimated_residual_value", "REAL DEFAULT 0"),
        ("vehicles", "guidance_price", "REAL DEFAULT 0"),
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
        ("return_inspections", "refund_company_name", "TEXT"),
        ("return_inspections", "refund_bank_name", "TEXT"),
        ("return_inspections", "refund_bank_card_no", "TEXT"),
        ("return_inspections", "lease_bank_name", "TEXT"),
        ("return_inspections", "lease_bank_card_no", "TEXT"),
        ("return_inspections", "leader_remark", "TEXT"),
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
        ("model_guidance_price_history", "price_kind", "TEXT DEFAULT 'legacy'"),
        ("contract_initial_payments", "received_amount", "REAL DEFAULT 0"),
        ("contract_initial_payments", "bank_serial", "TEXT"),
        ("repayments", "paid_amount", "REAL DEFAULT 0"),
        ("repayments", "verified_amount", "REAL DEFAULT 0"),
        ("repayments", "screenshot_path", "TEXT"),
        ("repayments", "bank_receipt_path", "TEXT"),
        ("repayments", "bank_serial", "TEXT"),
        ("repayments", "verified_by", "TEXT"),
        ("repayments", "verified_at", "TEXT"),
        ("repayments", "waterfall_summary", "TEXT"),
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
        '老板': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return', 'profit', 'settings'],
        '运营': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'risk', 'return'],
        '财务': ['dashboard', 'orders', 'assets', 'approvals', 'bills', 'reconciliation', 'profit', 'return'],
        '车管': ['dashboard', 'assets', 'approvals', 'return'],
        '销售': ['dashboard', 'orders', 'assets', 'approvals', 'risk', 'return'],
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
            'vehicles': ['purchase_price', 'tax_rate', 'estimated_residual_value', 'guidance_price'],
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

    default_model_guidance = [
        ('解放轻卡4米2-虎6G140度纯电-宁德电池', 98000, 3500, 128000),
        ('解放轻卡4米2-虎6G120度纯电-宁德电池', 98000, 3200, 120000),
        ('解放轻卡-虎VR纯电-轻盈版', 90000, 3000, 115000),
        ('解放轻卡4米2-虎6G 180混动-盟固利电池', 98000, 3300, 125000),
        ('解放轻卡4米2-领途190马力', 98000, 3200, 125000),
        ('解放轻卡4米2-领途150马力', 98000, 3000, 98000),
        ('解放轻卡3米8-云内150排半', 90000, 2800, 90000),
    ]
    for car_type, legacy_price, lease_price, sale_price in default_model_guidance:
        c.execute("""
            INSERT OR IGNORE INTO model_guidance_prices
                (car_type, guidance_price, lease_installment_price, sale_total_price, remark, updated_by, updated_at)
            VALUES (?, ?, ?, ?, '系统默认车型指导价', '系统', datetime('now','localtime'))
        """, (car_type, legacy_price, lease_price, sale_price))
    conn.commit()

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
