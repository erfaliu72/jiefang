#!/usr/bin/env python3
"""生成金聚源车辆管理系统产品功能清单 Word 文档"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

# ── 全局样式 ──
style = doc.styles['Normal']
font = style.font
font.name = '微软雅黑'
font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# ── 封面 ──
for _ in range(6):
    doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('金聚源车辆管理系统')
run.font.size = Pt(28)
run.bold = True
run.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('产品功能清单')
run.font.size = Pt(22)
run.font.color.rgb = RGBColor(0x2D, 0x5F, 0x8A)

doc.add_paragraph()
doc.add_paragraph()

version_info = [
    ('文档版本', 'v1.0'),
    ('功能描述文档基线', '功能描述文档-5.27-v3.md'),
    ('生成日期', datetime.date.today().strftime('%Y-%m-%d')),
    ('API 路由数', '106 个'),
    ('功能模块', '24 个'),
    ('功能项总计', '105 项'),
]
table = doc.add_table(rows=len(version_info), cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, (k, v) in enumerate(version_info):
    cell_k = table.rows[i].cells[0]
    cell_v = table.rows[i].cells[1]
    cell_k.text = k
    cell_v.text = v
    for cell in [cell_k, cell_v]:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(12)

doc.add_page_break()

# ── 辅助函数 ──
def set_cell_shading(cell, color):
    """设置单元格底色"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)

def make_module_header(doc, module_id, module_name):
    """添加模块标题"""
    p = doc.add_paragraph()
    p.space_before = Pt(16)
    p.space_after = Pt(6)
    run = p.add_run(f'模块 {module_id} — {module_name}')
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x1A, 0x3C, 0x6E)
    return p

def make_feature_table(doc, features, col_widths=None):
    """功能表格: 编号 | 功能 | 说明 | 状态"""
    table = doc.add_table(rows=1 + len(features), cols=4)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 表头
    headers = ['编号', '功能名称', '功能说明', '状态']
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_shading(cell, '1A3C6E')

    # 数据行
    for i, (fid, name, desc, status) in enumerate(features):
        row = table.rows[1 + i]
        row.cells[0].text = fid
        row.cells[1].text = name
        row.cells[2].text = desc
        row.cells[3].text = status
        for j in range(4):
            cell = row.cells[j]
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
            if j == 0:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if j == 3:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if status == '✅ 已实现':
                    for r in cell.paragraphs[0].runs:
                        r.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)
                elif status == '⚠️ 部分完善':
                    for r in cell.paragraphs[0].runs:
                        r.font.color.rgb = RGBColor(0xE6, 0x7E, 0x22)

    # 列宽
    if col_widths is None:
        col_widths = [Cm(2.0), Cm(4.5), Cm(8.5), Cm(2.5)]
    for i_row, row in enumerate(table.rows):
        for j, width in enumerate(col_widths):
            row.cells[j].width = width

    doc.add_paragraph()
    return table


# ── 一、概述 ──
h = doc.add_heading('一、概述', level=1)
doc.add_paragraph(
    '本文档为金聚源车辆管理系统的完整产品功能清单，覆盖系统全部 24 个功能模块、105 项具体功能。'
    '清单基于功能描述文档 v3.0（2026-05-29 修订版）与实际代码实现（106 个 API 路由）整理。'
)
doc.add_paragraph('角色说明：')
roles = [
    'boss — 老板',
    'sales — 销售',
    'fin — 财务',
    'ops — 运营',
    'fleet — 车管',
]
for r in roles:
    doc.add_paragraph(r, style='List Bullet')

doc.add_page_break()

# ── 二、功能清单 ──
h = doc.add_heading('二、功能清单', level=1)

# A — 用户与权限
make_module_header(doc, 'A', '用户与权限')
features_a = [
    ('A1', '登录 / 登出', '用户名+密码登录，5 角色（boss/sales/fin/ops/fleet），密码统一校验', '✅ 已实现'),
    ('A2', '角色矩阵', '按角色控制页面权限（role_pages）和操作权限（role_actions）', '✅ 已实现'),
    ('A3', '操作审计日志', '关键状态变更操作写入 audit_logs(action, target_type, target_id, operator)', '✅ 已实现'),
    ('A4', '字段级权限', '角色→字段→操作三维权限矩阵，前端组件+后端 DTO 双兜底（详见模块 U）', '✅ 已实现'),
]
make_feature_table(doc, features_a)

# B — 车辆资产管理
make_module_header(doc, 'B', '车辆资产管理')
features_b = [
    ('B1', '车辆入库', 'VIN(17位)/车牌/车型/采购价/开票价/税率/保险到期/年检到期/附件', '✅ 已实现'),
    ('B2', '车辆状态机', '在库→报单锁定中→租赁中/以租代售→退车中→待维修/在库→已售/已过户', '✅ 已实现'),
    ('B3', '车辆查询', 'VIN/车牌/车型/状态/公司/车厢/新车筛选，角色级字段可见性', '✅ 已实现'),
    ('B4', '车辆导入', 'Excel 批量导入车辆信息', '✅ 已实现'),
    ('B5', '车辆编辑/删除', '修改车辆信息或删除车辆', '✅ 已实现'),
    ('B6', '维修登记', '事故维修开始/完成状态切换（维修中→租赁中/以租代售）', '✅ 已实现'),
]
make_feature_table(doc, features_b)

# C — 指导价管理
make_module_header(doc, 'C', '指导价管理（双价+版本）')
features_c = [
    ('C1', '租赁分期指导价维护', '按车型维护 lease_installment_price，旧值入历史表，存量订单不变价', '✅ 已实现'),
    ('C2', '以租代售整车指导价维护', '同上 sale_total_price，price_kind=sale_total', '✅ 已实现'),
    ('C3', '指导价缺失硬性阻塞', '缺失时禁止提交报单；老板工作台红色提醒卡片', '✅ 已实现'),
    ('C4', '指导价历史查询', '历史变更记录可查', '✅ 已实现'),
    ('C5', '指导价导入', 'Excel 批量导入指导价', '✅ 已实现'),
]
make_feature_table(doc, features_c)

# D — 销售报单
make_module_header(doc, 'D', '销售报单（租赁/以租代售）')
features_d = [
    ('D1', '新建报单', '租赁/以租代售两种业务模式，VIN/客户/收款公司/起租日等', '✅ 已实现'),
    ('D2', '草稿保存', '7 天有效，不锁车，可反复编辑删除', '✅ 已实现'),
    ('D3', '报单提交即锁车', '提交后 vehicles.status→报单锁定中，黑名单校验', '✅ 已实现'),
    ('D4', '报单作废', '驳回/撤回/作废统一到已作废，释放车辆', '✅ 已实现'),
    ('D5', '报单查询', '状态/销售方式/顾问/客户/时间多维度筛选', '✅ 已实现'),
    ('D6', '草稿复制', '从已作废报单复制为草稿，修改后重新提交', '✅ 已实现'),
    ('D7', '报单编辑', '草稿及待审批状态可编辑修改', '✅ 已实现'),
]
make_feature_table(doc, features_d)

# E — 价格特批
make_module_header(doc, 'E', '价格特批')
features_e = [
    ('E1', '老板价格特批', '报价低于指导价自动触发特批流，24h SLA 超时提醒', '✅ 已实现'),
    ('E2', '财务确认报单', 'F1~F3 双表比对完成后激活报单（已激活），触发盈利数据生成', '✅ 已实现'),
    ('E3', '审批中心统一入口', '价格特批/退车/减免等各类型审批集中展示', '✅ 已实现'),
]
make_feature_table(doc, features_e)

# F — 还款计划
make_module_header(doc, 'F', '还款计划（双表生成与比对）')
features_f = [
    ('F1', '系统生成客户还款计划', '自动写入 repayments 表，状态=未激活（隔离 H1 扫描）', '✅ 已实现'),
    ('F2', '财务上传厂家还款计划', 'Excel 解析导入 factory_repayments', '✅ 已实现'),
    ('F3', '双表比对', '总额比对+首末期日期比对，差异行高亮', '✅ 已实现'),
    ('F4', '差异确认', '财务填写差异说明后可放行进入 E2', '✅ 已实现'),
]
make_feature_table(doc, features_f)

# G — 合同与出库
make_module_header(doc, 'G', '合同与出库')
features_g = [
    ('G1', '合同创建/上传 PDF', '线下签合同后上传 PDF，delivery_status=待出库', '✅ 已实现'),
    ('G2', '首款支付登记与审批', '运营登记→财务上传回单→审批通过，自动写回 contracts 字段', '✅ 已实现'),
    ('G3', '车辆出库', '照片≥3 张（车头/车尾/仪表），按业务类型区分出库门禁', '✅ 已实现'),
    ('G4', '出库文件上传', '出库照片/出库单附件上传', '✅ 已实现'),
]
make_feature_table(doc, features_g)

# H — 每日监控与催收
make_module_header(doc, 'H', '每日监控与催收')
features_h = [
    ('H1', '日终批处理', '每日 00:30 cron 执行 diff 区间式状态机扫描', '✅ 已实现'),
    ('H2', '催款执行', 'T+3 运营催款 + T+7 销售催款，截图必填，催款记录留存', '✅ 已实现'),
    ('H3', '顺序扣款（抵冲）', '租赁用押金池抵冲；以租代售无抵冲池，仅计提万5+禁止过户', '✅ 已实现'),
    ('H4', '锁车申请与审批', '建议锁车红标→申请→运营审核→老板审批→销售线下确认', '✅ 已实现'),
    ('H5', '解锁申请', '欠款还清后发起解锁流程', '✅ 已实现'),
]
make_feature_table(doc, features_h)

# I — 还款确认
make_module_header(doc, 'I', '还款确认与计划表调整')
features_i = [
    ('I1', '运营登记本期还款', '期数/实收金额/客户截图必填，多还抵扣月份与客户确认留痕', '✅ 已实现'),
    ('I2', '财务核销', '等额/少还（差额顺延）/多还（先冲应收再预抵）三分支', '✅ 已实现'),
    ('I3', '计划表展示', '双表两栏对照+预抵灰色徽标+差额顺延箭头', '✅ 已实现'),
    ('I4', '核销记录查询', '分配明细全量追溯', '✅ 已实现'),
]
make_feature_table(doc, features_i)

# J — 滞纳金
make_module_header(doc, 'J', '滞纳金')
features_j = [
    ('J1', '每日计提', '万分之5/日，基于原始未还本金无复利，UNIQUE防重', '✅ 已实现'),
    ('J2', '抵冲与收取', 'H3 中优先级第4位；客户主动还款时自动抵冲', '✅ 已实现'),
    ('J3', '减免审批', '销售发起→老板审批→财务执行三步，共用 waivers 表', '✅ 已实现'),
    ('J4', '结算时机', '退车/结清前自动汇总滞纳金，减免前置完成', '✅ 已实现'),
    ('J5', '滞纳金展示', '合同详情页包含本月新增/累计/已抵冲/已减免/待收卡片', '✅ 已实现'),
]
make_feature_table(doc, features_j)

# K — 租金减免
make_module_header(doc, 'K', '租金减免')
features_k = [
    ('K1', '销售发起申请', '目标期数/减免金额/原因，共用 waivers 表 waiver_kind=rent', '✅ 已实现'),
    ('K2', '老板审批', '48h SLA，销售可催办', '✅ 已实现'),
    ('K3', '财务复核执行', '更新 repayments.amount 并写核销分配记录', '✅ 已实现'),
    ('K4', '撤销', '仅已通过未生效可撤销，已生效需反向补登', '✅ 已实现'),
]
make_feature_table(doc, features_k)

# L — 发票申请
make_module_header(doc, 'L', '发票申请')
features_l = [
    ('L1', '运营发起申请', '开票主体与合同主体解耦，大额发票(≥5万)需老板审批', '✅ 已实现'),
    ('L2', '财务开票', '发票号≥8位+开票时间必填，开票主体按L1申请为准', '✅ 已实现'),
    ('L3', '发票纠错', '作废（本月内）/红冲（跨月）+重开完整流程', '✅ 已实现'),
    ('L4', '客户对账', '有效发票展示+历史版本链（蓝色/灰色/橙色徽标）', '⚠️ 部分完善'),
]
make_feature_table(doc, features_l)

# L-附
make_module_header(doc, 'L-附', '对账与上传规则（全局）')
features_la = [
    ('L-附1', '流水号幂等键', '(bank_serial, contract_id, period) 复合 UNIQUE', '✅ 已实现'),
    ('L-附2', '截图规则', '运营/销售侧必填，财务侧选填', '✅ 已实现'),
    ('L-附3', '字段必填校验', '前端禁用+后端400兜底', '✅ 已实现'),
]
make_feature_table(doc, features_la)

# M — 退车
make_module_header(doc, 'M', '退车（到期/提前）')
features_m = [
    ('M1', '销售发起退车', '到期/提前退车/转购车，提前退车违约金输入', '✅ 已实现'),
    ('M2', '车管验车', '随车工具/证件/公里数/照片，可标记待维修', '✅ 已实现'),
    ('M3', '运营查数据填单', '违章/ETC/事故扣款，滞纳金只读汇总', '✅ 已实现'),
    ('M4', '财务复核押金', '到期退款/提前退车违约金扣除', '✅ 已实现'),
    ('M5', '领导审批', '老板审批（approval_flows.ref_type=return_stock）', '✅ 已实现'),
    ('M6', '出款+入库', '银行流水号必填，车辆回库', '✅ 已实现'),
]
make_feature_table(doc, features_m)

# N — 过户
make_module_header(doc, 'N', '过户（以租代售结清/提前结清）')
features_n = [
    ('N1', '自然结清', '以租代售末期核销后自动创建过户单（contract_type=销售守卫）', '✅ 已实现'),
    ('N2', '提前结清买断', '剩余本金计算公式，运营报价→客户支付→财务确认', '✅ 已实现'),
    ('N3', '过户登记', '过户日期/新车主/材料附件', '✅ 已实现'),
    ('N4', '过户完成', '车辆→已售/已过户，合同→已结清/已提前结清', '✅ 已实现'),
]
make_feature_table(doc, features_n)

# O — 车辆返利
make_module_header(doc, 'O', '车辆返利')
features_o = [
    ('O1', '财务录入返利', '厂家返利/区域返利/其他，VIN/合同关联', '✅ 已实现'),
    ('O2', '可见范围控制', '仅财务/老板可见，非授权角色 403', '✅ 已实现'),
    ('O3', '返利报表', '按月/车型/车辆维度汇总，计入公司报表', '✅ 已实现'),
]
make_feature_table(doc, features_o)

# P — 单车盈利
make_module_header(doc, 'P', '单车盈利统计')
features_p = [
    ('P1', '基础数据生成', 'E2 激活后自动触发', '✅ 已实现'),
    ('P2', '财务/老板统计', '单车/客户/车型/合同多维度：计划收入/已收/未收/毛利/净利', '✅ 已实现'),
    ('P3', '销售/运营视图', '仅见客户视角数据（计划收入/已收/未收/减免明细），不见盈利', '✅ 已实现'),
]
make_feature_table(doc, features_p)

# Q — 老板穿透看板
make_module_header(doc, 'Q', '老板穿透看板')
features_q = [
    ('Q1', '总览卡', '在租/在售/在库/待维修+当月回款/到账/逾期/滞纳金', '✅ 已实现'),
    ('Q2', '待办提醒', '特批/退车审批/减免审批/指导价缺失/黑名单待审核', '✅ 已实现'),
    ('Q3', '穿透下钻', '报单→合同→还款→催收→退车→过户→盈利→返利全链路', '✅ 已实现'),
    ('Q4', 'SLA 超时提醒', '价格特批/减免/退车/开票/催款等 SLA 超时看板徽标', '✅ 已实现'),
]
make_feature_table(doc, features_q)

# R — 收款公司
make_module_header(doc, 'R', '收款公司主数据')
features_r = [
    ('R1', '收款公司维护', '全称/简称/信用代码/银行/账号/开户行/纳税人识别号/是否启用', '✅ 已实现'),
    ('R2', '联动', '报单/提醒文案/发票/提前结清下拉引用', '✅ 已实现'),
    ('R3', '审计', '字段变更写 audit_logs', '✅ 已实现'),
]
make_feature_table(doc, features_r)

# S — 客户黑名单
make_module_header(doc, 'S', '客户黑名单')
features_s = [
    ('S1', '录入', '客户姓名/身份证/电话/原因/等级（警示/拒绝）', '✅ 已实现'),
    ('S2', '审核与解禁', '老板解禁需填原因，记录保留供追溯', '✅ 已实现'),
    ('S3', '联动', '报单提交时校验，拒绝等级直接阻断提交', '✅ 已实现'),
]
make_feature_table(doc, features_s)

# U — 字段级权限
make_module_header(doc, 'U', '字段级权限矩阵')
features_u = [
    ('U1', '权限矩阵框架', '角色→字段→操作三维矩阵，默认拒绝显式授权', '✅ 已实现'),
    ('U2', '关键字段权限控制', '采购价/税率/返利/毛利/净利仅财务老板可见', '✅ 已实现'),
    ('U3', '前后端双校验', '后端 @require_field_perm 装饰器 + 前端 FieldGuard 组件', '✅ 已实现'),
]
make_feature_table(doc, features_u)

# V — 非功能要求
make_module_header(doc, 'V', '非功能要求（幂等/SLA/时间口径）')
features_v = [
    ('V1', '幂等性', '报单/核销/结清/开票/滞纳金等核心写操作幂等键保护', '✅ 已实现'),
    ('V2', 'SLA 超时提醒矩阵', '10+ 业务节点超时阈值与逐级升级处理', '✅ 已实现'),
    ('V3', '时间口径统一', '自然月/应还日锚/核销时间锚，跨月快照', '✅ 已实现'),
    ('V4', '缓存与性能', '看板首屏<2s，聚合查询5分钟缓存', '✅ 已实现'),
    ('V5', '数据备份与审计', '每日 02:00 全量备份保留30日，审计日志不可删', '✅ 已实现'),
]
make_feature_table(doc, features_v)

# 审批中心（跨模块）
make_module_header(doc, 'AC', '审批中心（跨模块）')
features_ac = [
    ('AC1', '统一审批列表', '价格特批/退车/减免/锁车/出库等各类型待审批集中展示', '✅ 已实现'),
    ('AC2', '审批通过/驳回', '通用审批/驳回接口，操作写审计日志', '✅ 已实现'),
    ('AC3', '重新提交', '驳回后可修改重新提审', '✅ 已实现'),
    ('AC4', '红点醒目提示', '待审批事项数量计数提示', '✅ 已实现'),
]
make_feature_table(doc, features_ac)

# 风险预警
make_module_header(doc, 'RS', '风险预警')
features_rs = [
    ('RS1', '逾期风险', '逾期合同及天数列表', '✅ 已实现'),
    ('RS2', '厂家逾期', '厂家计划逾期提醒', '✅ 已实现'),
    ('RS3', '保险/年审到期预警', '即将到期车辆预警列表', '✅ 已实现'),
]
make_feature_table(doc, features_rs)

doc.add_page_break()

# ── 三、统计汇总 ──
h = doc.add_heading('三、统计汇总', level=1)

summary_data = [
    ('A 用户与权限', '4', '✅ 完整'),
    ('B 车辆资产管理', '6', '✅ 完整'),
    ('C 指导价管理', '5', '✅ 完整'),
    ('D 销售报单', '7', '✅ 完整'),
    ('E 价格特批', '3', '✅ 完整'),
    ('F 还款计划', '4', '✅ 完整'),
    ('G 合同与出库', '4', '✅ 完整'),
    ('H 每日监控与催收', '5', '✅ 完整'),
    ('I 还款确认', '4', '✅ 完整'),
    ('J 滞纳金', '5', '✅ 完整'),
    ('K 租金减免', '4', '✅ 完整'),
    ('L 发票申请', '4', '✅ 完整'),
    ('L-附 对账规则', '3', '✅ 完整'),
    ('M 退车流程', '6', '✅ 完整'),
    ('N 过户', '4', '✅ 完整'),
    ('O 车辆返利', '3', '✅ 完整'),
    ('P 单车盈利', '3', '✅ 完整'),
    ('Q 老板看板', '4', '✅ 完整'),
    ('R 收款公司', '3', '✅ 完整'),
    ('S 客户黑名单', '3', '✅ 完整'),
    ('U 字段权限', '3', '✅ 完整'),
    ('V 非功能要求', '5', '✅ 完整'),
    ('AC 审批中心', '4', '✅ 完整'),
    ('RS 风险预警', '3', '✅ 完整'),
]

table = doc.add_table(rows=1 + len(summary_data), cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ['模块', '功能数', '生命周期']
for j, h in enumerate(headers):
    cell = table.rows[0].cells[j]
    cell.text = h
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(10)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    set_cell_shading(cell, '1A3C6E')

total = 0
for i, (mod, count, status) in enumerate(summary_data):
    row = table.rows[1 + i]
    row.cells[0].text = mod
    row.cells[1].text = count
    row.cells[2].text = status
    total += int(count)
    for j in range(3):
        cell = row.cells[j]
        for p in cell.paragraphs:
            for r in p.runs:
                r.font.size = Pt(9)
        if j in (1, 2):
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if j == 2 and '完整' in status:
            for r in cell.paragraphs[0].runs:
                r.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)

# 合计行
row_total = table.add_row()
row_total.cells[0].text = '合计'
row_total.cells[1].text = str(total)
row_total.cells[2].text = '功能完整度 ~95%'
for j in range(3):
    cell = row_total.cells[j]
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.bold = True
            r.font.size = Pt(10)
    set_cell_shading(cell, 'E8F0FE')

doc.add_paragraph()
p = doc.add_paragraph()
p.add_run('备注：').bold = True
doc.add_paragraph('• 本文档基于 功能描述文档-5.27-v3.md 与实际 app.py 代码（106 个 API 路由）整理', style='List Bullet')
doc.add_paragraph('• 全部 24 个模块中，23 个模块功能完整，1 项（L4 客户对账）前端展示在优化中', style='List Bullet')
doc.add_paragraph('• 涵盖 5 类角色（boss/sales/fin/ops/fleet）的完整业务流程闭环', style='List Bullet')

# ── 保存 ──
output_path = '/Users/liuyuanchang/code/jiefang/金聚源车辆管理系统_产品功能清单.docx'
doc.save(output_path)
print(f'✅ 文档已生成: {output_path}')
