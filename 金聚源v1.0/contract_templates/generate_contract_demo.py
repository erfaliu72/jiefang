from pathlib import Path
from datetime import datetime
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent
RAW_DIR = BASE_DIR / "raw"
OUT_DIR = BASE_DIR / "generated"
OUT_DIR.mkdir(parents=True, exist_ok=True)
pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))

TEMPLATES = {
    "销售": "【金聚源】车辆买卖合同（新车）-陈律师-20250905.docx",
    "租赁": "2026年  金聚源 车辆租赁合同.docx",
    "以租代售": "金聚源 （以租代购）-陈律师-20250905.docx",
}


def format_money(value):
    return f"{float(value):,.2f}"


def format_date_parts(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return str(dt.year), str(dt.month), str(dt.day)


def build_context(contract):
    start_y, start_m, start_d = format_date_parts(contract["start_date"])
    end_y, end_m, end_d = format_date_parts(contract["end_date"])
    total_days = str((datetime.strptime(contract["end_date"], "%Y-%m-%d") - datetime.strptime(contract["start_date"], "%Y-%m-%d")).days)
    return {
        "contract_no": contract["contract_no"],
        "customer_name": contract["customer_name"],
        "customer_phone": contract["customer_phone"],
        "customer_id_card": contract["customer_id_card"],
        "car_type": contract["car_type"],
        "plate_number": contract["plate_number"],
        "vin": contract["vin"],
        "yard": contract["yard"],
        "start_date": contract["start_date"],
        "end_date": contract["end_date"],
        "start_year": start_y,
        "start_month": start_m,
        "start_day": start_d,
        "end_year": end_y,
        "end_month": end_m,
        "end_day": end_d,
        "loan_periods": str(contract["loan_periods"]),
        "loan_period_years": f"{round(contract['loan_periods'] / 12, 2):g}",
        "total_days": total_days,
        "total_price": format_money(contract["total_price"]),
        "rent": format_money(contract["rent"]),
        "deposit": format_money(contract["deposit"]),
        "down_payment": format_money(contract["down_payment"]),
        "loan_amount": format_money(contract["loan_amount"]),
        "repayment_day": str(contract["repayment_day"]),
        "contract_type": contract["contract_type"],
    }


def replacements(contract):
    ctx = build_context(contract)
    if contract["contract_type"] == "销售":
        return [
            ("合同编号：【  】", f"合同编号：【{ctx['contract_no']}】"),
            ("乙方（买受人）：                  身份证号：", f"乙方（买受人）：{ctx['customer_name']}    身份证号：{ctx['customer_id_card']}"),
            ("电话                  ", f"电话 {ctx['customer_phone']}"),
            ("车辆品牌：  ，车型： ，车架号：", f"车辆品牌：解放，车型：{ctx['car_type']}，车架号：{ctx['vin']}"),
            ("2.1车辆含税金额为          元，税率为13%。", f"2.1车辆含税金额为 {ctx['total_price']} 元，税率为13%。"),
            ("按揭贷款方式付款：乙方应当于签订之日起     日内向甲方支付车辆首付款     元，余款     元乙方向相关贷款机构申请贷款支付。",
             f"按揭贷款方式付款：乙方应当于签订之日起 7 日内向甲方支付车辆首付款 {ctx['down_payment']} 元，余款 {ctx['loan_amount']} 元乙方向相关贷款机构申请贷款支付。"),
            ("分期付款：乙方应当于     年     月     日前分     期支付该车辆的全部价款，其中乙方应当于签订之日起     日内向甲方支付车辆首付款     元，剩余款项于每月     日前向甲方支付。",
             f"分期付款：乙方应当于 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日前分 {ctx['loan_periods']} 期支付该车辆的全部价款，其中乙方应当于签订之日起 7 日内向甲方支付车辆首付款 {ctx['down_payment']} 元，剩余款项于每月 {ctx['repayment_day']} 日前向甲方支付。"),
        ]
    if contract["contract_type"] == "租赁":
        return [
            ("合同编号：【2026032902】", f"合同编号：【{ctx['contract_no']}】"),
            ("乙方（承租人）：            身份证号：               电话", f"乙方（承租人）：{ctx['customer_name']}    身份证号：{ctx['customer_id_card']}    电话 {ctx['customer_phone']}"),
            ("车型：       ， 车牌号       ，车架号：                          。", f"车型：{ctx['car_type']}， 车牌号 {ctx['plate_number']}，车架号：{ctx['vin']}。"),
            ("2.1起算日：     年    月    日至   年   月   日（实际以租赁车辆实际交付乙方之日起算），共    年（    天）",
             f"2.1起算日：{ctx['start_year']} 年 {ctx['start_month']} 月 {ctx['start_day']} 日至 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日（实际以租赁车辆实际交付乙方之日起算），共 {ctx['loan_period_years']} 年（{ctx['total_days']} 天）"),
            ("3.1.1本合同租赁车辆的租金标准为      元/月。", f"3.1.1本合同租赁车辆的租金标准为 {ctx['rent']} 元/月。"),
            ("3.1.2乙方应当于甲方交付车辆前7日内向甲方交纳首期租金，剩余期限的租金乙方应当于每月【 】日前向甲方支付下一个支付周期的租金。",
             f"3.1.2乙方应当于甲方交付车辆前7日内向甲方交纳首期租金，剩余期限的租金乙方应当于每月【{ctx['repayment_day']}】日前向甲方支付下一个支付周期的租金。"),
            ("3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【    】元，作为乙方履行本合同的保证。",
             f"3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【{ctx['deposit']}】元，作为乙方履行本合同的保证。"),
        ]
    return [
        ("合同编号：【       】", f"合同编号：【{ctx['contract_no']}】"),
        ("乙方（承租人）：            ，身份证号：", f"乙方（承租人）：{ctx['customer_name']}，身份证号：{ctx['customer_id_card']}"),
        ("车型：          ， 车牌号         ，车架号：               。", f"车型：{ctx['car_type']}， 车牌号 {ctx['plate_number']}，车架号：{ctx['vin']}。"),
        ("起算日：      年     月     日至      年    月    日（实际以租赁车辆实际交付乙方之日起算），共/年（    个月）",
         f"起算日：{ctx['start_year']} 年 {ctx['start_month']} 月 {ctx['start_day']} 日至 {ctx['end_year']} 年 {ctx['end_month']} 月 {ctx['end_day']} 日（实际以租赁车辆实际交付乙方之日起算），共 {ctx['loan_period_years']} 年（{ctx['loan_periods']} 个月）"),
        ("3.1.1本合同租赁车辆的租金标准为租金标准为每月     元/月，租期       个月，合计       元。",
         f"3.1.1本合同租赁车辆的租金标准为每月 {ctx['rent']} 元/月，租期 {ctx['loan_periods']} 个月，合计 {ctx['total_price']} 元。"),
        ("3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【 】元，作为乙方履行本合同的保证。",
         f"3.2.1乙方应当于甲方交付车辆前7日内向甲方交纳合同保证金【{ctx['deposit']}】元，作为乙方履行本合同的保证。"),
        ("4.1甲方应当在签订本合同且收到乙方交付的首期租金及全额合同保证金后15日内在                         （车辆交付地点）将车辆交付给乙方，届时应当检测确认租赁车辆设备及租赁车辆状况，双方无异议后应签署《车辆交接单》（详见附件1）；乙方应签署《车辆交接单》，该清单签署后代表乙方已对车辆质量以及性能",
         f"4.1甲方应当在签订本合同且收到乙方交付的首期租金及全额合同保证金后15日内在 {ctx['yard']}（车辆交付地点）将车辆交付给乙方，届时应当检测确认租赁车辆设备及租赁车辆状况，双方无异议后应签署《车辆交接单》（详见附件1）；乙方应签署《车辆交接单》，该清单签署后代表乙方已对车辆质量以及性能"),
    ]


def replace_in_doc(doc, reps):
    for paragraph in doc.paragraphs:
        text = paragraph.text
        new_text = text
        for old, new in reps:
            new_text = new_text.replace(old, new)
        if new_text != text and paragraph.runs:
            paragraph.runs[0].text = new_text
            for run in paragraph.runs[1:]:
                run.text = ""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    text = paragraph.text
                    new_text = text
                    for old, new in reps:
                        new_text = new_text.replace(old, new)
                    if new_text != text and paragraph.runs:
                        paragraph.runs[0].text = new_text
                        for run in paragraph.runs[1:]:
                            run.text = ""


def render_pdf(contract, target):
    ctx = build_context(contract)
    c = canvas.Canvas(str(target), pagesize=A4)
    w, h = A4
    y = h - 50
    c.setFont("STSong-Light", 16)
    c.drawString(50, y, f"{ctx['contract_type']}合同")
    y -= 28
    c.setFont("STSong-Light", 11)
    lines = [
        f"合同编号：{ctx['contract_no']}",
        f"客户：{ctx['customer_name']}    电话：{ctx['customer_phone']}    身份证号：{ctx['customer_id_card']}",
        f"车型：{ctx['car_type']}    车牌号：{ctx['plate_number']}    VIN：{ctx['vin']}",
        f"合同起止：{ctx['start_date']} 至 {ctx['end_date']}",
        f"总价：¥{ctx['total_price']}    月租：¥{ctx['rent']}    押金：¥{ctx['deposit']}",
        f"首付：¥{ctx['down_payment']}    贷款金额：¥{ctx['loan_amount']}    期数：{ctx['loan_periods']}",
        f"说明：该 PDF 为系统基于合同模板自动生成的便捷版。"
    ]
    for raw in lines:
        for line in simpleSplit(raw, "STSong-Light", 11, w - 100):
            c.drawString(50, y, line)
            y -= 18
    c.save()


def generate(contract):
    template = RAW_DIR / TEMPLATES[contract["contract_type"]]
    doc = Document(str(template))
    replace_in_doc(doc, replacements(contract))
    base = f"{contract['contract_no']}_{contract['customer_name']}_{contract['contract_type']}"
    docx_target = OUT_DIR / f"{base}.docx"
    pdf_target = OUT_DIR / f"{base}.pdf"
    doc.save(str(docx_target))
    render_pdf(contract, pdf_target)
    return docx_target, pdf_target


if __name__ == "__main__":
    samples = [
        {
            "contract_type": "销售",
            "contract_no": "JGY-销售-000001",
            "customer_name": "张三",
            "customer_phone": "13800001111",
            "customer_id_card": "610102199001010011",
            "car_type": "虎VR单排厢货",
            "plate_number": "陕A12345",
            "vin": "LFNA4LC73TAE16802",
            "yard": "西安市未央区丰业大道88号国坤物流园内",
            "start_date": "2026-05-12",
            "end_date": "2026-08-12",
            "loan_periods": 3,
            "repayment_day": 15,
            "total_price": 128000,
            "rent": 0,
            "deposit": 5000,
            "down_payment": 30000,
            "loan_amount": 98000,
        },
        {
            "contract_type": "租赁",
            "contract_no": "JGY-租赁-000002",
            "customer_name": "李四",
            "customer_phone": "13900002222",
            "customer_id_card": "610102199002020022",
            "car_type": "解放轻卡4米2-领途150马力",
            "plate_number": "陕A23456",
            "vin": "T0000000000000001",
            "yard": "陕西金聚源",
            "start_date": "2026-05-12",
            "end_date": "2027-05-11",
            "loan_periods": 12,
            "repayment_day": 15,
            "total_price": 0,
            "rent": 3800,
            "deposit": 5000,
            "down_payment": 0,
            "loan_amount": 0,
        },
        {
            "contract_type": "以租代售",
            "contract_no": "JGY-以租代售-000003",
            "customer_name": "王五",
            "customer_phone": "13700003333",
            "customer_id_card": "610102199003030033",
            "car_type": "解放轻卡4米2-虎6G140度纯电-宁德电池",
            "plate_number": "陕A34567",
            "vin": "T0000000000000002",
            "yard": "陕西金聚源",
            "start_date": "2026-05-12",
            "end_date": "2029-05-11",
            "loan_periods": 36,
            "repayment_day": 15,
            "total_price": 136800,
            "rent": 4200,
            "deposit": 8000,
            "down_payment": 12000,
            "loan_amount": 98000,
        },
    ]
    for item in samples:
        word_file, pdf_file = generate(item)
        print(word_file)
        print(pdf_file)
