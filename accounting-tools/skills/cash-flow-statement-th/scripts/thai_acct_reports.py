#!/usr/bin/env python3
"""
thai_acct_reports.py — เครื่องมือสร้างรายงานบัญชี 5 รายงาน (ไฟล์ Excel) จากไฟล์ CSV

รายงานที่รองรับ
  cashflow   งบกระแสเงินสด (วิธีทางอ้อม)          input: financial_statements.csv
  ar-aging   รายงานอายุลูกหนี้                     input: ar_open_invoices.csv
  ap-aging   รายงานอายุเจ้าหนี้                    input: ap_open_invoices.csv
  vat        รายงานภาษีซื้อ (พร้อมตรวจความถูกต้อง)  input: purchase_vat_invoices.csv
  ratios     รายงานอัตราส่วนทางการเงิน             input: financial_statements.csv

ตัวอย่าง
  python thai_acct_reports.py cashflow  financial_statements.csv -o cashflow.xlsx
  python thai_acct_reports.py ar-aging  ar_open_invoices.csv --as-of 2026-10-31 --gl 2214900 -o ar.xlsx
  python thai_acct_reports.py ap-aging  ap_open_invoices.csv --as-of 2026-10-31 -o ap.xlsx
  python thai_acct_reports.py vat       purchase_vat_invoices.csv --month 2026-10 --company-tax-id 0999900001135 -o vat.xlsx
  python thai_acct_reports.py ratios    financial_statements.csv -o ratios.xlsx

ข้อกำหนดไฟล์ CSV: UTF-8 (มี/ไม่มี BOM ก็ได้), วันที่รูปแบบ YYYY-MM-DD (ค.ศ.), ตัวเลขไม่ต้องมีจุลภาค
ต้องการ: openpyxl
"""
import argparse
import csv
import sys
from collections import OrderedDict, defaultdict
from datetime import date, datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit("ต้องติดตั้ง openpyxl ก่อน: pip install openpyxl")

NUM = '#,##0.00;(#,##0.00);"-"'
FONT = "TH Sarabun New"
HEAD_FILL = PatternFill("solid", fgColor="1A3C34")
SUB_FILL = PatternFill("solid", fgColor="E8EEF2")
WARN_FILL = PatternFill("solid", fgColor="FDECEA")
THIN = Side(style="thin", color="B0B7BB")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TH_MONTHS = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
             "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]


# ---------------------------------------------------------------- utilities
def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return [{k.strip(): (v or "").strip() for k, v in r.items()} for r in csv.DictReader(f)]


def num(v):
    v = (v or "").replace(",", "").strip()
    if v in ("", "-"):
        return 0.0
    if v.startswith("(") and v.endswith(")"):
        return -float(v[1:-1])
    return float(v)


def d(v):
    return datetime.strptime(v.strip(), "%Y-%m-%d").date()


def th_date(x):
    return f"{x.day:02d}/{x.month:02d}/{x.year + 543}"


def r2(x):
    return round(x + 0.0, 2)


def valid_tax_id(tid):
    """ตรวจเลขประจำตัวผู้เสียภาษี 13 หลักด้วย check digit (mod 11)"""
    if len(tid) != 13 or not tid.isdigit():
        return False
    s = sum(int(tid[i]) * (13 - i) for i in range(12))
    return (11 - s % 11) % 10 == int(tid[12])


class Sheet:
    """ตัวช่วยเขียนชีตแบบเรียงบรรทัด"""

    def __init__(self, ws, widths):
        self.ws, self.row = ws, 1
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        self.ncol = len(widths)

    def title(self, *lines):
        for i, t in enumerate(lines):
            c = self.ws.cell(self.row, 1, t)
            c.font = Font(name=FONT, size=16 if i == 0 else 14, bold=i == 0)
            self.ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=self.ncol)
            c.alignment = Alignment(horizontal="center")
            self.row += 1
        self.row += 1

    def header(self, cols):
        for i, t in enumerate(cols, 1):
            c = self.ws.cell(self.row, i, t)
            c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
            c.fill, c.border = HEAD_FILL, BOX
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.ws.freeze_panes = self.ws.cell(self.row + 1, 1)
        self.row += 1

    def line(self, values, bold=False, fill=None, indent=0, box=True):
        for i, v in enumerate(values, 1):
            c = self.ws.cell(self.row, i, v)
            c.font = Font(name=FONT, size=14, bold=bold)
            if isinstance(v, (int, float)):
                c.number_format = NUM
            if fill:
                c.fill = fill
            if box:
                c.border = BOX
            if i == 1 and indent:
                c.alignment = Alignment(indent=indent)
        self.row += 1
        return self.row - 1

    def blank(self, n=1):
        self.row += n


def save(wb, out):
    wb.save(out)
    print(f"บันทึกไฟล์: {out}")


# ---------------------------------------------------------------- financial statements
def load_fs(path):
    rows = read_csv(path)
    years = sorted(int(k[1:]) for k in rows[0] if k.startswith("y") and k[1:].isdigit())
    fs = {}
    for r in rows:
        fs[r["code"]] = {"item": r["item"], **{y: (num(r[f"y{y}"]) if r.get(f"y{y}", "") != "" else None) for y in years}}
    return fs, years


def g(fs, code, y):
    v = fs.get(code, {}).get(y)
    return 0.0 if v is None else v


def compute_cashflow(fs, y0, y1):
    P = lambda c: g(fs, c, y1)
    O = lambda c: g(fs, c, y0)
    pbt = P("PL_REV") - P("PL_COGS") + P("PL_OTHINC") - P("PL_SGA") - P("PL_INT")
    net_profit = pbt - P("PL_TAX")
    op = OrderedDict()
    op["กำไรก่อนภาษีเงินได้"] = pbt
    adj = OrderedDict([
        ("ค่าเสื่อมราคา", P("SUP_DEP")),
        ("หนี้สงสัยจะสูญ", P("SUP_BADDEBT")),
        ("(กำไร) ขาดทุนจากการขายสินทรัพย์", -P("SUP_GAIN")),
        ("ต้นทุนทางการเงิน", P("PL_INT")),
    ])
    wc = OrderedDict([
        ("ลูกหนี้การค้า (ลดลง) เพิ่มขึ้น", -(P("BS_AR") - O("BS_AR"))),
        ("สินค้าคงเหลือ (ลดลง) เพิ่มขึ้น", -(P("BS_INV") - O("BS_INV"))),
        ("ค่าใช้จ่ายจ่ายล่วงหน้า (ลดลง) เพิ่มขึ้น", -(P("BS_PREPAID") - O("BS_PREPAID"))),
        ("เจ้าหนี้การค้า เพิ่มขึ้น (ลดลง)", P("BS_AP") - O("BS_AP")),
        ("ค่าใช้จ่ายค้างจ่าย เพิ่มขึ้น (ลดลง)", P("BS_ACCRUED") - O("BS_ACCRUED")),
    ])
    before_wc = pbt + sum(adj.values())
    cash_gen = before_wc + sum(wc.values())
    tax_paid = O("BS_TAXPAY") + P("PL_TAX") - P("BS_TAXPAY")
    paid = OrderedDict([("จ่ายดอกเบี้ย", -P("SUP_INT_PAID")), ("จ่ายภาษีเงินได้", -tax_paid)])
    cfo = cash_gen + sum(paid.values())
    inv = OrderedDict([
        ("เงินสดจ่ายซื้อที่ดิน อาคารและอุปกรณ์", -P("SUP_PPE_BUY")),
        ("เงินสดรับจากการขายอุปกรณ์", P("SUP_PPE_PROCEEDS")),
    ])
    cfi = sum(inv.values())
    fin = OrderedDict([
        ("เงินกู้ยืมระยะสั้น เพิ่มขึ้น (ลดลง)", P("BS_STLOAN") - O("BS_STLOAN")),
        ("เงินกู้ยืมระยะยาว เพิ่มขึ้น (ลดลง)", P("BS_LTLOAN") - O("BS_LTLOAN")),
        ("เงินสดรับจากการเพิ่มทุน", P("BS_CAPITAL") - O("BS_CAPITAL")),
        ("จ่ายเงินปันผล", -P("SUP_DIV")),
    ])
    cff = sum(fin.values())
    net = cfo + cfi + cff
    checks = []
    dcash = P("BS_CASH") - O("BS_CASH")
    checks.append(("เงินสดเพิ่มขึ้นสุทธิ = ผลต่างเงินสดในงบแสดงฐานะการเงิน", net, dcash))
    checks.append(("กำไรสะสมต้นงวด + กำไรสุทธิ − ปันผล = กำไรสะสมปลายงวด",
                   O("BS_RE") + net_profit - P("SUP_DIV"), P("BS_RE")))
    checks.append(("ค่าเผื่อหนี้สงสัยจะสูญเพิ่มขึ้น = หนี้สงสัยจะสูญ (ถ้าไม่มีตัดจำหน่ายหนี้สูญ)",
                   P("BS_ALLOW") - O("BS_ALLOW"), P("SUP_BADDEBT")))
    disp_cost = O("BS_PPE") + P("SUP_PPE_BUY") - P("BS_PPE")
    disp_acc = O("BS_ACCDEP") + P("SUP_DEP") - P("BS_ACCDEP")
    checks.append(("มูลค่าตามบัญชีสินทรัพย์ที่ขาย = เงินรับ − กำไรจากการขาย",
                   disp_cost - disp_acc, P("SUP_PPE_PROCEEDS") - P("SUP_GAIN")))
    return dict(op=op, adj=adj, wc=wc, before_wc=before_wc, cash_gen=cash_gen, paid=paid, cfo=cfo,
                inv=inv, cfi=cfi, fin=fin, cff=cff, net=net, open_cash=O("BS_CASH"),
                close_cash=P("BS_CASH"), checks=checks, net_profit=net_profit)


def report_cashflow(args):
    fs, years = load_fs(args.input)
    y0, y1 = years[-2], years[-1]
    cf = compute_cashflow(fs, y0, y1)
    wb = Workbook()
    s = Sheet(wb.active, [62, 22])
    wb.active.title = "งบกระแสเงินสด"
    s.title(args.company, "งบกระแสเงินสด (วิธีทางอ้อม)", f"สำหรับปีสิ้นสุดวันที่ 31 ธันวาคม {y1 + 543}", "(หน่วย: บาท)")
    s.header(["รายการ", f"ปี {y1 + 543}"])
    s.line(["กระแสเงินสดจากกิจกรรมดำเนินงาน", None], bold=True, fill=SUB_FILL)
    s.line(["กำไรก่อนภาษีเงินได้", r2(cf["op"]["กำไรก่อนภาษีเงินได้"])], indent=1)
    s.line(["รายการปรับกระทบกำไรก่อนภาษีเป็นเงินสดรับ (จ่าย)", None], indent=1)
    for k, v in cf["adj"].items():
        s.line([k, r2(v)], indent=2)
    s.line(["กำไรจากการดำเนินงานก่อนการเปลี่ยนแปลงในสินทรัพย์และหนี้สินดำเนินงาน", r2(cf["before_wc"])], bold=True)
    for k, v in cf["wc"].items():
        s.line([k, r2(v)], indent=2)
    s.line(["เงินสดรับจากการดำเนินงาน", r2(cf["cash_gen"])], bold=True)
    for k, v in cf["paid"].items():
        s.line([k, r2(v)], indent=2)
    s.line(["เงินสดสุทธิได้มาจาก (ใช้ไปใน) กิจกรรมดำเนินงาน", r2(cf["cfo"])], bold=True, fill=SUB_FILL)
    s.line(["กระแสเงินสดจากกิจกรรมลงทุน", None], bold=True, fill=SUB_FILL)
    for k, v in cf["inv"].items():
        s.line([k, r2(v)], indent=2)
    s.line(["เงินสดสุทธิได้มาจาก (ใช้ไปใน) กิจกรรมลงทุน", r2(cf["cfi"])], bold=True, fill=SUB_FILL)
    s.line(["กระแสเงินสดจากกิจกรรมจัดหาเงิน", None], bold=True, fill=SUB_FILL)
    for k, v in cf["fin"].items():
        s.line([k, r2(v)], indent=2)
    s.line(["เงินสดสุทธิได้มาจาก (ใช้ไปใน) กิจกรรมจัดหาเงิน", r2(cf["cff"])], bold=True, fill=SUB_FILL)
    s.line(["เงินสดและรายการเทียบเท่าเงินสดเพิ่มขึ้น (ลดลง) สุทธิ", r2(cf["net"])], bold=True)
    s.line([f"เงินสดและรายการเทียบเท่าเงินสด ณ วันต้นปี", r2(cf["open_cash"])])
    s.line([f"เงินสดและรายการเทียบเท่าเงินสด ณ วันสิ้นปี", r2(cf["open_cash"] + cf["net"])], bold=True)
    ws2 = wb.create_sheet("ตรวจสอบ")
    s2 = Sheet(ws2, [70, 20, 20, 16, 12])
    s2.title("การตรวจสอบความถูกต้อง (Cross-check)")
    s2.header(["รายการตรวจ", "คำนวณได้", "ตามงบ/ข้อมูล", "ผลต่าง", "ผล"])
    ok_all = True
    for name, a, b in cf["checks"]:
        diff = r2(a - b)
        ok = abs(diff) < 0.5
        ok_all &= ok
        s2.line([name, r2(a), r2(b), diff, "ผ่าน" if ok else "ไม่ผ่าน"], fill=None if ok else WARN_FILL)
    save(wb, args.output)
    print(f"CFO {cf['cfo']:,.2f} | CFI {cf['cfi']:,.2f} | CFF {cf['cff']:,.2f} | สุทธิ {cf['net']:,.2f}")
    print("ผลตรวจสอบ: " + ("ผ่านทุกข้อ" if ok_all else "มีรายการไม่ผ่าน — ดูชีต 'ตรวจสอบ'"))


# ---------------------------------------------------------------- aging
BUCKETS = [("ยังไม่ถึงกำหนด", None, 0), ("1-30 วัน", 1, 30), ("31-60 วัน", 31, 60), ("61-90 วัน", 61, 90),
           ("91-180 วัน", 91, 180), ("181-365 วัน", 181, 365), ("เกิน 365 วัน", 366, None)]
DEFAULT_ALLOW_RATES = [0, 1, 5, 10, 25, 50, 100]  # % ตัวอย่างนโยบาย — ต้องปรับตามนโยบายบริษัท


def bucket_of(days):
    if days <= 0:
        return 0
    for i, (_, lo, hi) in enumerate(BUCKETS[1:], 1):
        if (hi is None and days >= lo) or (lo <= days <= hi):
            return i
    return len(BUCKETS) - 1


def compute_aging(rows, as_of, party_code, party_name):
    detail, issues = [], []
    for r in rows:
        bal = r2(num(r["amount"]) - num(r["paid"]))
        if abs(bal) < 0.005:
            continue
        due = d(r["due_date"])
        if d(r["doc_date"]) > as_of:
            issues.append(f"{r['doc_no']}: วันที่เอกสารหลังวันที่รายงาน — ไม่นำมารวม")
            continue
        days = (as_of - due).days
        b = bucket_of(days)
        if bal < 0:
            issues.append(f"{r['doc_no']}: ยอดติดลบ {bal:,.2f} (ใบลดหนี้/รับชำระเกิน) ควรตัดชำระกับเอกสารที่เกี่ยวข้อง")
        detail.append(dict(code=r[party_code], name=r[party_name], doc=r["doc_no"], doc_date=d(r["doc_date"]),
                           due=due, days=max(days, 0), bucket=b, bal=bal, remark=r.get("remark", "")))
    detail.sort(key=lambda x: (x["code"], x["due"]))
    summary = OrderedDict()
    for x in detail:
        s = summary.setdefault(x["code"], {"name": x["name"], "b": [0.0] * len(BUCKETS)})
        s["b"][x["bucket"]] += x["bal"]
    return detail, summary, issues


def report_aging(args, kind):
    is_ar = kind == "ar"
    rows = read_csv(args.input)
    as_of = d(args.as_of)
    pc, pn = ("customer_code", "customer_name") if is_ar else ("vendor_code", "vendor_name")
    detail, summary, issues = compute_aging(rows, as_of, pc, pn)
    label = "ลูกหนี้" if is_ar else "เจ้าหนี้"
    wb = Workbook()
    ws = wb.active
    ws.title = f"สรุปอายุ{label}"
    s = Sheet(ws, [10, 40] + [15] * len(BUCKETS) + [16, 9])
    s.title(args.company, f"รายงานวิเคราะห์อายุ{label} (Aging Report)", f"ณ วันที่ {th_date(as_of)}  (นับจากวันครบกำหนดชำระ)")
    s.header(["รหัส", f"ชื่อ{label}"] + [b[0] for b in BUCKETS] + ["รวม", "%"])
    grand = [0.0] * len(BUCKETS)
    total_all = sum(sum(v["b"]) for v in summary.values()) or 1
    for code, v in summary.items():
        t = sum(v["b"])
        for i, x in enumerate(v["b"]):
            grand[i] += x
        overdue90 = sum(v["b"][4:])
        s.line([code, v["name"]] + [r2(x) for x in v["b"]] + [r2(t), round(t / total_all * 100, 1)],
               fill=WARN_FILL if overdue90 > 0 else None)
    s.line(["", "รวมทั้งสิ้น"] + [r2(x) for x in grand] + [r2(sum(grand)), 100.0], bold=True, fill=SUB_FILL)
    s.line(["", "สัดส่วน (%)"] + [round(x / total_all * 100, 1) for x in grand] + [100.0, None], fill=SUB_FILL)
    s.blank()
    if is_ar:
        rates = [float(x) for x in args.rates.split(",")] if args.rates else DEFAULT_ALLOW_RATES
        s.line(["", "อัตราค่าเผื่อฯ ตามนโยบาย (%)"] + rates + [None, None], fill=SUB_FILL)
        allow = [grand[i] * rates[i] / 100 for i in range(len(BUCKETS))]
        s.line(["", "ค่าเผื่อหนี้สงสัยจะสูญที่ควรมี"] + [r2(x) for x in allow] + [r2(sum(allow)), None], bold=True)
        s.blank()
    total = r2(sum(grand))
    if args.gl is not None:
        s.line(["", "ยอดคงเหลือตามบัญชีแยกประเภท (GL)", args.gl], box=False)
        s.line(["", "ยอดรวมตามรายงานอายุ", total], box=False)
        diff = r2(total - args.gl)
        s.line(["", "ผลต่าง (ต้องอธิบายได้)", diff], bold=True, box=False, fill=None if abs(diff) < 0.5 else WARN_FILL)
    if not is_ar:
        s.blank()
        due7 = sum(x["bal"] for x in detail if 0 <= (x["due"] - as_of).days <= 7)
        due30 = sum(x["bal"] for x in detail if 0 <= (x["due"] - as_of).days <= 30)
        overdue = sum(x["bal"] for x in detail if x["due"] < as_of)
        s.line(["", "ประมาณการเงินสดจ่าย: เกินกำหนดแล้ว", r2(overdue)], box=False)
        s.line(["", "ครบกำหนดภายใน 7 วัน", r2(due7)], box=False)
        s.line(["", "ครบกำหนดภายใน 30 วัน", r2(due30)], box=False)
    ws2 = wb.create_sheet("รายละเอียด")
    s2 = Sheet(ws2, [10, 38, 18, 14, 14, 12, 18, 16, 36])
    s2.title(f"รายละเอียดเอกสารค้างชำระ ณ วันที่ {th_date(as_of)}")
    s2.header(["รหัส", "ชื่อ", "เลขที่เอกสาร", "วันที่เอกสาร", "วันครบกำหนด", "เกินกำหนด (วัน)", "ช่วงอายุ", "ยอดคงค้าง", "หมายเหตุ"])
    for x in detail:
        s2.line([x["code"], x["name"], x["doc"], th_date(x["doc_date"]), th_date(x["due"]), x["days"],
                 BUCKETS[x["bucket"]][0], x["bal"], x["remark"]], fill=WARN_FILL if x["bucket"] >= 4 else None)
    if issues:
        ws3 = wb.create_sheet("ข้อสังเกต")
        s3 = Sheet(ws3, [110])
        s3.title("ข้อสังเกตจากการตรวจข้อมูล")
        for t in issues:
            s3.line([t])
    save(wb, args.output)
    print(f"ยอดรวม {label}คงค้าง {total:,.2f}")
    for i, b in enumerate(BUCKETS):
        print(f"  {b[0]:<16}{grand[i]:>16,.2f}")
    for t in issues:
        print("  ! " + t)


# ---------------------------------------------------------------- purchase VAT
FORBIDDEN_KEYWORDS = {
    "รับรอง": "ค่ารับรอง — ภาษีซื้อต้องห้าม (ม.82/5)",
    "รถยนต์นั่ง": "รถยนต์นั่ง/รถโดยสารไม่เกิน 10 คน — ภาษีซื้อต้องห้าม (ม.82/5)",
    "รถเก๋ง": "รถยนต์นั่ง/รถโดยสารไม่เกิน 10 คน — ภาษีซื้อต้องห้าม (ม.82/5)",
    "ใบกำกับภาษีอย่างย่อ": "ใบกำกับภาษีอย่างย่อ — ภาษีซื้อต้องห้าม (ม.82/5)",
}


def months_between(a, b):
    """จำนวนเดือนจาก (ปี,เดือน) a ถึง b"""
    return (b.year - a.year) * 12 + (b.month - a.month)


def check_vat_rows(rows, tax_month, vat_rate, company_tax_id=None):
    out, seen = [], {}
    for i, r in enumerate(rows, 1):
        amt, vat = num(r["amount"]), num(r["vat"])
        flags, block = [], False
        tid = r["vendor_tax_id"].replace("-", "").replace(" ", "")
        if not valid_tax_id(tid):
            flags.append("เลขผู้เสียภาษีผู้ขายไม่ถูกต้อง (check digit/จำนวนหลัก)")
            block = True
        if company_tax_id and tid == company_tax_id:
            flags.append("เลขผู้เสียภาษีผู้ขายตรงกับของบริษัทเอง")
            block = True
        exp = r2(amt * vat_rate / 100)
        if abs(exp - vat) > 1.0:
            flags.append(f"VAT ไม่เท่ากับ {vat_rate:g}% ของมูลค่า (ควรเป็น {exp:,.2f})")
            block = True
        key = (tid, r["tax_invoice_no"].upper())
        if key in seen:
            flags.append(f"เลขที่ใบกำกับซ้ำกับลำดับที่ {seen[key]}")
            block = True
        else:
            seen[key] = i
        dd = d(r["doc_date"])
        gap = months_between(dd, tax_month)
        if gap < 0:
            flags.append("วันที่ใบกำกับอยู่หลังเดือนภาษีที่รายงาน")
            block = True
        elif gap > 6:
            flags.append("เกิน 6 เดือนนับจากเดือนที่ออกใบกำกับ — ใช้เครดิตไม่ได้ ต้องพิจารณายื่นเพิ่มเติม/ขอคืน")
            block = True
        elif gap > 0:
            flags.append(f"ใบกำกับเดือน {TH_MONTHS[dd.month - 1]} — ใช้สิทธิ์ล่าช้า {gap} เดือน (ยังอยู่ในกำหนด 6 เดือน)")
        text = r.get("description", "")
        for kw, msg in FORBIDDEN_KEYWORDS.items():
            if kw in text:
                flags.append(msg)
                block = True
                break
        br = r.get("branch", "").strip()
        branch = "สำนักงานใหญ่" if br in ("", "สำนักงานใหญ่", "สนญ", "00000", "0") else f"สาขาที่ {br.zfill(5)}"
        out.append(dict(no=i, date=dd, inv=r["tax_invoice_no"], vendor=r["vendor_name"], tid=tid, branch=branch,
                        desc=text, amt=amt, vat=vat, flags=flags, ok=not block))
    return out


def report_vat(args):
    rows = read_csv(args.input)
    y, m = map(int, args.month.split("-"))
    tax_month = date(y, m, 1)
    data = check_vat_rows(rows, tax_month, args.rate, args.company_tax_id)
    wb = Workbook()
    ws = wb.active
    ws.title = "รายงานภาษีซื้อ"
    s = Sheet(ws, [7, 13, 18, 38, 17, 15, 17, 15])
    s.title("รายงานภาษีซื้อ", f"เดือนภาษี {TH_MONTHS[m - 1]} ปี {y + 543}")
    s.line(["", f"ชื่อผู้ประกอบการ: {args.company}"], box=False)
    s.line(["", f"เลขประจำตัวผู้เสียภาษีอากร: {args.company_tax_id or '-'}     สถานประกอบการ: {args.branch}"], box=False)
    s.blank()
    s.header(["ลำดับ", "วัน เดือน ปี", "เลขที่ใบกำกับภาษี", "ชื่อผู้ขายสินค้า/ผู้ให้บริการ",
              "เลขประจำตัวผู้เสียภาษีอากรของผู้ขาย", "สถานประกอบการ", "มูลค่าสินค้าหรือบริการ", "จำนวนเงินภาษีมูลค่าเพิ่ม"])
    ok_rows = sorted([x for x in data if x["ok"]], key=lambda x: (x["date"], x["inv"]))
    for n, x in enumerate(ok_rows, 1):
        s.line([n, th_date(x["date"]), x["inv"], x["vendor"], x["tid"], x["branch"], r2(x["amt"]), r2(x["vat"])])
    ta, tv = sum(x["amt"] for x in ok_rows), sum(x["vat"] for x in ok_rows)
    s.line(["", "", "", "รวม", "", "", r2(ta), r2(tv)], bold=True, fill=SUB_FILL)
    ws2 = wb.create_sheet("ตรวจสอบ")
    s2 = Sheet(ws2, [7, 13, 18, 34, 16, 15, 14, 12, 70])
    s2.title("ผลการตรวจสอบใบกำกับภาษีซื้อทุกรายการ")
    s2.header(["ลำดับ", "วันที่", "เลขที่ใบกำกับ", "ผู้ขาย", "เลขผู้เสียภาษี", "มูลค่า", "VAT", "สถานะ", "ข้อสังเกต"])
    for x in data:
        s2.line([x["no"], th_date(x["date"]), x["inv"], x["vendor"], x["tid"], r2(x["amt"]), r2(x["vat"]),
                 "เครดิตได้" if x["ok"] else "ตัดออก", "; ".join(x["flags"])],
                fill=None if x["ok"] else WARN_FILL)
    save(wb, args.output)
    bad = [x for x in data if not x["ok"]]
    print(f"รายการทั้งหมด {len(data)} | เครดิตได้ {len(ok_rows)} | ตัดออก {len(bad)}")
    print(f"ยอดยกไป ภ.พ.30 ช่องยอดซื้อที่มีสิทธิ์ {ta:,.2f}  ภาษีซื้อ {tv:,.2f}")
    for x in data:
        if x["flags"]:
            print(f"  [{x['no']}] {x['inv']}: " + "; ".join(x["flags"]))


# ---------------------------------------------------------------- ratios
def compute_ratios(fs, y, prev=None, average=False):
    def bal(code):
        cur = g(fs, code, y)
        if average and prev is not None and fs.get(code, {}).get(prev) is not None:
            return (cur + g(fs, code, prev)) / 2
        return cur

    G = lambda c: g(fs, c, y)
    rev, cogs = G("PL_REV"), G("PL_COGS")
    gp = rev - cogs
    pbt = rev - cogs + G("PL_OTHINC") - G("PL_SGA") - G("PL_INT")
    ebit = pbt + G("PL_INT")
    np_ = pbt - G("PL_TAX")
    ca = G("BS_CASH") + G("BS_AR") - G("BS_ALLOW") + G("BS_INV") + G("BS_PREPAID")
    cl = G("BS_AP") + G("BS_ACCRUED") + G("BS_TAXPAY") + G("BS_STLOAN")
    ta_end = ca + G("BS_PPE") - G("BS_ACCDEP")
    tl = cl + G("BS_LTLOAN")
    eq = G("BS_CAPITAL") + G("BS_RE")
    ar_net = bal("BS_AR") - bal("BS_ALLOW")
    inv, ap = bal("BS_INV"), bal("BS_AP")
    ta = (bal("BS_CASH") + bal("BS_AR") - bal("BS_ALLOW") + bal("BS_INV") + bal("BS_PREPAID")
          + bal("BS_PPE") - bal("BS_ACCDEP"))
    eq_b = bal("BS_CAPITAL") + bal("BS_RE")
    div = lambda a, b: (a / b) if b else None
    art, invt, apt = div(rev, ar_net), div(cogs, inv), div(cogs, ap)
    dso, dio, dpo = div(365, art), div(365, invt), div(365, apt)
    R = OrderedDict()
    R["สภาพคล่อง"] = [
        ("อัตราส่วนทุนหมุนเวียน (Current Ratio)", div(ca, cl), "เท่า", "สินทรัพย์หมุนเวียน ÷ หนี้สินหมุนเวียน"),
        ("อัตราส่วนทุนหมุนเวียนเร็ว (Quick Ratio)", div(ca - G("BS_INV") - G("BS_PREPAID"), cl), "เท่า",
         "(เงินสด + ลูกหนี้สุทธิ) ÷ หนี้สินหมุนเวียน"),
        ("อัตราส่วนเงินสด (Cash Ratio)", div(G("BS_CASH"), cl), "เท่า", "เงินสด ÷ หนี้สินหมุนเวียน"),
    ]
    R["ประสิทธิภาพการบริหารสินทรัพย์"] = [
        ("อัตราการหมุนเวียนลูกหนี้", art, "รอบ", "รายได้จากการขาย ÷ ลูกหนี้การค้าสุทธิ"),
        ("ระยะเวลาเก็บหนี้เฉลี่ย (DSO)", dso, "วัน", "365 ÷ อัตราการหมุนเวียนลูกหนี้"),
        ("อัตราการหมุนเวียนสินค้าคงเหลือ", invt, "รอบ", "ต้นทุนขาย ÷ สินค้าคงเหลือ"),
        ("ระยะเวลาขายสินค้าเฉลี่ย (DIO)", dio, "วัน", "365 ÷ อัตราการหมุนเวียนสินค้า"),
        ("อัตราการหมุนเวียนเจ้าหนี้", apt, "รอบ", "ต้นทุนขาย ÷ เจ้าหนี้การค้า (ใช้ต้นทุนขายแทนยอดซื้อ)"),
        ("ระยะเวลาชำระหนี้เฉลี่ย (DPO)", dpo, "วัน", "365 ÷ อัตราการหมุนเวียนเจ้าหนี้"),
        ("วงจรเงินสด (Cash Cycle)", (dso + dio - dpo) if None not in (dso, dio, dpo) else None, "วัน", "DSO + DIO − DPO"),
        ("อัตราการหมุนเวียนของสินทรัพย์รวม", div(rev, ta), "รอบ", "รายได้จากการขาย ÷ สินทรัพย์รวม"),
    ]
    R["ความสามารถในการชำระหนี้ (โครงสร้างเงินทุน)"] = [
        ("อัตราส่วนหนี้สินต่อส่วนของผู้ถือหุ้น (D/E)", div(tl, eq), "เท่า", "หนี้สินรวม ÷ ส่วนของผู้ถือหุ้น"),
        ("อัตราส่วนหนี้สินต่อสินทรัพย์รวม", div(tl, ta_end) and div(tl, ta_end) * 100, "%", "หนี้สินรวม ÷ สินทรัพย์รวม"),
        ("อัตราส่วนความสามารถชำระดอกเบี้ย", div(ebit, G("PL_INT")), "เท่า", "กำไรก่อนดอกเบี้ยและภาษี ÷ ดอกเบี้ยจ่าย"),
    ]
    pct = lambda a, b: (a / b * 100) if b else None
    R["ความสามารถในการทำกำไร"] = [
        ("อัตรากำไรขั้นต้น", pct(gp, rev), "%", "กำไรขั้นต้น ÷ รายได้จากการขาย"),
        ("อัตรากำไรจากการดำเนินงาน (EBIT Margin)", pct(ebit, rev), "%", "EBIT ÷ รายได้จากการขาย"),
        ("อัตรากำไรสุทธิ", pct(np_, rev), "%", "กำไรสุทธิ ÷ รายได้จากการขาย"),
        ("อัตราผลตอบแทนจากสินทรัพย์ (ROA)", pct(np_, ta), "%", "กำไรสุทธิ ÷ สินทรัพย์รวม"),
        ("อัตราผลตอบแทนผู้ถือหุ้น (ROE)", pct(np_, eq_b), "%", "กำไรสุทธิ ÷ ส่วนของผู้ถือหุ้น"),
    ]
    balance_diff = ta_end - (tl + eq)
    return R, balance_diff


def report_ratios(args):
    fs, years = load_fs(args.input)
    y0, y1 = years[-2], years[-1]
    r1, diff1 = compute_ratios(fs, y1, y0, args.average)
    r0, diff0 = compute_ratios(fs, y0, None, False)
    wb = Workbook()
    ws = wb.active
    ws.title = "อัตราส่วนทางการเงิน"
    s = Sheet(ws, [48, 15, 15, 14, 9, 52])
    basis = "ยอดเฉลี่ยต้นปี-ปลายปี (เฉพาะปีล่าสุด)" if args.average else "ยอดคงเหลือ ณ วันสิ้นปี"
    s.title(args.company, "รายงานอัตราส่วนทางการเงิน", f"ปี {y1 + 543} เทียบ ปี {y0 + 543}  (ฐานยอดงบแสดงฐานะการเงิน: {basis})")
    s.header(["อัตราส่วน", f"ปี {y1 + 543}", f"ปี {y0 + 543}", "เปลี่ยนแปลง", "หน่วย", "สูตร"])
    for grp, items in r1.items():
        s.line([grp, None, None, None, None, None], bold=True, fill=SUB_FILL)
        for (name, v1, unit, f), (_, v0, _, _) in zip(items, r0[grp]):
            ch = (v1 - v0) if (v1 is not None and v0 is not None) else None
            s.line([name, None if v1 is None else round(v1, 2), None if v0 is None else round(v0, 2),
                    None if ch is None else round(ch, 2), unit, f], indent=1)
    s.blank()
    for y, df in ((y1, diff1), (y0, diff0)):
        s.line([f"ตรวจสอบ: สินทรัพย์ − (หนี้สิน + ส่วนของผู้ถือหุ้น) ปี {y + 543}", r2(df),
                "ผ่าน" if abs(df) < 0.5 else "ไม่สมดุล"], box=False, fill=None if abs(df) < 0.5 else WARN_FILL)
    save(wb, args.output)
    for grp, items in r1.items():
        print(grp)
        for (name, v1, unit, _), (_, v0, _, _) in zip(items, r0[grp]):
            f = lambda v: "-" if v is None else f"{v:,.2f}"
            print(f"  {name:<46}{f(v1):>12}{f(v0):>12} {unit}")


# ---------------------------------------------------------------- CLI
def main():
    p = argparse.ArgumentParser(description="สร้างรายงานบัญชีจากไฟล์ CSV")
    p.add_argument("report", choices=["cashflow", "ar-aging", "ap-aging", "vat", "ratios"])
    p.add_argument("input")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--company", default="บริษัท สยามตัวอย่างการค้า จำกัด")
    p.add_argument("--company-tax-id")
    p.add_argument("--branch", default="สำนักงานใหญ่")
    p.add_argument("--as-of", help="วันที่รายงาน YYYY-MM-DD (aging)")
    p.add_argument("--gl", type=float, help="ยอดคงเหลือตาม GL เพื่อกระทบยอด (aging)")
    p.add_argument("--rates", help="อัตราค่าเผื่อฯ % ต่อช่วงอายุ 7 ช่วง คั่นด้วยจุลภาค (ar-aging)")
    p.add_argument("--month", help="เดือนภาษี YYYY-MM (vat)")
    p.add_argument("--rate", type=float, default=7.0, help="อัตรา VAT %% (ค่าเริ่มต้น 7)")
    p.add_argument("--average", action="store_true", help="ใช้ยอดเฉลี่ยในการคำนวณอัตราหมุนเวียน (ratios)")
    a = p.parse_args()
    if a.report in ("ar-aging", "ap-aging") and not a.as_of:
        p.error("ต้องระบุ --as-of")
    if a.report == "vat" and not a.month:
        p.error("ต้องระบุ --month")
    {"cashflow": report_cashflow, "ratios": report_ratios, "vat": report_vat,
     "ar-aging": lambda x: report_aging(x, "ar"), "ap-aging": lambda x: report_aging(x, "ap")}[a.report](a)


if __name__ == "__main__":
    main()
