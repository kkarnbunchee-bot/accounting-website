# ชุดเครื่องมือลดเวลางานฝ่ายบัญชี (5 รายงาน)

| # | รายงาน | Prompt | Skill | App (แท็บ) | ข้อมูลตัวอย่าง |
|---|---|---|---|---|---|
| 1 | งบกระแสเงินสด (วิธีทางอ้อม) | `prompts/01-cash-flow-statement.md` | `skills/cash-flow-statement-th` | 1 | `financial_statements.csv` |
| 2 | รายงานอายุลูกหนี้ | `prompts/02-ar-aging.md` | `skills/ar-aging-report-th` | 2 | `ar_open_invoices.csv` |
| 3 | รายงานอายุเจ้าหนี้ | `prompts/03-ap-aging.md` | `skills/ap-aging-report-th` | 3 | `ap_open_invoices.csv` |
| 4 | รายงานภาษีซื้อ | `prompts/04-purchase-vat-report.md` | `skills/purchase-vat-report-th` | 4 | `purchase_vat_invoices.csv` |
| 5 | อัตราส่วนทางการเงิน | `prompts/05-financial-ratios.md` | `skills/financial-ratios-th` | 5 | `financial_statements.csv` |

## 3 ระดับการใช้งาน
1. **Prompt** — คัดลอกไปวางใน Claude พร้อมแนบไฟล์ เหมาะกับงานครั้งคราว
2. **Skill** — อัปโหลดโฟลเดอร์ skill (zip) ที่ claude.ai → Settings → Capabilities → Skills หรือวางใน `.claude/skills/` ของ Claude Code
   Claude จะเรียกใช้อัตโนมัติเมื่อพูดถึงรายงานนั้น และรันสคริปต์ `scripts/thai_acct_reports.py` เพื่อคำนวณแบบตรวจสอบได้ (ไม่ใช้ AI คิดเลข)
3. **App** — เปิด `app/index.html` ในเบราว์เซอร์ (ไม่ต้องติดตั้ง ไม่ส่งข้อมูลออกนอกเครื่อง) วาง CSV → ได้รายงาน + ตรวจสอบ + ส่งออก CSV/พิมพ์ PDF

## สคริปต์ (ใช้ได้โดยตรง)
```bash
pip install openpyxl
python scripts/thai_acct_reports.py cashflow sample-data/financial_statements.csv -o cashflow.xlsx
python scripts/thai_acct_reports.py ar-aging sample-data/ar_open_invoices.csv --as-of 2026-10-31 --gl 2053500 -o ar.xlsx
python scripts/thai_acct_reports.py ap-aging sample-data/ap_open_invoices.csv --as-of 2026-10-31 --gl 523190 -o ap.xlsx
python scripts/thai_acct_reports.py vat sample-data/purchase_vat_invoices.csv --month 2026-10 --company-tax-id 0999900001135 -o vat.xlsx
python scripts/thai_acct_reports.py ratios sample-data/financial_statements.csv -o ratios.xlsx
```
ผลลัพธ์ตัวอย่างอยู่ใน `sample-output/`

## ข้อมูลตัวอย่าง (สมมติทั้งหมด)
บริษัท สยามตัวอย่างการค้า จำกัด — ชื่อบริษัท/เลขผู้เสียภาษีทั้งหมดเป็นข้อมูลสมมติ (ขึ้นต้น 09999)

| รายงาน | ผลลัพธ์ที่คาดหวัง | จุดทดสอบที่ใส่ไว้ |
|---|---|---|
| งบกระแสเงินสด ปี 2568 | CFO 1,342,000 · CFI (830,000) · CFF (400,000) · เพิ่มขึ้นสุทธิ 112,000 | กระทบยอดผ่านครบ 4 ข้อ |
| อายุลูกหนี้ ณ 31/10/2569 | ยอดค้าง 2,042,800 · เกิน 90 วัน 216,800 · ค่าเผื่อฯ 149,916 | ใบลดหนี้ติดลบ, ผลต่าง GL (10,700), ลูกหนี้เกิน 365 วัน |
| อายุเจ้าหนี้ ณ 31/10/2569 | ยอดค้าง 523,190 · ครบกำหนดใน 30 วัน 171,200 | เจ้าหนี้โต้แย้งยอด 190 วัน |
| ภาษีซื้อ ต.ค. 2569 | 13 ใบ → เครดิตได้ 7 ใบ · ยอดซื้อ 321,191.59 · ภาษีซื้อ 22,483.41 | เลขผู้เสียภาษีผิด, VAT ไม่ตรง 7%, ใบซ้ำ, เกิน 6 เดือน, ค่ารับรอง, รถยนต์นั่ง |
| อัตราส่วน 2568 vs 2567 | Current 2.53 · D/E 0.61 · NPM 7.95% · ROE 17.89% · Cash cycle 80.30 วัน | งบดุลสมดุลทั้งสองปี |

## รูปแบบไฟล์นำเข้า
- CSV UTF-8, วันที่ `YYYY-MM-DD` (ค.ศ.), ตัวเลขไม่มีจุลภาค — รายงานแสดงผลเป็น พ.ศ.
- คอลัมน์ดูจากไฟล์ใน `sample-data/` (export จากโปรแกรมบัญชีแล้วเปลี่ยนชื่อหัวคอลัมน์ให้ตรง)

## ข้อจำกัด / สิ่งที่ต้องยืนยัน
- แบบฟอร์มรายงานจริงของฝ่ายบัญชี (ที่ระบุว่าแนบมา) ยังไม่ได้รับ — ชุดนี้ออกแบบตามรูปแบบมาตรฐาน ปรับหัวคอลัมน์/ลำดับได้เมื่อได้ฟอร์มจริง
- อัตราค่าเผื่อหนี้สงสัยจะสูญเป็นตัวอย่าง ต้องใช้ตามนโยบายบริษัท
- การตรวจภาษีซื้อต้องห้ามใช้คำสำคัญ ต้องมีผู้ตรวจทานทุกเดือน · อัตรา VAT ให้ยืนยันตามพระราชกฤษฎีกาที่ใช้บังคับ
- งบกระแสเงินสดรองรับโครงสร้างบัญชีพื้นฐาน รายการพิเศษ (เช่าซื้อ, ตัดหนี้สูญ, ปันผลค้างจ่าย, เงินลงทุน) ต้องปรับเพิ่ม
