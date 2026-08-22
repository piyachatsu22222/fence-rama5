# -*- coding: utf-8 -*-
"""
วิเคราะห์ PO งานรั้ว โครงการ 4111 (บ้านกลางเมือง พระราม 5 / บางศรีเมือง)

อ่านไฟล์ SAP ME2N export "4111_po.xlsx" แบบออฟไลน์ 100% (openpyxl)
แล้วสรุปยอดเงินแยกตาม WBS Element และเลขที่ PO

โครงสร้างไฟล์ต้นทาง (sheet "4111_po"):
    แถวที่ 4 = หัวตาราง, แถวที่ 5 เป็นต้นไป = ข้อมูล
    คอลัมน์ที่ใช้ (index เริ่มจาก 0):
        1  WBS Element      4  Purch.Doc.      5  Item
        2  Name             7  Doc. Date      12  Vendor
       14  Short Text      15  Quantity       16  Net price
    มูลค่า = Quantity x Net price

การใช้งาน:
    py -3 วิเคราะห์_PO_รั้ว_4111.py [path ของ 4111_po.xlsx] [path ไฟล์ผลลัพธ์]
"""

import os
import sys
import collections

import openpyxl

DEFAULT_SOURCE = r"C:\Users\piyac\Documents\0.04 งานรั้วแบ่งแปลง\OneDrive\4111_po.xlsx"
DEFAULT_OUTPUT = r"C:\Users\piyac\Documents\รั้วแบ่งแปลง\สรุป_PO_รั้ว_4111.txt"
SHEET_NAME = "4111_po"
HEADER_ROW = 4

COL_WBS, COL_NAME, COL_DOC, COL_ITEM = 1, 2, 4, 5
COL_DATE, COL_VENDOR, COL_TEXT = 7, 12, 14
COL_QTY, COL_PRICE = 15, 16

# WBS ของงานรั้วแต่ละประเภท เรียงตามลำดับที่ต้องการแสดงผล
FENCE_WBS = [
    ("FH-001", "รั้วแบ่งแปลง / รั้วบ้าน (รายแปลง)"),
    ("FN-001", "รั้วโครงการ เหนือดิน H 3.0 ม."),
    ("FN-002", "รั้วชั่วคราว / รั้วกั้นโซนก่อสร้าง"),
    ("FN-003", "รั้วเขื่อน Type P2, P3, P4 และ Type พิเศษ"),
    ("FN-004", "งานวางหมุด ทาสีรั้ว รั้วโปร่งซุ้ม"),
    ("FC-001", "ซุ้ม ป้อม ป้าย และประตูรั้วทางเข้าหลัก"),
]


def to_number(value):
    """แปลงค่าในเซลล์เป็นตัวเลข ถ้าแปลงไม่ได้ให้เป็น 0"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return 0.0


def load_rows(source):
    workbook = openpyxl.load_workbook(source, data_only=True, read_only=True)
    try:
        sheet = workbook[SHEET_NAME]
        rows = [
            row
            for row in sheet.iter_rows(min_row=HEADER_ROW + 1, values_only=True)
            if row[COL_DOC] is not None
        ]
    finally:
        workbook.close()
    return rows


def sort_key_by_date(row):
    """เรียงตามวันที่เอกสารรูปแบบ dd.mm.yyyy"""
    date_text = str(row[COL_DATE])
    return date_text[-4:] + date_text[3:5] + date_text[:2]


def build_report(rows):
    lines = []
    lines.append("สรุปมูลค่า PO งานรั้ว โครงการ 4111 (บ้านกลางเมือง พระราม 5)")
    lines.append("แหล่งข้อมูล: SAP ME2N export วันที่ตัดยอด 04.11.2023")
    lines.append("มูลค่า = Quantity x Net price (ยังไม่รวม VAT)")
    lines.append("=" * 96)
    lines.append("")

    grand_total = 0.0
    summary = []

    for tag, description in FENCE_WBS:
        selected = [row for row in rows if tag in str(row[COL_WBS])]
        if not selected:
            continue

        subtotal = sum(to_number(r[COL_QTY]) * to_number(r[COL_PRICE]) for r in selected)
        grand_total += subtotal
        summary.append((tag, description, len(selected), subtotal))

        lines.append(f"### 4/111-UT-{tag}  {description}")
        lines.append(f"    จำนวนรายการ {len(selected)} รายการ   รวม {subtotal:,.2f} บาท")
        lines.append(
            f"    {'เลขที่ PO':<13}{'Item':>5}{'ปริมาณ':>10}{'ราคา/หน่วย':>15}"
            f"{'มูลค่า':>16}  {'วันที่':<12}รายละเอียด"
        )
        for row in sorted(selected, key=sort_key_by_date):
            qty = to_number(row[COL_QTY])
            price = to_number(row[COL_PRICE])
            detail = str(row[COL_TEXT] or row[COL_NAME] or "")[:46]
            lines.append(
                f"    {str(row[COL_DOC]):<13}{str(row[COL_ITEM]):>5}{qty:>10,.2f}"
                f"{price:>15,.2f}{qty * price:>16,.2f}  {str(row[COL_DATE])[:10]:<12}{detail}"
            )
        lines.append("")

    lines.append("=" * 96)
    lines.append("สรุปรวมทุกหมวดงานรั้ว")
    lines.append(f"{'WBS':<20}{'รายการ':>8}{'มูลค่า (บาท)':>18}   คำอธิบาย")
    for tag, description, count, subtotal in summary:
        lines.append(f"4/111-UT-{tag:<11}{count:>8}{subtotal:>18,.2f}   {description}")
    lines.append(f"{'รวมทั้งสิ้น':<20}{'':>8}{grand_total:>18,.2f}")
    lines.append("")

    # แยกเฉพาะรั้วแบ่งแปลงรายแปลง (ตัดรายการที่ไม่ใช่ตัวรั้วออก เช่น ตู้รับจดหมาย)
    plot_rows = [
        row
        for row in rows
        if "FH-001" in str(row[COL_WBS])
        and "รั้วกั้นระหว่างแปลง" in str(row[COL_TEXT] or "")
    ]
    plot_total = sum(to_number(r[COL_QTY]) * to_number(r[COL_PRICE]) for r in plot_rows)
    by_po = collections.Counter()
    for row in plot_rows:
        by_po[str(row[COL_DOC])] += to_number(row[COL_QTY]) * to_number(row[COL_PRICE])

    lines.append("=" * 96)
    lines.append("เฉพาะ 'ค่าก่อสร้างรั้วกั้นระหว่างแปลง' (ตัดตู้รับจดหมายและรายการอื่นออก)")
    lines.append(f"จำนวน {len(plot_rows)} แปลง  รวม {plot_total:,.2f} บาท")
    for doc, amount in sorted(by_po.items()):
        lines.append(f"    PO {doc}   {amount:>16,.2f}")

    return "\n".join(lines)


def main():
    source = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SOURCE
    output = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_OUTPUT

    if not os.path.exists(source):
        print(f"ไม่พบไฟล์ต้นทาง: {source}")
        return 1

    rows = load_rows(source)
    report = build_report(rows)

    with open(output, "w", encoding="utf-8") as handle:
        handle.write(report)

    print(f"อ่านข้อมูล {len(rows)} แถว")
    print(f"บันทึกผลสรุปที่: {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
