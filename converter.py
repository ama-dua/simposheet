#!/usr/bin/env python3
"""Prototype v0: convert a text-based bank statement PDF into a clean,
formatted Excel workbook.

Usage: convert.py <statement.pdf> <output.xlsx>

Design notes (v0 scope):
- Handles the most common statement layout: each transaction is a line that
  starts with a date, followed by a description, ending with one amount
  (transaction amount) or two amounts (amount + running balance).
- Writes real Excel types (dates as dates, money as numbers), a Debit/Credit
  split, and a Summary sheet with a reconciliation check:
  beginning balance + credits - debits == ending balance.
- Per-bank quirks (odd layouts, scanned PDFs/OCR) are future work.
"""
import json
import re
import sys
from datetime import date

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

AMOUNT = r"\(?-?\$?[\d,]+\.\d{2}\)?"
AMOUNT_RE = re.compile(AMOUNT)
DATE_START_RE = re.compile(
    r"^\s*(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s+(.*)$"
)
MONTH_DATE_START_RE = re.compile(
    r"^\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+"
    r"(\d{1,2})(?:,?\s+(\d{4}))?\s+(.*)$",
    re.IGNORECASE,
)
MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun",
     "jul", "aug", "sep", "oct", "nov", "dec"])}

def xl_date(d):
    """Store dates at midday so no timezone-shifting previewer can move
    them to the previous/next calendar day. Still a true Excel date."""
    from datetime import datetime as _dt
    if isinstance(d, date):
        return _dt(d.year, d.month, d.day, 12, 0)
    return d


SKIP_HINTS = (
    "page ", "continued", "account summary", "summary of",
    "daily balance", "interest charged", "fees charged",
)


def parse_amount(tok: str) -> float:
    neg = tok.strip().startswith("(") and tok.strip().endswith(")")
    val = float(tok.replace("(", "").replace(")", "").replace("$", "")
                .replace(",", "").replace("-", ""))
    if neg or tok.strip().startswith("-"):
        return -val
    return val


def find_label_amount(lines, labels):
    for ln in lines:
        low = ln.lower()
        if any(lbl in low for lbl in labels):
            amounts = AMOUNT_RE.findall(ln)
            if amounts:
                return parse_amount(amounts[-1])
    return None


def find_statement_year(lines):
    for ln in lines:
        m = re.search(r"(20\d{2})\s*[-–]\s*\w+\s+\d{1,2},?\s*(20\d{2})", ln)
        if m:
            return int(m.group(2))
        m = re.search(r"period.*?(\d{1,2})/(\d{1,2})/(\d{4})", ln, re.IGNORECASE)
        if m:
            return int(m.group(3))
    for ln in lines:
        m = re.search(r"\b(20\d{2})\b", ln)
        if m:
            return int(m.group(1))
    return date.today().year


def parse_transactions(lines, year):
    rows = []
    for ln in lines:
        low = ln.lower()
        if any(h in low for h in SKIP_HINTS):
            continue
        m = DATE_START_RE.match(ln)
        mon = None
        if m:
            mon, day = int(m.group(1)), int(m.group(2))
            y = m.group(3)
            rest = m.group(4)
            yr = int(y) if y else year
            if y and len(y) == 2:
                yr = 2000 + yr
        else:
            m2 = MONTH_DATE_START_RE.match(ln)
            if not m2:
                # possible wrapped description line
                if rows and not AMOUNT_RE.search(ln) and ln.strip():
                    rows[-1]["description"] = (
                        rows[-1]["description"] + " " + ln.strip()).strip()
                continue
            mon = MONTHS[m2.group(1).lower()[:3]]
            day = int(m2.group(2))
            yr = int(m2.group(3)) if m2.group(3) else year
            rest = m2.group(4)
        amounts = AMOUNT_RE.findall(rest)
        if not amounts:
            continue
        desc = AMOUNT_RE.sub("", rest).strip(" -")
        desc = re.sub(r"\s+", " ", desc)
        amt = parse_amount(amounts[0])
        balance = parse_amount(amounts[1]) if len(amounts) > 1 else None
        try:
            d = date(yr, mon, day)
        except ValueError:
            continue
        rows.append({"date": d, "description": desc,
                     "amount": amt, "balance": balance})
    # debit/credit split: prefer balance deltas when balances exist
    prev_bal = None
    for r in rows:
        signed = r["amount"]
        if r["balance"] is not None and prev_bal is not None:
            signed = round(r["balance"] - prev_bal, 2)
        if r["balance"] is not None:
            prev_bal = r["balance"]
        r["credit"] = signed if signed > 0 else None
        r["debit"] = -signed if signed < 0 else None
    return rows


def extract_receipt(lines):
    """Receipt mode: a receipt is one expense record, not a transaction list.
    Extract the standard fields bookkeepers need. Card numbers are never
    captured beyond the last-4 digits printed on the receipt itself."""
    text = "\n".join(lines)
    rec = {"date": None, "vendor": None, "amount": None,
           "reference": None, "payment_method": None}

    m = re.search(r"([A-Za-z]+ \d{1,2}, \d{4})", text)
    if m:
        for fmt in ("%B %d, %Y", "%b %d, %Y"):
            try:
                rec["date"] = __import__("datetime").datetime.strptime(
                    m.group(1), fmt).date()
                break
            except ValueError:
                continue

    for ln in lines:
        if ln.strip():
            rec["vendor"] = ln.strip()
            break

    m = re.search(r"(?:Amount Paid|Total|Amount Due)\s*:?\s*\$?([\d,]+\.\d{2})",
                  text, re.IGNORECASE)
    if m:
        rec["amount"] = float(m.group(1).replace(",", ""))
    else:
        amounts = [parse_amount(a) for a in AMOUNT_RE.findall(text)]
        if amounts:
            rec["amount"] = max(amounts)

    # Reference #: first standalone token like AR033336 in reading order
    # (payment/receipt numbers sit near the top of the document).
    m = re.search(r"\b([A-Z]{2}\d{5,})\b", text)
    if m:
        rec["reference"] = m.group(1)

    for ln in lines:
        m = re.search(r"\b(MasterCard|Visa|Amex|American Express|Discover|"
                      r"Cash|Check|ACH)\b", ln, re.IGNORECASE)
        if m:
            rec["payment_method"] = m.group(1).title()
            last4 = re.search(r"\b(\d{4})\b", ln)
            if last4:
                rec["payment_method"] += " ****" + last4.group(1)
            break
    return rec


def write_receipt_workbook(rec, pdf_path, xlsx_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"
    headers = ["Date", "Vendor", "Category", "Description",
               "Amount", "Payment Method"]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")
    desc = f"Receipt {rec['reference']}" if rec["reference"] else "Receipt"
    values = [xl_date(rec["date"]), rec["vendor"], "Uncategorized", desc,
              rec["amount"], rec["payment_method"]]
    for col, v in enumerate(values, 1):
        c = ws.cell(row=2, column=col, value=v)
        if col == 1:
            c.number_format = "mm/dd/yyyy"
        if col == 5:
            c.number_format = '"$"#,##0.00'
    for i, w in enumerate([12, 42, 16, 24, 14, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    sm = wb.create_sheet("Receipt Detail")
    sm.cell(row=1, column=1, value="Extracted Receipt").font = Font(bold=True, size=13)
    detail = [
        ("Source file", pdf_path.split("/")[-1]),
        ("Vendor", rec["vendor"]),
        ("Date", xl_date(rec["date"])),
        ("Amount", rec["amount"]),
        ("Reference #", rec["reference"]),
        ("Payment method", rec["payment_method"]),
    ]
    for i, (k, v) in enumerate(detail, 3):
        sm.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = sm.cell(row=i, column=2, value=v)
        if isinstance(v, float):
            c.number_format = '"$"#,##0.00'
        if isinstance(v, date):
            c.number_format = "mm/dd/yyyy"
    sm.column_dimensions["A"].width = 18
    sm.column_dimensions["B"].width = 44
    wb.save(xlsx_path)


BILL_BRANDS = ["T-Mobile", "Verizon", "AT&T", "Xfinity", "Comcast",
               "Spectrum", "CenturyLink", "Sprint", "Cricket", "Metro"]


def parse_date_str(s):
    for fmt in ("%B %d, %Y", "%b %d, %Y"):
        try:
            return __import__("datetime").datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def extract_bill(lines):
    """Bill mode: a bill is a summary + line items, not one payment.
    Extract the headline fields and the per-line totals table."""
    text = "\n".join(lines)
    bill = {"vendor": None, "bill_date": None, "due_date": None,
            "total_due": None, "past_due": None, "plans": None,
            "equipment": None, "services": None, "current_total": None,
            "items": []}

    for brand in BILL_BRANDS:
        if brand.lower() in text.lower():
            bill["vendor"] = brand
            break
    if not bill["vendor"] and lines:
        bill["vendor"] = lines[0].strip()

    m = re.search(r"([A-Za-z]+ \d{1,2}, \d{4})", text)
    if m:
        bill["bill_date"] = parse_date_str(m.group(1))
    m = re.search(r"due by ([A-Za-z]+ \d{1,2}, \d{4})", text, re.IGNORECASE)
    if m:
        bill["due_date"] = parse_date_str(m.group(1))

    m = re.search(r"TOTAL DUE\s*\$([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        bill["total_due"] = float(m.group(1).replace(",", ""))

    m = re.search(r"PAST DUE BALANCE\s*\$([\d,]+\.\d{2})", text,
                  re.IGNORECASE)
    if m:
        bill["past_due"] = float(m.group(1).replace(",", ""))
    else:
        bill["past_due"] = None

    m = re.search(r"Totals\s+\$([\d,]+\.\d{2})\s+\$([\d,]+\.\d{2})\s+"
                  r"\$([\d,]+\.\d{2})\s+\$([\d,]+\.\d{2})", text)
    if m:
        bill["plans"], bill["equipment"], bill["services"], \
            bill["current_total"] = (float(g.replace(",", ""))
                                      for g in m.groups())

    item_re = re.compile(r"^(Account|\(\d{3}\)\s*\d{3}-\d{4})\s+(.*)$")
    # Per-line totals live in the bill's summary table; the detailed-charges
    # pages repeat the same lines with usage figures, so only scan between
    # the summary and detailed-charges markers when they exist.
    start = end = None
    for i, ln in enumerate(lines):
        if "THIS BILL SUMMARY" in ln.upper() and start is None:
            start = i
        if "DETAILED CHARGES" in ln.upper() and start is not None:
            end = i
            break
    scan = lines[start:end] if start is not None else lines
    for ln in scan:
        m = item_re.match(ln)
        if not m:
            continue
        amounts = AMOUNT_RE.findall(ln)
        if not amounts:
            continue
        rest = m.group(2)
        typ = "Wearable" if "Wearable" in rest else (
            "Voice" if "Voice" in rest else "")
        bill["items"].append({"line": m.group(1), "type": typ,
                              "total": parse_amount(amounts[-1])})
    return bill


def write_bill_workbook(bill, pdf_path, xlsx_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bill Summary"
    ws.cell(row=1, column=1, value="Bill Summary").font = Font(bold=True, size=13)
    items_sum = round(sum(i["total"] for i in bill["items"]), 2) or None
    check = None
    if items_sum is not None and bill["current_total"] is not None:
        check = "MATCH ✓" if abs(items_sum - bill["current_total"]) < 0.01 \
            else f"MISMATCH by {items_sum - bill['current_total']:,.2f}"
    past_due = bill["past_due"]
    if past_due is None and bill["total_due"] is not None and \
            bill["current_total"] is not None:
        past_due = round(bill["total_due"] - bill["current_total"], 2)
    due_check = None
    if past_due is not None and bill["current_total"] is not None and \
            bill["total_due"] is not None:
        s = round(past_due + bill["current_total"], 2)
        due_check = "MATCH ✓" if abs(s - bill["total_due"]) < 0.01 \
            else f"MISMATCH by {s - bill['total_due']:,.2f}"
    fields = [
        ("Source file", pdf_path.split("/")[-1]),
        ("Vendor", bill["vendor"]),
        ("Bill date", bill["bill_date"]),
        ("Due date", bill["due_date"]),
        ("Previous balance (past due)", past_due),
        ("This month - Plans", bill["plans"]),
        ("This month - Equipment", bill["equipment"]),
        ("This month - Services", bill["services"]),
        ("This month - Total charges", bill["current_total"]),
        ("Total due", bill["total_due"]),
        ("Total due check (prev + this month)", due_check),
        ("Line items sum (this month)", items_sum),
        ("Line items check (vs this month)", check),
    ]
    for i, (k, v) in enumerate(fields, 3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = ws.cell(row=i, column=2, value=xl_date(v) if isinstance(v, date) else v)
        if isinstance(v, float):
            c.number_format = '"$"#,##0.00'
        if isinstance(v, date):
            c.number_format = "mm/dd/yyyy"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 30

    li = wb.create_sheet("Line Items")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(["Line", "Type", "Total"], 1):
        c = li.cell(row=1, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")
    for i, item in enumerate(bill["items"], 2):
        li.cell(row=i, column=1, value=item["line"])
        li.cell(row=i, column=2, value=item["type"])
        li.cell(row=i, column=3, value=item["total"]).number_format = '"$"#,##0.00'
    for col, w in [("A", 20), ("B", 12), ("C", 14)]:
        li.column_dimensions[col].width = w
    li.freeze_panes = "A2"
    wb.save(xlsx_path)


def convert(pdf_path, xlsx_path):
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        npages = len(pdf.pages)
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(text.splitlines())
    year = find_statement_year(lines)
    beginning = find_label_amount(
        lines, ["beginning balance", "opening balance", "previous balance"])
    ending = find_label_amount(
        lines, ["ending balance", "closing balance", "new balance"])
    rows = parse_transactions(lines, year)
    # drop rows that are actually the beginning/ending balance lines
    rows = [r for r in rows
            if "balance" not in r["description"].lower()
            or "daily" in r["description"].lower()]

    if not rows:
        joined = "\n".join(lines)
        if re.search(r"TOTAL DUE|Amount Due", joined, re.IGNORECASE) and \
                re.search(r"\bbill\b", joined, re.IGNORECASE):
            bill = extract_bill(lines)
            if bill["total_due"] is not None:
                write_bill_workbook(bill, pdf_path, xlsx_path)
                result = {"mode": "bill", "vendor": bill["vendor"],
                          "total_due": bill["total_due"],
                          "current_total": bill["current_total"],
                          "line_items": len(bill["items"]),
                          "output": xlsx_path}
                print(json.dumps(result, indent=2, default=str))
                return result
        rec = extract_receipt(lines)
        if rec["amount"] is not None:
            write_receipt_workbook(rec, pdf_path, xlsx_path)
            result = {"mode": "receipt", "vendor": rec["vendor"],
                      "amount": rec["amount"], "reference": rec["reference"],
                      "output": xlsx_path}
            print(json.dumps(result, indent=2, default=str))
            return result

    total_credits = round(sum(r["credit"] or 0 for r in rows), 2)
    total_debits = round(sum(r["debit"] or 0 for r in rows), 2)

    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    headers = ["Date", "Description", "Debit", "Credit", "Balance"]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center")
    for i, r in enumerate(rows, 2):
        ws.cell(row=i, column=1, value=xl_date(r["date"])).number_format = "mm/dd/yyyy"
        ws.cell(row=i, column=2, value=r["description"])
        if r["debit"] is not None:
            ws.cell(row=i, column=3, value=r["debit"]).number_format = '"$"#,##0.00'
        if r["credit"] is not None:
            ws.cell(row=i, column=4, value=r["credit"]).number_format = '"$"#,##0.00'
        if r["balance"] is not None:
            ws.cell(row=i, column=5, value=r["balance"]).number_format = '"$"#,##0.00'
    widths = [12, 60, 14, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{max(len(rows) + 1, 2)}"

    sm = wb.create_sheet("Summary")
    computed_ending = None
    status = "UNKNOWN (statement balances not found)"
    if beginning is not None:
        computed_ending = round(beginning + total_credits - total_debits, 2)
        if ending is not None:
            diff = round(computed_ending - ending, 2)
            status = "MATCH ✓" if abs(diff) < 0.01 else f"MISMATCH by {diff:,.2f}"
    summary = [
        ("Source file", pdf_path.split("/")[-1]),
        ("Pages", npages),
        ("Transactions found", len(rows)),
        ("Total credits", total_credits),
        ("Total debits", total_debits),
        ("Beginning balance (from statement)", beginning),
        ("Ending balance (from statement)", ending),
        ("Ending balance (computed)", computed_ending),
        ("Reconciliation", status),
    ]
    sm.cell(row=1, column=1, value="Conversion Summary").font = Font(bold=True, size=13)
    for i, (k, v) in enumerate(summary, 3):
        sm.cell(row=i, column=1, value=k).font = Font(bold=True)
        c = sm.cell(row=i, column=2, value=v)
        if isinstance(v, float):
            c.number_format = '"$"#,##0.00'
    sm.column_dimensions["A"].width = 38
    sm.column_dimensions["B"].width = 30

    wb.save(xlsx_path)
    result = {
        "transactions": len(rows),
        "total_credits": total_credits,
        "total_debits": total_debits,
        "beginning_balance": beginning,
        "ending_balance_statement": ending,
        "ending_balance_computed": computed_ending,
        "reconciliation": status,
        "output": xlsx_path,
    }
    print(json.dumps(result, indent=2))
    return result


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: convert.py <statement.pdf> <output.xlsx>")
        sys.exit(2)
    convert(sys.argv[1], sys.argv[2])
