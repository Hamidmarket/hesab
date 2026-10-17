from openpyxl import load_workbook
import json
import os
import subprocess
from datetime import datetime, date, timedelta
import re
import jdatetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CUSTOMERS_FILE = os.path.join(BASE_DIR, "customers.xlsx")
CUSTOMERS_FOLDER = os.path.join(BASE_DIR, "Customers")
JSON_FILE = os.path.join(BASE_DIR, "customers.json")

WINDOW_DAYS = 90


def format_money(value):
    try:
        value = float(value or 0)
        return f"{int(value):,}".replace(",", ".") + " تومان"
    except (TypeError, ValueError):
        return str(value)


def to_number(value):
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    text = str(value).strip().replace(",", "").replace("٬", "").replace("٫", ".")
    text = text.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789"))
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def parse_date(value):
    """Parse Excel dates and common Jalali/Gregorian date strings."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789"))
    text = text.replace("-", "/").replace(".", "/")
    match = re.search(r"(\d{4})\s*/\s*(\d{1,2})\s*/\s*(\d{1,2})", text)
    if not match:
        return None
    year, month, day = map(int, match.groups())
    try:
        if year >= 1700:
            return date(year, month, day)
        # Dates like 1404/0130 are malformed and deliberately ignored.
        return jdatetime.date(year, month, day).togregorian()
    except (ValueError, OverflowError):
        return None


def get_balance(customer_ws):
    value = to_number(customer_ws["H2"].value)
    return value if value is not None else 0.0


def grade_for(score):
    if score >= 65:
        return "A+", "خوش‌حساب ممتاز"
    if score >= 55:
        return "A", "خوش‌ حساب"
    if score >= 40:
        return "B", "مشتری معمولی"
    if score >= 25:
        return "C", "مشتری ضعیف"
    return "D", "مشتری بد حساب"

def score_customer(customer_ws, balance, today):
    transactions = []
    has_any_transaction = False

    # In the customer's workbook: B = credit/invoice, C = payment, E = invoice date.
    # Data rows usually begin at row 3; rows without an amount are ignored.
    for row in range(3, customer_ws.max_row + 1):
        credit = to_number(customer_ws[f"B{row}"].value) or 0.0
        payment = to_number(customer_ws[f"C{row}"].value) or 0.0
        if credit > 0 or payment > 0:
            has_any_transaction = True
            txn_date = parse_date(customer_ws[f"E{row}"].value)
            if txn_date:
                transactions.append({
                    "date": txn_date,
                    "credit": max(0.0, credit),
                    "payment": max(0.0, payment),
                })

    if not has_any_transaction:
        return {
            "score": None,
            "grade": "—",
            "status": "مشتری جدید / سابقه ناکافی" if balance >= 0 else "بستانکار",
            "score_details": None,
            "last_payment_date": None,
            "days_since_payment": None,
            "credit_90_days": 0,
            "payment_90_days": 0,
        }

    # If there are transactions but none has a readable date, do not label the customer bad
    # based on missing/invalid dates. Show insufficient history instead.
    if not transactions:
        return {
            "score": None,
            "grade": "—",
            "status": "سابقه تاریخ‌دار کافی نیست",
            "score_details": None,
            "last_payment_date": None,
            "days_since_payment": None,
            "credit_90_days": 0,
            "payment_90_days": 0,
        }

    cutoff = today - timedelta(days=WINDOW_DAYS)
    recent = [t for t in transactions if cutoff <= t["date"] <= today]
    recent_credit = sum(t["credit"] for t in recent)
    recent_payment = sum(t["payment"] for t in recent)
    recent_payment_events = sum(1 for t in recent if t["payment"] > 0)

    payment_dates = [t["date"] for t in transactions if t["payment"] > 0 and t["date"] <= today]
    last_payment = max(payment_dates) if payment_dates else None
    days_since_payment = (today - last_payment).days if last_payment else None

    # 40 points: payments relative to new credit during the last 90 days.
    # If there were payments but no new credit in this window, do not penalize paying old debt.
    if recent_credit > 0:
        ratio = recent_payment / recent_credit
        payment_score = 40 if ratio >= 1 else 32 if ratio >= .75 else 24 if ratio >= .50 else 12 if ratio >= .25 else 5 if ratio > 0 else 0
    elif recent_payment > 0:
        payment_score = 40
    else:
        payment_score = 0

    # 25 points: current balance compared with credit taken during the period.
    if balance <= 0:
        balance_score = 25
    elif recent_credit > 0:
        debt_ratio = balance / recent_credit
        balance_score = 25 if debt_ratio <= .25 else 20 if debt_ratio <= .50 else 15 if debt_ratio <= .75 else 10 if debt_ratio <= 1 else 5
    else:
        balance_score = 0

    # 20 points: regularity, measured by separate payment invoices in the last 90 days.
    consistency_score = 20 if recent_payment_events >= 4 else 16 if recent_payment_events == 3 else 12 if recent_payment_events == 2 else 7 if recent_payment_events == 1 else 0

    # 15 points: recency of last recorded payment.
    if days_since_payment is None:
        recency_score = 0
    elif days_since_payment < 20:
        recency_score = 15
    elif days_since_payment < 31:
        recency_score = 12
    elif days_since_payment < 46:
        recency_score = 8
    elif days_since_payment < 61:
        recency_score = 4
    else:
        recency_score = 0

    score = int(payment_score + balance_score + consistency_score + recency_score)
    grade, grade_label = grade_for(score)

    if balance < 0:
        status = "بستانکار"
    elif balance == 0:
        status = "تسویه‌شده"
    else:
        status = grade_label

    return {
        "score": score,
        "grade": grade,
        "status": status,
        "score_details": {
            "payment_ratio": payment_score,
            "balance": balance_score,
            "consistency": consistency_score,
            "recency": recency_score,
        },
        "last_payment_date": jdatetime.date.fromgregorian(date=last_payment).strftime("%Y/%m/%d") if last_payment else None,
        "days_since_payment": days_since_payment,
        "credit_90_days": int(recent_credit),
        "payment_90_days": int(recent_payment),
    }


wb = load_workbook(CUSTOMERS_FILE)
ws = wb.active
customers = {}
today = datetime.now().date()

for row in range(2, ws.max_row + 1):
    code = ws[f"A{row}"].value
    name = ws[f"B{row}"].value
    filename = ws[f"E{row}"].value
    if code is None or filename is None:
        continue

    customer_file = os.path.join(CUSTOMERS_FOLDER, str(filename))
    amount = 0.0
    updated_date = "نامشخص"
    scored = {
        "score": None, "grade": "—", "status": "سابقه ناکافی",
        "score_details": None, "last_payment_date": None,
        "days_since_payment": None, "credit_90_days": 0, "payment_90_days": 0,
    }

    if os.path.exists(customer_file):
        file_time = os.path.getmtime(customer_file)
        miladi = datetime.fromtimestamp(file_time)
        updated_date = jdatetime.datetime.fromgregorian(datetime=miladi).strftime("%Y/%m/%d - %H:%M")
        try:
            customer_wb = load_workbook(customer_file, data_only=True, read_only=True)
            customer_ws = customer_wb.active
            amount = get_balance(customer_ws)
            scored = score_customer(customer_ws, amount, today)
            customer_wb.close()
        except Exception as exc:
            print(f"خطا در خواندن فایل مشتری {filename}: {exc}")
    else:
        print(f"فایل مشتری پیدا نشد: {customer_file}")

    ws[f"C{row}"] = amount
    ws[f"D{row}"] = updated_date
    pdf_filename = os.path.splitext(str(filename))[0] + ".pdf"

    customer_record = {
        "name": str(name),
        "money": format_money(amount),
        "date": str(updated_date),
        "file": pdf_filename,
        **scored,
    }
    customers[str(code).strip()] = customer_record

wb.save(CUSTOMERS_FILE)
with open(JSON_FILE, "w", encoding="utf-8") as f:
    json.dump(customers, f, ensure_ascii=False, indent=4)

print("customers.json created successfully.")
print(f"Number of customers processed: {len(customers)}")

msg = "Update credit score " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")
try:
    subprocess.run(["git", "add", "."], check=True)
    result = subprocess.run(["git", "commit", "-m", msg], capture_output=True, text=True)
    if result.returncode != 0:
        print("چیزی برای commit جدید وجود نداشت یا commit انجام نشد.")
    subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
    subprocess.run(["git", "push", "origin", "main"], check=True)
    print("GitHub updated successfully.")
except Exception as exc:
    print("khata dar beroz rasani github")
    print(exc)
