from openpyxl import load_workbook
import json
import os
import subprocess
from datetime import datetime
import re
import jdatetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CUSTOMERS_FILE = os.path.join(BASE_DIR, "customers.xlsx")
CUSTOMERS_FOLDER = os.path.join(BASE_DIR, "Customers")
JSON_FILE = os.path.join(BASE_DIR, "customers.json")


def format_money(value):
    if value is None:
        return "0 تومان"

    try:
        value = int(float(value))
        return f"{value:,}".replace(",", ".") + " تومان"
    except:
        return str(value)


def to_number(value):
    """تبدیل مبلغ اکسل به عدد؛ سلول‌های متنی و عنوان‌ها نادیده گرفته می‌شوند."""
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    try:
        text = str(value).strip()
        text = text.replace(",", "").replace("٬", "").replace("٫", ".")
        text = text.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789"))
        return float(text)
    except:
        return None


def parse_jalali_date(value):
    """تاریخ شمسیِ ستون E را می‌خواند؛ تاریخ‌های نامعتبر را نادیده می‌گیرد."""
    if value is None or str(value).strip() == "":
        return None

    # اگر اکسل تاریخ واقعی ذخیره کرده باشد، آن را از میلادی به شمسی تبدیل می‌کنیم.
    if isinstance(value, datetime):
        try:
            return jdatetime.date.fromgregorian(date=value.date())
        except:
            return None

    text = str(value).strip().translate(
        str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")
    )
    # پشتیبانی از جداکننده‌های رایج و تاریخ‌هایی مثل 1405/05/.20
    parts = re.findall(r"\d+", text)
    if len(parts) != 3:
        return None

    try:
        year, month, day = map(int, parts)
        return jdatetime.date(year, month, day)
    except:
        return None


def get_payment_info(customer_ws):
    """
    پرداخت‌ها از ستون C خوانده می‌شوند و تاریخ همان ردیف از ستون E.
    فقط ردیف‌هایی در نظر گرفته می‌شوند که مبلغ پرداخت مثبت و تاریخ معتبر داشته باشند.
    """
    latest_date = None

    for row in range(1, customer_ws.max_row + 1):
        payment_amount = to_number(customer_ws.cell(row=row, column=3).value)
        if payment_amount is None or payment_amount <= 0:
            continue

        payment_date = parse_jalali_date(customer_ws.cell(row=row, column=5).value)
        if payment_date is not None and (latest_date is None or payment_date > latest_date):
            latest_date = payment_date

    if latest_date is None:
        return None, None

    today = jdatetime.date.fromgregorian(date=datetime.now().date())
    days_since = (today - latest_date).days
    return latest_date, max(days_since, 0)


def calculate_score(amount, days_since_payment):
    """امتیاز از 40: حداکثر 20 برای تاریخ پرداخت و 20 برای مانده بدهی."""
    debt = to_number(amount)
    if debt is None:
        debt = 0

    # امتیازِ فاصله از آخرین پرداخت
    if days_since_payment is None:
        payment_points = 0
    elif days_since_payment < 20:
        payment_points = 20
    elif days_since_payment < 30:
        payment_points = 15
    elif days_since_payment < 40:
        payment_points = 5
    else:
        payment_points = 0

    # امتیازِ مانده بدهی
    if debt < 2_000_000:
        debt_points = 20
    elif debt < 3_000_000:
        debt_points = 10
    elif debt < 4_000_000:
        debt_points = 5
    else:
        debt_points = 0

    score = payment_points + debt_points

    if score >= 30:
        grade, status = "A+", "مشتری نمونه"
    elif score >= 20:
        grade, status = "A", "مشتری عالی"
    elif score >= 10:
        grade, status = "B", "مشتری عادی"
    else:
        grade, status = "C", "مشتری بدحساب"

    return score, grade, status


wb = load_workbook(CUSTOMERS_FILE)
ws = wb.active

customers = {}

for row in range(2, ws.max_row + 1):
    code = ws[f"A{row}"].value
    name = ws[f"B{row}"].value
    filename = ws[f"E{row}"].value

    if code is None or filename is None:
        continue

    customer_file = os.path.join(CUSTOMERS_FOLDER, str(filename))

    amount = 0
    date = "نامشخص"
    last_payment_date_text = "نامشخص"
    days_since_payment = None

    if os.path.exists(customer_file):
        # تاریخ آخرین تغییر فایل مشتری؛ این تاریخ فقط تاریخ به‌روزرسانی فایل است.
        file_time = os.path.getmtime(customer_file)
        miladi = datetime.fromtimestamp(file_time)
        shamsi = jdatetime.datetime.fromgregorian(datetime=miladi)
        date = shamsi.strftime("%Y/%m/%d - %H:%M")
        ws[f"D{row}"] = date

        try:
            customer_wb = load_workbook(customer_file, data_only=True)
            customer_ws = customer_wb.active

            # مانده بدهیِ محاسبه‌شده توسط خود اکسل
            amount = customer_ws["H2"].value
            if amount is None:
                amount = 0

            last_payment_date, days_since_payment = get_payment_info(customer_ws)
            if last_payment_date is not None:
                last_payment_date_text = last_payment_date.strftime("%Y/%m/%d")

            customer_wb.close()

        except Exception as e:
            print("Error reading customer file:", customer_file)
            print(e)
    else:
        print("file not found:", customer_file)

    ws[f"C{row}"] = amount

    score, grade, status = calculate_score(amount, days_since_payment)
    pdf_filename = os.path.splitext(str(filename))[0] + ".pdf"

    customers[str(code)] = {
        "name": str(name),
        "money": format_money(amount),
        "date": str(date),
        "file": pdf_filename,
        "score": score,
        "grade": grade,
        "status": status,
        "last_payment_date": last_payment_date_text,
        "days_since_payment": days_since_payment
    }

wb.save(CUSTOMERS_FILE)

with open(JSON_FILE, "w", encoding="utf-8") as f:
    json.dump(customers, f, ensure_ascii=False, indent=4)

print("customers.json created.")

msg = "Update " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")

try:
    subprocess.run(["git", "add", "."], check=True)

    result = subprocess.run(
        ["git", "commit", "-m", msg],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        print("No changes to commit, or commit was not created.")

    subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=True)
    subprocess.run(["git", "push", "origin", "main"], check=True)

    print("")
    print("===================================")
    print("GitHub Updated Successfully.")
    print("===================================")

except Exception as e:
    print("")
    print("Github error")
    print(e)
