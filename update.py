from openpyxl import load_workbook
import json
import subprocess
from datetime import datetime

wb = load_workbook("customers.xlsx")
ws = wb.active

customers = {}

for row in ws.iter_rows(min_row=2, values_only=True):
    code, name, money, date, extra = row

    code = str(code).zfill(4)

    customers[code] = {
        "name": str(name),
        "money": str(money),
        "date": str(date)
    }

with open("customers.json", "w", encoding="utf-8") as f:
    json.dump(customers, f, ensure_ascii=False, indent=4)

print("customers.json created.")

msg = "Update " + datetime.now().strftime("%Y-%m-%d %H:%M:%S")

subprocess.run(["git", "add", "."], check=True)
subprocess.run(["git", "commit", "-m", msg], check=True)
subprocess.run(["git", "push"], check=True)

print("GitHub Updated Successfully.")
