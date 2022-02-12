from openpyxl import Workbook, load_workbook
from datetime import datetime

file_name = "expedia_report_monthly_january_2018.xlsx"

file_month_year = file_name[23:35].replace("_", " ")

wb = load_workbook(file_name)
ws = wb.active

dateTimeObj = ws['A12'].value

dateStr = dateTimeObj.strftime("%B %Y").lower()

print(dateStr)
print(file_month_year)

if file_month_year == dateStr:
    print("Oh my goooodness")


first_row = ws["A1":"F13"].value

print(first_row)
