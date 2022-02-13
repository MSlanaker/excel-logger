from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

file_name = "expedia_report_monthly_january_2018.xlsx"

wb = load_workbook(file_name)
ws = wb.worksheets[0]

# Extract the month and year from the file name to compare against the sheet data
file_name_lst = (os.path.splitext(file_name)[0].split("_"))
print(file_name_lst)

# Convert back to a string without the ".xlsx" or "_"
convert_string = " ".join([str(item) for item in file_name_lst])


# Isolate the part of the string that has the month and year
month_year_string = convert_string[23:]

print(month_year_string)

# Conver the month and year to the same format as the date value on the sheet
formatted_date = datetime.strptime(month_year_string, "%B %Y")

print(formatted_date)


worksheet_rows_1st_sheet = list(ws.iter_rows(values_only=True))

# For loop to go through each row and compare the date value to the value pulled from the file name
# If those values match, then it prints the cell values as a list
for row in worksheet_rows_1st_sheet:
    if row[0] == formatted_date:
        row_data = [row for row in row if row != None]
        print(row_data)

# Take the raw values and convert them into more readable values for the log
# Convert the date to a readable string and turn the decimals back to percentages
month_converted = row_data[0].strftime("%B %Y")
calls_offered = row_data[1]
abandon_after_30 = row_data[2] * 100
fcr = row_data[3] * 100
dsat = row_data[4] * 100
csat = row_data[5] * 100

print(month_converted, calls_offered, abandon_after_30, fcr, dsat, csat)
