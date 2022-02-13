# Importing needed modules ################################################

from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import logging

# Setting up logging configurations ######################################
logging.basicConfig(filename="value_log.log", level=logging.INFO,
                    format="%(message)s")

# File input
file_name = input(
    "Please enter the name of the file you would like to search: ")

# Path for the file and loading up the given excel file
file = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Call Spreadsheets\\' + file_name
wb = load_workbook(file)
ws = wb.worksheets[0]

# Getting the month and date from the file name ##################################

# Extract the month and year from the file name to compare against the sheet data
file_name_lst = (os.path.splitext(file_name)[0].split("_"))

# Convert back to a string without the ".xlsx" or "_"
convert_string = " ".join([str(item) for item in file_name_lst])

# Isolate the part of the string that has the month and year
month_year_string = convert_string[23:]

# Conver the month and year to the same format as the date value on the sheet
formatted_date = datetime.strptime(month_year_string, "%B %Y")


worksheet_rows = list(ws.iter_rows(values_only=True))


# For loop to go through each row and compare the date value to the value pulled from the file name
# If those values match, then it prints the cell values as a list
for row in worksheet_rows:
    if row[0] == formatted_date:
        row_data = [row for row in row if row != None]
        print("Finding your information.")

# Take the raw values and convert them into more readable values for the log
# Convert the date to a readable string and turn the decimals back to percentages
month_converted = row_data[0].strftime("%B %Y")
calls_offered = row_data[1]
abandon_after_30 = row_data[2] * 100
fcr = row_data[3] * 100
dsat = row_data[4] * 100
csat = row_data[5] * 100

# Log the info to the value_log.log file
logging.info("Month: " + month_converted)
logging.info("Calls Offered: " + str(calls_offered))
logging.info("Abandon After 30: " + str(abandon_after_30) + "%")
logging.info("FCR: " + str(fcr) + "%")
logging.info("DSAT: " + str(dsat) + "%")
logging.info("CSAT: " + str(csat) + "%")
