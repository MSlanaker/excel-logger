# Importing needed modules ################################################
import os
from datetime import datetime
from openpyxl import load_workbook
import logging

# Setting up logging configurations ######################################
logging.basicConfig(filename="value_log.log", level=logging.INFO,
                    format="%(message)s")

file_path = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Call Spreadsheets\\'
arc_path = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Archive\\'
err_path = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Error\\'

# Initial check to make sure non-excel files and improperly named files are removed ################################


def get_error_files(path):
    os.chdir(path)
    lst = os.listdir()

    err_files = [s for s in lst if '.xlsx' not in s or s[:23]
                 != 'expedia_report_monthly_']

    return err_files


bad_files = get_error_files(file_path)

print("These file(s) could not be processed and are being moved to the 'Error' folder:", bad_files)

# Moving any errored files into the error folder #########################################################
source = file_path
destination = err_path

for f in bad_files:
    os.rename(source + f, destination + f)


# Grabbing the remaining files and putting them into a list to be processed #########################################
def get_files(path):
    os.chdir(path)
    lst = os.listdir()

    excel_files = [s for s in lst if '.xlsx' in s]

    return excel_files


files = get_files(file_path)

print(files)


def log_first_sheet_info(file_list):
    logging.basicConfig(filename="value_log.log", level=logging.INFO,
                        format="%(message)s")
    for x in file_list:
        try:
            # Path for the file and loading up the given excel file
            file = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Call Spreadsheets\\' + x
            wb = load_workbook(file)
            ws = wb.worksheets[0]
        except:
            print("There was an issue with this file:", x)

        try:
            # Extract the month and year from the file name to compare against the sheet data
            file_name_lst = (os.path.splitext(x)[0].split("_"))
            # Convert back to a string without the ".xlsx" or "_"
            convert_string = " ".join([str(item) for item in file_name_lst])
            # Isolate the part of the string that has the month and year
            month_year_string = convert_string[23:].capitalize()
            # Conver the month and year to the same format as the date value on the sheet
            formatted_date = datetime.strptime(month_year_string, "%B %Y")
        except:
            print("There was an issue with this file:", x)
            os.rename(source + x, destination + x)

        # Iterates through the rows using the built in method
        worksheet_rows = list(ws.iter_rows(values_only=True))

        # For loop to go through each row and the data type to the formatted month and year
        # Checking if the item in position [0] is a datetime object, if so grab that row and append it to a new list
        for row in worksheet_rows:
            if type(row[0]) == type(formatted_date):
                row_data = [row for row in row if row != None]

                # If the fist item in the list is a datetime object it then converts it to a string
                # Then creates a new, nested list with the date string as the first item
                if type(row_data[0]) == type(formatted_date):
                    date = row_data[0].strftime("%B %Y")
                    date_list = (date, row_data)

                    # If the the date string matches the date string from the file name
                    if date_list[0] == month_year_string:

                        # Take the raw values and convert them into more readable values for the log
                        # Convert the date to a readable string and turn the decimals back to percentages
                        month_converted = date_list[0]
                        calls_offered = date_list[1][1]
                        abandon_after_30 = date_list[1][2] * 100
                        fcr = date_list[1][3] * 100
                        dsat = date_list[1][4] * 100
                        csat = date_list[1][5] * 100

                        # Log the info to the value_log.log file
                        logging.info("----Summary Rolling MoM----")
                        logging.info("Month: " + month_converted)
                        logging.info("Calls Offered: " + str(calls_offered))
                        logging.info("Abandon After 30: " +
                                     str(abandon_after_30) + "%")
                        logging.info("FCR: " + str(fcr) + "%")
                        logging.info("DSAT: " + str(dsat) + "%")
                        logging.info("CSAT: " + str(csat) + "%")

                        return files


log_first_sheet_info(files)


def log_second_sheet_info(file_list):
    logging.basicConfig(filename="value_2.log", level=logging.INFO,
                        format="%(message)s")
    for x in file_list:

        try:
            file = 'C:\\Users\\Owner\\OneDrive\\Desktop\\excel-logger\\Call Spreadsheets\\' + x
            wb = load_workbook(file)
            ws_voc = wb.worksheets[1]
        except:
            print("There was an issue with this file:", x)

        try:
            file_name_lst = (os.path.splitext(x))[0].split("_")
            convert_string = " ".join([str(item) for item in file_name_lst])
            month_year_string = convert_string[23:].capitalize()
            formatted_date = datetime.strptime(month_year_string, "%B %Y")
        except:
            print("There was an issue with this file:", x)

        worksheet_rows_2nd_sheet = list(ws_voc.iter_rows(values_only=True))

        worksheet_columns = list(ws_voc.iter_cols(values_only=True))

        for col in worksheet_columns:
            if type(col[0]) == type(formatted_date):
                col_data = [col for col in col if col != None]

                if type(col_data[0]) == type(formatted_date):
                    date = col_data[0].strftime("%B %Y")
                    date_list = (date, col_data)

                    if date_list[0] == month_year_string:

                        month_converted = date_list[0]
                        promoters = date_list[1][2]
                        passives = date_list[1][4]
                        detractors = date_list[1][6]

                        logging.info("----VOC Rolling MoM----")

                        logging.info("Month: " + month_converted)

                        if promoters > 200:
                            logging.info("Promotors are " + str(
                                promoters) + " - good!")
                        elif promoters < 200:
                            logging.info("Promotors are " +
                                         str(promoters) + " - bad.")
                        if passives > 100:
                            logging.info("Passives are " + str(
                                passives) + " - good!")
                        elif passives < 100:
                            logging.info("Passives are " +
                                         str(passives) + " - bad.")
                        if detractors > 100:
                            logging.info("Passives are " + str(
                                detractors) + " - good!")
                        elif detractors < 100:
                            logging.info("Passives are " + str(
                                detractors) + " - bad.")


log_second_sheet_info(files)


source = file_path
destination = arc_path

for f in files:
    os.rename(source + f, destination + f)
