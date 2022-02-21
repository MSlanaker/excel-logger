# Importing needed modules ################################################
import os
from datetime import datetime
from openpyxl import load_workbook
import logging
import log_functions

# Setting up logging configurations ######################################
logging.basicConfig(filename="value_log.log", level=logging.INFO,
                    format="%(message)s")

# Declaring some pathing variables for file location manipulation #####################
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

# Moving any errored files into the error folder #########################################################
print("These file(s) could not be processed and are being moved to the 'Error' folder:", bad_files)


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

# Calling the logging functions from our log_functions module ###########################################

log_functions.log_first_sheet_info(files)

log_functions.log_second_sheet_info(files)

# Taking the processed files and moving them to the archive folder ##################################################
source = file_path
destination = arc_path

for f in files:
    os.rename(source + f, destination + f)
