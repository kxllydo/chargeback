import os
import datetime

from steps.stepone import step_one
from steps.stepthree import step_three

def check_for_perms(path):
    if not os.path.exists(path):
        print(f"{os.path.abspath(path)} does not exist.")
        exit(-1)
    if not os.access(path, os.R_OK):
        print(f"Please grant READ permissions to {os.path.abspath(path)}")
        exit(-1)
    if not os.access(path, os.W_OK):
        print(f"Please grant WRITE permissions to {os.path.abspath(path)}")
        exit(-1)

if __name__ == "__main__":
    #current_date   = datetime.datetime.now()
    #month          = current_date.strftime("%B")
    #year           = current_date.strftime("%Y")

    # DELETE LATER
    month = "February"
    year = "2024"
    # DELETE LATER

    archive_directory = "../../archives"  # Path to top level parent folder of all files
    check_for_perms(archive_directory)
    check_for_perms(archive_directory + os.sep + year)
    check_for_perms(archive_directory + os.sep + year + os.sep + month)

    archive_directory   = archive_directory + os.sep + year + os.sep + month + os.sep
    billing_file        = archive_directory + "billing.csv"
    excel_file          = archive_directory + "chargeback.xlsx"
    check_for_perms(billing_file)
    check_for_perms(excel_file)

    step_one(billing_file, excel_file) # Completes current month's RG Comparison sheet, filling Current Month RG and Cost
                                       # Completes next months' RG Comparison sheet, fliling Last Month RG
    step_three(excel_file, 
               f"{month} ({year})")    # Adds current month's RGs to Summary Sheet