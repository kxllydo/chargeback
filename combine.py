#from kelly import summary, rg
from openpyxl import Workbook, load_workbook
import pandas as pd
import validator
import csv
import json
import math

from steven.steps import stepone, stepthree
from kelly import summary, rg

if __name__ == "__main__":
    print("")
    
    archive_directory = "../archives/2024/February"
    billing = archive_directory + "/billing.csv"
    chargeback = archive_directory + "/chargeback.xlsx"

    stepone.step_one(billing, chargeback)
    rg.rgComparer(chargeback)
    stepthree.step_three(chargeback, "Feb (2024)")

    wb = load_workbook(chargeback)
    sumSheet = pd.read_excel(chargeback, sheet_name = "Summary")
    summary.creategroupSummarySheet(wb, sumSheet, chargeback)
    summary.createChargeback(wb, sumSheet, chargeback)


if __name__ == "__main__3":
    excel = "c:\\Users\\do-kelly\\Downloads\\chargeback (1).xlsx"
    
    option = validator.validateOption("Hi Jason! What are you trying to accomplish:\n (1) Upload billing csv and extract data\n (2) Compare rgs, and get chargeback\n **NOTE: Only do option 2 if you have done option 1**\n", [1, 2])

    if (option == 1):
        path = input ("Provide the path to your billing csv file:")
    elif (option == 2):
        print("Remember to add last month's rg before providing the path")
        path = input("Provide the path to the generated excel workbook:")
        wb = load_workbook(excel)
        sumSheet = pd.read_excel(excel, sheet_name="Summary")
        summary.creategroupSummarySheet(wb, sumSheet, excel)
        # rg.rgComparer(excel)
        print("Successfully added data analysis and chargeback sheets")

    # while not(validOption):
    #     try:
    #         option = int(input("Hi Jason! What are you trying to accomplish:\n (1) Upload billing csv and extract data\n (2) Compare rgs, and get chargeback\n **NOTE: Only do option 2 if you have done option 1**\n"))
    #         if (option == 1):
    #             path = input ("Provide the path to your billing csv file:")
    #             validOption = True
    #         elif (option == 2):
    #             validOption = True
    #             print("Remember to add last month's rg before providing the path")
    #             path = input("Provide the path to the generated excel workbook:")
    #             wb = load_workbook(excel)
    #             sumSheet = pd.read_excel(excel, sheet_name="Summary")
    #             summary.creategroupSummarySheet(wb, sumSheet, excel)
    #             rg.rgComparer(excel)
    #             print("Successfully added data analysis and chargeback sheets")
    #         else:
    #             print("That was an invalid response. Try again")
    #             validOption = False
    #     except ValueError:
    #         print ("Type in a number. Try again")

        