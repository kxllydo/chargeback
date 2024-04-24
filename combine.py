from kelly import summary
from kelly import rg
from steven import resourcegroups
from steven import summary as summary2
from openpyxl import load_workbook
import pandas as pd

import openpyxl as excel

if __name__ == "__main__":
    # Set CSV path
    billing = "../archives/2024/February/billing.csv"
    # Set XLSX path
    chargeback = "../archives/2024/February/chargeback.xlsx"
    chargeback2 = chargeback.replace(".xlsx", "2.xlsx")
    excel.Workbook().save(chargeback2)
    # Set Billing Cycle: (Example: February (2024))
    billingCycle = "February (2024)"

    resourcegroups.step_one(billing, chargeback)
    rg.rgComparer(chargeback)

    resourcegroups.step_two(chargeback, chargeback2)
    summary2.step_three(chargeback, billingCycle)
    summary2.step_four(chargeback, chargeback2)

    wb = load_workbook("chargeback")
    summarySheet = pd.read_excel("Summary")
    summary.createChargebacks(wb, chargeback, summarySheet, 0.0305)


# def validateOption (prompt, options):
#     valid = False

#     while not valid:
#         try:
#             option = int(input(prompt))
#             if option in options:
#                 valid = True
#             else:
#                 print("That was an invalid response. Try again.\n")
#         except ValueError:
#             print ("Input a number. Try again.\n")
#     return option


# if __name__ == "__main__":
#     billing = "../archives/2024/February/billing.csv"
#     chargeback = "../archives/2024/February/chargeback.xlsx"

#     resourcegroups.step_one(billing, chargeback)
#     rg.rgComparer(chargeback)

# if __name__ == "__ma2in__":
#     print("")
    
#     archive_directory = "../archives/2024/February"
#     billing = archive_directory + "/billing.csv"
#     chargeback = archive_directory + "/chargeback.xlsx"

#     resourcegroups.step_one(billing, chargeback)
#     rg.rgComparer(chargeback)
#     stepthree.step_three(chargeback, "Feb (2024)")

#     wb = load_workbook(chargeback)
#     sumSheet = pd.read_excel(chargeback, sheet_name = "Summary")
#     summary.creategroupSummarySheet(wb, sumSheet, chargeback)
#     summary.createChargebacks(wb, chargeback, sumSheet, 0.0305)


# if __name__ == "__main__3":
#     excel = "c:\\Users\\do-kelly\\Downloads\\chargeback (1).xlsx"
    
#     option = validateOption("Hi Jason! What are you trying to accomplish:\n (1) Upload billing csv and extract data\n (2) Compare rgs, and get chargeback\n **NOTE: Only do option 2 if you have done option 1**\n", [1, 2])

#     if (option == 1):
#         path = input ("Provide the path to your billing csv file:")
#     elif (option == 2):
#         print("Remember to add last month's rg before providing the path")
#         path = input("Provide the path to the generated excel workbook:")
#         wb = load_workbook(excel)
#         sumSheet = pd.read_excel(excel, sheet_name="Summary")
#         summary.creategroupSummarySheet(wb, sumSheet, excel)
#         # rg.rgComparer(excel)
#         print("Successfully added data analysis and chargeback sheets")

#     # while not(validOption):
#     #     try:
#     #         option = int(input("Hi Jason! What are you trying to accomplish:\n (1) Upload billing csv and extract data\n (2) Compare rgs, and get chargeback\n **NOTE: Only do option 2 if you have done option 1**\n"))
#     #         if (option == 1):
#     #             path = input ("Provide the path to your billing csv file:")
#     #             validOption = True
#     #         elif (option == 2):
#     #             validOption = True
#     #             print("Remember to add last month's rg before providing the path")
#     #             path = input("Provide the path to the generated excel workbook:")
#     #             wb = load_workbook(excel)
#     #             sumSheet = pd.read_excel(excel, sheet_name="Summary")
#     #             summary.creategroupSummarySheet(wb, sumSheet, excel)
#     #             rg.rgComparer(excel)
#     #             print("Successfully added data analysis and chargeback sheets")
#     #         else:
#     #             print("That was an invalid response. Try again")
#     #             validOption = False
#     #     except ValueError:
#     #         print ("Type in a number. Try again")