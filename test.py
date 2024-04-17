from openpyxl import Workbook, load_workbook
import pandas as pd
import csv
import json
import math


def applicationMerger(excelSheet, applications):
    sumDict = {}
    uniqueApps = []
    for index, value in enumerate(applications, start=2):
        cost = (excelSheet.loc[index,"Unnamed: 47"])
        if math.isnan(cost):
            cost = 0
        else:
            cost = round(cost, 2)
        if value not in uniqueApps:
            uniqueApps.append(value)
            sumDict[value] = cost
        else:
            sumDict[value] += cost
    return sumDict

if __name__ == "__main__":
    excel = "c:\\Users\\do-kelly\\Downloads\\summary.xlsx"
    sheet = "Summary"

    summary = pd.read_excel(excel, sheet_name = sheet)
    applications = summary.loc[2:221, "Unnamed: 0"]
    # print(summary)

    print(applicationMerger(summary, applications))



    #for app in applications:

    # testCol = test["Resource Group"]
    # for value in testCol:
    #     print(value)
   