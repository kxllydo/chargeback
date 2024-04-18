from openpyxl import Workbook, load_workbook
import pandas as pd
import csv
import json
import math

def addDataColumn (wb, ws, path, dataList, columnNum, header):
    cell = ws.cell(row = 1, column = columnNum)
    cell.value = header
    for index, value in enumerate(dataList, start=2):
        cell = ws.cell(row=index, column=columnNum)
        cell.value = value
    wb.save(path)

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

def rgComparer():
    excel = "c:\\Users\\do-kelly\\Downloads\\help.xlsx"
    comparison = pd.read_excel(excel, sheet_name = "RG Comparison")
    lastMonthRg = comparison["Last Month RG"].tolist()
    currentMonthRg = comparison["Current Month RG"].tolist()

    print(comparison.columns)
    deleted = []
    added = []
    for rgs in lastMonthRg:
        if lastMonthRg not in currentMonthRg:
            deleted.append(rgs)

    for rg in currentMonthRg:
        if currentMonthRg not in lastMonthRg:
            added.append(rg)
    data = {'Deleted' : deleted,
            "Added" : added}
    
    dataFrame = pd.DataFrame(data)
    
    wb = load_workbook(excel)
    sheet = wb["RG Comparison"]
    addDataColumn(wb, sheet, excel, deleted, 4, "Deleted")
    addDataColumn(wb, sheet, excel, added, 5, "Added")
    # cell = sheet.cell(row = 1, column = 4)
    # cell.value = "Deleted"
    # for index, value in enumerate(deleted, start=2):
    #     cell = sheet.cell(row=index, column=4)
    #     cell.value = value

    # wb.save(excel)







if __name__ == "__main__":
    # excel = "c:\\Users\\do-kelly\\Downloads\\summary.xlsx"
    # sheet = "Summary"

    # summary = pd.read_excel(excel, sheet_name = sheet)
    # applications = summary.loc[2:221, "Unnamed: 0"]
    # print(summary)

    rgComparer()
    # print(applicationMerger(summary, applications))
