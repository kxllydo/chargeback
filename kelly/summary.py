from openpyxl import Workbook, load_workbook
import pandas as pd
import csv
import json
import math

def addDataAndHeader (wb, ws, path, columnNum, header, width = 0, dataList = []):
    '''
    This adds all of the headers to the group summary sheet
    @param wb is the workbook loaded using openpyxl
    @param ws is the worksheet opened using openpyxl
    @path is a string of the path to the excel workbook
    @columnNum is the column you want to add the header to
    @header is the header you want to add to the excel sheet
    '''
    cell = ws.cell(row = 1, column = columnNum)
    cell.value = header
    for index, value in enumerate(dataList, start=2):
        cell = ws.cell(row=index, column=columnNum)
        cell.value = value

    if width != 0:
       ws.column_dimensions[chr(64 + columnNum)].width = width 

    wb.save(path)

def groupCostMerger(sheet):
    """
    Combines the same groups and sums their costs
    @param sheet is the sheet read in by using pandas
    @return a dictionary of the groups and their costs
    """
    colLength = len(sheet["Group"]) - 1
    applications = sheet.loc[0:colLength, "Group"]
    sumDict = {}
    uniqueApps = []

    for index, value in enumerate(applications, start = 0):
        cost = sheet.loc[index,"February (2024)"]
        if math.isnan(cost):
            cost = 0
        if value not in uniqueApps:
            uniqueApps.append(value)
            sumDict[value] = cost
        else:
            sumDict[value] += cost
    return sumDict

def merger(sheet, header):
    """
    General merger to assign relationship between non cost categories
    @param sheet is the sheet read in by using pandas
    @header is the header of the data you want to access
    @return a dictionary of the group and other dataset
    """
    colLength = len(sheet["Group"]) - 1
    applications = sheet.loc[0:colLength, "Group"]
    dict = {}
    uniqueApps = []

    for index, value in enumerate(applications, start = 0):
        if value not in uniqueApps:
            dict[value] = sheet.loc[index, header]
            uniqueApps.append(value)
    return dict

def creategroupSummarySheet(wb, ws, path):
    """
    Creates the group summary sheet and fills it with cost, PC, and AC
    @param wb is the workbook loaded using openpyxl
    @param path is a string path to the workbook
    """
    pc = merger(ws, "PC")
    ac = merger(ws, "AC")
    cost = groupCostMerger(ws)

    wb.create_sheet("Group Summary")
    ws = wb["Group Summary"]
    headers = ["Applications", "Amount", "Infra Charge", "Adjustments", "Total Charge",	"Profit Center", "Account Code"]
    
    addDataAndHeader(wb, ws, path, 1, headers[0], 49, list(cost))
    addDataAndHeader(wb, ws, path, 2, headers[1], 14.55, list(cost.values()))
    addDataAndHeader(wb, ws, path, 3, headers[2], 12.64)
    addDataAndHeader(wb, ws, path, 4, headers[3], 13.91)
    addDataAndHeader(wb, ws, path, 5, headers[4], 16.64)
    addDataAndHeader(wb, ws, path, 5, headers[5], 15.45, list(pc.values()))
    addDataAndHeader(wb, ws, path, 6, headers[6], 15.45, list(ac.values()))
