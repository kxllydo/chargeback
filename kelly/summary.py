from openpyxl import Workbook, load_workbook
from openpyxl.styles import NamedStyle
import pandas as pd
import csv
import json
import math
import time

def addDataAndHeader (wb, ws, path, columnNum, header, width = 0, dataList = []):
    '''
    This adds all of the headers to the group summary sheet
    @param wb is the workbook loaded using openpyxl
    @param ws is the worksheet opened using openpyxl
    @param path is a string of the path to the excel workbook
    @param columnNum is the column you want to add the header to
    @param header is the header you want to add to the excel sheet
    @param width is the width of the column
    @param dataList is the list of the data you want in the column
    @return completed data and header
    '''
    cell = ws.cell(row = 1, column = columnNum)
    cell.value = header
    for index, value in enumerate(dataList, start=2):
        cell = ws.cell(row=index, column=columnNum)
        cell.value = value

    if width != 0:
       ws.column_dimensions[chr(64 + columnNum)].width = width 

    if (len(dataList) != 0 ) and (isinstance(dataList[0], float)):
        # if 'number_format' in wb.named_styles:
        #     number_format = wb.named_styles['number_format']
        # else:
        #     # Create a new named style
        #     number_format = NamedStyle(name='number_format')
        #     wb.add_named_style(number_format)

        for index, value in enumerate(dataList, start = 2):
            cell = ws.cell(row=index, column=columnNum)
            cell.number_format = "0.00"

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


def infracharge(total):
    """
    Finds the value of the infracharge
    @param sumSheet is the summary sheet opened with pandas
    @param total is the dictionary of all the groups and the total cost
    """
    groupNum = len(total) - 1
    print(groupNum)
    cost = round((total["Infrastructure"]/groupNum), 2)
    length = len(total.values())
    data  = []
    for i in range(length):
        data.append(cost)
    return data

def addCharges(cost, infracharge):
    """
    Adds sales tax 3.05% and infracharge to the total costs
    @param cost is the dictionary of groups and total costs
    @param infracharge is the list of the infracharge
    """
    infra = infracharge[0]
    for key in cost:
        originalCost = cost[key]
        taxes = originalCost * .0305
        cost[key] += taxes + infra
    
    return cost

def creategroupSummarySheet(wb, sumSheet, path):
    """
    Creates the group summary sheet and fills it with cost, PC, and AC
    @param wb is the workbook loaded using openpyxl
    @param ws is the summary worksheet opened with pandas
    @param path is a string path to the workbook
    """
    pc = merger(sumSheet, "PC")
    ac = merger(sumSheet, "AC")
    cost = groupCostMerger(sumSheet)
    charge = infracharge(cost)

    wb.create_sheet("Group Summary")
    ws = wb["Group Summary"]
    headers = ["Applications", "Amount", "Infra Charge", "Adjustments", "Total Charge",	"Profit Center", "Account Code"]
    
    addDataAndHeader(wb, ws, path, 1, headers[0], 49, list(cost))
    addDataAndHeader(wb, ws, path, 2, headers[1], 14.55, list(cost.values()))
    addDataAndHeader(wb, ws, path, 3, headers[2], 12.64, charge)
    addDataAndHeader(wb, ws, path, 4, headers[3], 13.91)
    addDataAndHeader(wb, ws, path, 5, headers[4], 16.64)
    addDataAndHeader(wb, ws, path, 5, headers[5], 15.45, list(pc.values()))
    addDataAndHeader(wb, ws, path, 6, headers[6], 15.45, list(ac.values()))

def createChargeback(wb, sumSheet, path):
    """
    Creates chargeback sheet for the month
    @param is the workbook opened by openpyxl
    @param is the summary sheet opened by pandas
    """
    pc = merger(sumSheet, "PC")
    ac = merger(sumSheet, "AC")
    cost = groupCostMerger(sumSheet)
    charge = infracharge(cost)
    owner = merger(sumSheet, "Owner")

    finalCosts = addCharges(cost, charge)

    finalCosts.pop("Infrastructure")
    pc.pop("Infrastructure")
    ac.pop("Infrastructure")
    owner.pop("Infrastructure")

    wb.create_sheet("Customer Chargeback")

    ws = wb["Customer Chargeback"]
    headers = ["Owner","Applications","February (2024)","Profit Center", "AC"]
    addDataAndHeader(wb, ws, path, 1, headers[0], 31.64, list(owner.values()))
    addDataAndHeader(wb, ws, path, 2, headers[1], 32.64,  list(owner))
    addDataAndHeader(wb, ws, path, 3, headers[2], 19, list(finalCosts.values()))
    addDataAndHeader(wb, ws, path, 4, headers[3], 15.91, list(pc.values()))
    addDataAndHeader(wb, ws, path, 5, headers[4], 16.64, list(ac.values()))
