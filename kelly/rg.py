from openpyxl import load_workbook
import pandas as pd
from .summary import addDataAndHeader

def rgComparer(path):
    """
    Tells you which resource groups were added or deleted from the previous month
    @param path is the path to the excel workbook
    """
    comparison = pd.read_excel(path, sheet_name = "RG Comparison")
    lastMonthRg = comparison["Last Month RG"].tolist()
    currentMonthRg = comparison["Current Month RG"].tolist()

    deleted, added = [], []
    for rgs in lastMonthRg:
        if lastMonthRg not in currentMonthRg:
            deleted.append(rgs)

    for rg in currentMonthRg:
        if currentMonthRg not in lastMonthRg:
            added.append(rg)
    
    wb = load_workbook(path)
    sheet = wb["RG Comparison"]
    addDataAndHeader(wb, sheet, path, 4, "Deleted", deleted)
    addDataAndHeader(wb, sheet, path, 5, "Added", added)