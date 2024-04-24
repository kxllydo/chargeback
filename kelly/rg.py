from openpyxl import load_workbook
import pandas as pd
from summary import addDataAndHeader

def rgComparer(path):
    """
    Tells you which resource groups were added or deleted from the previous month
    @param path is the path to the excel workbook
    """
    comparison = pd.read_excel(path, sheet_name = "RG Comparison")
    lastMonthRg = comparison["Last Month RG"].tolist()
    currentMonthRg = comparison["Current Month RG"].tolist()

    deleted, added = [], []

    for rg in lastMonthRg:
        if rg not in currentMonthRg:
            deleted.append(rg)

    for rgs in currentMonthRg:
        if rgs not in lastMonthRg:
            added.append(rgs)
    
    wb = load_workbook(path)
    sheet = wb["RG Comparison"]
    addDataAndHeader(wb, sheet, path, 4, "Deleted", 45, deleted)
    addDataAndHeader(wb, sheet, path, 5, "Added", 45, added)