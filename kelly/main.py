import summary, rg
from openpyxl import Workbook, load_workbook
import pandas as pd
import csv
import json
import math

if __name__ == "__main__":
    excel = "c:\\Users\\do-kelly\\Downloads\\chargeback (1).xlsx"

    wb = load_workbook(excel)
    sumSheet = pd.read_excel(excel, sheet_name="Summary")
    summary.creategroupSummarySheet(wb, sumSheet, excel)
    # rg.rgComparer(excel)

