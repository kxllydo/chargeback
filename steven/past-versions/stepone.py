from helper import check_for_perms
import openpyxl as excel
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

def stepone(billing, chargeback):
    check_for_perms(billing)
    check_for_perms(chargeback)

    data = pd.read_csv(billing)
    data = data.filter(items = ["ConsumedService", "ResourceId", "ResourceGroup", "Cost"])
    data.fillna("(blank)", inplace = True)
    data.loc[(data["ConsumedService"] == "microsoft.visualstudio") & (data["ResourceGroup"] == "(blank)"), "ResourceGroup"] = data["ResourceId"].str.split("/").str[-1]
    data["RG_Lower"] = data["ResourceGroup"].str.lower()
    data.rename(columns = {"ResourceGroup" : "Current Month RG"}, inplace = True)
    data = data.groupby("RG_Lower", as_index = False).agg(
        {"Current Month RG" : "first",
         "Cost" : "sum"})
    print(type(data))

    df = pd.read_excel(chargeback, sheet_name = "RG Comparison")
    
    return df