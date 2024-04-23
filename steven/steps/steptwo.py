import openpyxl as excel
import pandas as pd

def step_one(billing, chargeback):
    print(billing, chargeback)



if __name__ == "__main__":
    billing = "../../../archives/2024/February/billing.csv"
    chargeback = "../../../archives/2024/February/chargeback.xlsx"

    df = pd.read_csv(billing)
    df = df.filter(items = ["ConsumedService", "ResourceId", "ResourceGroup", "Cost"])
    df.fillna("(blank)", inplace = True)
    df = df.str.strip()

    # ConsumedService == microsoft.visualstudio
    # Assigns RG to organization group found in column ResourceId (AB)
    # Under assumption that ResourceId is formatted: .../organizations/
    df.loc[(df["ConsumedService"] == "microsoft.visualstudio") & (df["ResourceGroup"] == "(blank)"), "ResourceGroup"] = df["ResourceId"].str.split("/").str[-1]

    # For any ResourceGroup that is still blank, output into 
    
    data = data.groupby("ResourceGroup", as_index = False).agg(
        {"ResourceGroup" : "first",
         "Cost" : "sum"}
    )  

    #data = df.loc[df['ResourceGroup'] == '(blank)']
    #data = data.groupby("ConsumedService", as_index = False)
    #data = data.agg({"ConsumedService" : "first", "ResourceGroup" : "count"})

    print(data)

    #    data = data.groupby("RG_Lower", as_index = False).agg(
    #    {"Current Month RG": "first",
    #     "Cost": "sum"})