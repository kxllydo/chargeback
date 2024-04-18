import pandas
import openpyxl
import os
import datetime, time

if __name__ == "__main__":
    #billing_file_path = None
    #while not billing_file_path:
    #    billing_file_path = input("Location of billing.csv:\t")
    #    
    #    if not os.path.exists(billing_file_path) or not os.access(billing_file_path, os.R_OK):
    #        billing_file_path = None

    #current_date = datetime.datetime.now()
    #month = current_date.strftime("%B")
    #year = current_date.strftime("%Y")

    billing_file_path = "../../hide/billing.xlsx"   ##############
    month = "February"                              # TEMPORARY! #
    year = "2024"                                   ##############

    destination_path = billing_file_path

    xl_wb = openpyxl.load_workbook(destination_path)
    rg_cost_sheet = xl_wb.create_sheet("RG Comparison")
    xl_wb.active = rg_cost_sheet

    startt = time.time()
    xl_wb.save(destination_path)
    print(f"{time.time() - startt} to SAVE excel shett\n")


    df = pandas.DataFrame(xl_wb["billing"].values)

    print(df.head(5))

    df = df.filter(items = ["ResourceGroup", "Cost"])

    print(df.head(5))

    #df.fillna("(blank)", inplace = True)
    #df["RG_Lower"] = df["ResourceGroup"].str.lower()
    #df["Last Month RG"] = df["ResourceGroup"]
    #df["Current Month RG"] = df["ResourceGroup"]

    #df = df.groupby("RG_Lower", as_index = False).agg({"Last Month RG": "first", "Current Month RG": "first", "Cost": "sum"})
    #df.drop(columns = "RG_Lower", inplace = True)

    #for row in df.head(10):
    #    rg_cost_sheet.append(row)
    #df.to_excel("../../hide/PivotTable_RG-Cost.xlsx", index = False)

    #with pandas.ExcelWriter(destination_path, engine = "openpyxl") as writer:
    #    #writer.book = xl_wb
    #   df.to_excel(writer, sheet_name="RG Comparison", index = False)
    
    xl_wb.close()