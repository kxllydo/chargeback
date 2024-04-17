import pandas
import openpyxl
import csv, json

if __name__ == "__main__":
    billing_file_path = "../../hide/billing.csv"

    df = pandas.read_csv(billing_file_path, na_values = "(blank)").filter(items = ["ResourceGroup", "Cost"])
    df.fillna("(blank)", inplace = True)
    df["RG_Lower"] = df["ResourceGroup"].str.lower()

    df = df.groupby("RG_Lower", as_index = False).agg({"ResourceGroup": "first", "Cost": "sum"})
    df.drop(columns = "RG_Lower", inplace = True)
    df.to_excel("../../hide/PivotTable_RG-Cost.xlsx", index = False)



if __name__ == "__main9__":
    billing_file_path = "../../hide/billing.csv"

    with open(billing_file_path, "r", encoding = "utf-8-sig") as csvf:
        csvr = csv.DictReader(csvf)

        infodump = {}

        for row in csvr:
            row_k = row["ResourceGroup"].lower()
            if not row_k: row_k = "(blank)"

            if row_k not in infodump:
                infodump[row_k] = "---"

            if row["Tags"]:
                asda = json.loads("{" + row["Tags"] + "}")
                if "Profit Center" in asda.keys():
                    infodump[row_k] = asda["Profit Center"]

        with open(billing_file_path + "aads.txt", "w", encoding = "utf-8-sig") as writable:
            for key, val in infodump.items():
                writable.write(f"{key} -> {val}\n")