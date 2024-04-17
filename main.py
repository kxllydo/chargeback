import csv
import json

def csv_to_json(csvf, jsonf):
    '''
        Pre-processing of data to clean up irrelevant fields.
        Only relevant fields are: Resource Group, Cost, and Tags.

        Args:
        
            csvf (string)   : Path to .CSV file to be read
            jsonf (string)  : Path to .JSON file to be written

        Returns:
            jsonf
    '''

    with open(csvf, 'r', encoding = 'utf-8') as csvf_o:
        with open(jsonf, 'w', encoding = 'utf-8') as jsonf_o:
            json_array = []
            csvr = csv.DictReader(csvf_o)

            for row in csvr:
                # print(row)
                # new_row = {}
                # new_row['RG'] = row['Resource Group']
                # new_row['Cost'] = row['Cost']
                # new_row['Tags'] = row['Tags']

                json_array.append(row)
                print(row)

            jsonString = json.dumps(json_array, indent = 4)
            jsonf_o.write(jsonString)

if __name__ == "__main__":
    csv_file = input("Input the CSV file path\n") #r"./billing.csv"
    json_file = input("Input the JSON file path\n") #r"./resource-group-costs-json.txt"

    csv_to_json(csv_file, json_file)
