#Convert Excel File into Json File

import pandas as pd
import json
import sys
import os
from datetime import datetime

# Path to your Excel file
data_path = sys.argv[1]
#data_path = "D:/Atologist/demo_excel/demo_file.xlsx"
#print(f"AAA: {data_path}")

def date_converter(o):
    if isinstance(o, (datetime, pd.Timestamp)):
        return o.strftime("%m-%d-%Y")
    raise TypeError("Type not serializable")

try:
    # Read Excel data into a DataFrame
    df = pd.read_excel(data_path, dtype=str)

    # Ensure all data is read as strings to handle mixed types
    df.fillna('', inplace=True)

    # Replace common null representations with empty strings
    df.replace(['NaN', 'nan', 'null', 'N/A'], '', inplace=True)

    # Get the header row as a list
    headers_list = df.columns.tolist()

    # Get all the data rows as a list of lists
    data_values_list = df.values.tolist()
    data = [headers_list] + data_values_list

    # Serialize data to JSON
    json_data = json.dumps(data, default=date_converter, ensure_ascii=False, separators=(',', ':'))

    # Create a file name based on the input file name
    file_name = os.path.splitext(os.path.basename(data_path))[0]

    # Optional: Save JSON to a file
    #json_file_path = "D:/Atologist/demo_excel/converted_output.json"
    json_file_path = os.path.join(os.path.dirname(data_path), file_name + ".json")

    # Write data to the JSON file
    with open(json_file_path, 'w', encoding='utf-8') as json_file:
        json_file.write(json_data)

    print(f'Excel data from "{data_path}" has been converted to JSON and saved to "{json_file_path}".')
except Exception as e:
    print(f"Error converting Excel to JSON: {str(e)}")