import csv
import json
import openpyxl

input_file_path = "test.xlsx"  # change this to the path of your input file
output_file_path = "test.json"  # change this to the path of your output file

# Determine the file type based on the file extension
if input_file_path.endswith(".csv"):
    file_type = "csv"
elif input_file_path.endswith((".xlsx", ".xls")):
    file_type = "excel"
else:
    raise ValueError("Unsupported file type")

# Read the data from the input file
if file_type == "csv":
    with open(input_file_path, "r") as csv_file:
        csv_reader = csv.DictReader(csv_file)
        data = list(csv_reader)
elif file_type == "excel":
    workbook = openpyxl.load_workbook(input_file_path)
    worksheet = workbook.active
    headers = [cell.value for cell in worksheet[1]]
    data = []
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        data.append({headers[i]: row[i] for i in range(len(headers))})

# Write the data to the output file in JSON format
with open(output_file_path, "w") as json_file:
    json.dump(data, json_file)
