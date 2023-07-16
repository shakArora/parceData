import openpyxl
import json

# Load the workbook
workbook = openpyxl.load_workbook("data.xlsx")

# Select the active sheet
sheet = workbook.active

# Read the existing data from data.json
with open("data.json") as json_file:
    data = json.load(json_file)

# Iterate over the rows in the sheet starting from the second row
for row in sheet.iter_rows(min_row=2, values_only=True):
    date = row[0].strftime('%Y-%m-%d')
    event_type = row[1]
    event_comment = row[2] if row[2] else ""

    # Check if the date already exists in the data dictionary
    if date in data:
        # Update the event type and comment
        data[date]["type"] = event_type
        data[date]["comment"] = event_comment
    else:
        # Add a new entry for the date
        data[date] = {
            "date": date,
            "type": event_type,
            "comment": event_comment
        }

# Write the updated data back to data.json
with open("data.json", "w") as json_file:
    json.dump(data, json_file, indent=2)
