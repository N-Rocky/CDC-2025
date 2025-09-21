import pandas as pd
import json

path = "dataset.xlsx"
spreadsheet = pd.ExcelFile(path)

# Skip the first sheet (ReadMe)
sheets = spreadsheet.sheet_names[1:] # Excluding ReadMe sheet
tableNames = [f"Table{index + 1}" for index in range(len(sheets))]

tables = {}

for index, sheet in enumerate(sheets):
    df = pd.read_excel(path, sheet_name=sheet, header=None, skiprows=5, nrows=100)

    # Disregard empty cells
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)

    years = df.iloc[0, 2:].astype(str).tolist()
    categories = df.iloc[1:, 1].fillna("Unknown")

    table = {}
    for col, year in enumerate(years, start=2):
        yearData = {}
        for row, category in enumerate(categories, start=1):
            value = df.iloc[row, col]
            if pd.api.types.is_number(value):
                yearData[str(category).strip()] = float(value)
        table[year] = yearData
    tables[tableNames[index]] = table

with open("data.json", "w") as file:
    json.dump(tables, file, indent=2)

print("Converted to json")