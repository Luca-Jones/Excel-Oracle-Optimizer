"""
Reads the current input vector from the oracle spreadsheet and writes
it to a json file of choice

"""

import xlwings as xw
import json

file = "RWH.xlsx"
range = ""

# load workbook
wb = xw.Book(file)
ws = wb.sheets["Main"]

params = {}
ref = json.load(open("parameters.json"))

# set all parameters to their default values
for param in ws.range("C2:C32"):
    p_address = param.get_address(row_absolute=False, column_absolute=False)
    if p_address in ref:
        params[p_address] = param.value

# Save the parameters in a json
with open("state0.json", "w") as outfile:
    outfile.write(json.dumps(params))
