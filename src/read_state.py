"""
Reads the current input vector from an oracle spreadsheet and writes
it to a json file of choice

args:
    oracle
    params list out
    params bounds

"""

import xlwings as xw
import json

# args
in_file = "oracle.xlsx"
out_file = "init.json"
params_file = "params.json"

# load workbook
wb = xw.Book(in_file)
ws = wb.sheets[0]

# load params dict
params = {}
ref_params = json.load(open(params_file))

# discrete params
for key, val in ref_params["Discrete"].items():
    params[key] = ws.range(key).value

# continuous params
for key, val in ref_params["Continuous"].items():
    params[key] = ws.range(key).value

# prints params for confirmation
print(json.dumps(params))

# saves params to a json
with open(out_file, "w") as o:
    o.write(json.dumps(params))
