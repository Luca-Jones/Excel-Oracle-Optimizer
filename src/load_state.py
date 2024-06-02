"""
loads an input vector from a json file into the oracle spreadsheet

args:
    oracle
    params list

"""

import xlwings as xw
import json

# args
oracle = "oracle.xlsx"
in_file = "init.json"

# load workbook
wb = xw.Book(oracle)
ws = wb.sheets[0]

# load params
params = json.load(open(in_file))

# set all param values
for key, val in params.items():
    ws.range(key).raw_value = val
