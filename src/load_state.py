"""
loads an input vector from a json file into the oracle spreadsheet

args:
    oracle
    input vector

"""

import xlwings as xw
import json

# args
oracle = "../in/oracle.xlsx"
in_file = "../in/init.json"


def load_state(oracle, in_file):

    # loads workbook
    wb = xw.Book(oracle)
    ws = wb.sheets[0]

    # loads params
    params = json.load(open(in_file))

    # sets all param values
    for key, val in params.items():
        ws.range(key).raw_value = val


if __name__ == "__main__":
    print("running load_state...")
    load_state(oracle, in_file)
