"""
Reads the current input vector from an oracle spreadsheet and writes
it to a json file of choice

args:
    oracle
    params vector
    params bounds

"""

import xlwings as xw
import json

# args
oracle = "oracle.xlsx"
out_file = "init.json"
params_config = "params.json"


def get_state(oracle, params_config):

    # loads workbook
    wb = xw.Book(oracle)
    ws = wb.sheets[0]

    # params dict
    params = {}
    ref_params = json.load(open(params_config))

    # loads discrete params
    for key, val in ref_params["Discrete"].items():
        params[key] = ws.range(key).value

    # loads continuous params
    for key, val in ref_params["Continuous"].items():
        params[key] = ws.range(key).value

    return params


def save_state(oracle, params_config, out_file):

    # gets the current params dict from the spreadsheet
    params = get_state(oracle, params_config)

    # prints params for confirmation
    print(json.dumps(params))

    # saves params to a json
    with open(out_file, "w") as o:
        o.write(json.dumps(params))


if __name__ == "__main__":
    print("running save_state...")
    save_state(oracle, params_config, out_file)
