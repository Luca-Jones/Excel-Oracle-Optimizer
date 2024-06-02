"""
performs a step of the descent on a starting input vector
for a given oracle

args
    oracle
    input vector
    param bounds

"""

import xlwings as xw
import matplotlib.pyplot as plt
import json
from save_state import save_state

# args
oracle = "oracle.xlsx"
in_file = "init.json"
params_config = "params.json"


def descend(oracle, params, params_config):

    # loads workbook
    wb = xw.Book(oracle)
    ws = wb.sheets[0]

    # defines parameter boundaries
    bounds = json.load(open(params_config))

    # key = CELL, value = CELL_VALUE
    grad = list(next(iter(params.items())))

    # saves the original satisfaction
    sat_cell = bounds["Satisfaction"]
    current_sat = ws.range(sat_cell).value
    best_sat = current_sat

    for param, val in params.items():

        if param in bounds["Discrete"]:

            # tries each option
            for opt in bounds["Discrete"][param]:

                ws.range(param).raw_value = opt
                if ws.range(sat_cell).value > best_sat:
                    best_sat = ws.range(sat_cell).value
                    grad[0] = param
                    grad[1] = opt

            # resets param
            ws.range(param).raw_value = val

        elif param in bounds["Continuous"]:

            # gets continuous step
            dx = bounds["Continuous"][param][2]

            # compares +dx, 0, and -dx
            if val + dx <= bounds["Continuous"][param][1]:

                ws.range(param).raw_value = val + dx
                if ws.range(sat_cell).value > best_sat:
                    best_sat = ws.range(sat_cell).value
                    grad[0] = param
                    grad[1] = val + dx

            if val - dx >= bounds["Continuous"][param][0]:

                ws.range(param).raw_value = val - dx
                if ws.range(sat_cell).value > best_sat:
                    best_sat = ws.range(sat_cell).value
                    grad[0] = param
                    grad[1] = val - dx

            # resets param
            ws.range(param).raw_value = val

    # update ideal solution
    params[grad[0]] = grad[1]
    print(f"\t{(grad[0], grad[1])}\t{current_sat} --> {best_sat}")
    ws.range(grad[0]).raw_value = grad[1]

    return grad[0], grad[1]


if __name__ == "__main__":
    print("running descend...")
    params = json.load(open(in_file))
    descend(oracle, params, params_config)
    save_state(oracle, out_file=in_file, params_config=params_config)
