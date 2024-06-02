"""
performs the descent on a starting input vector until it plateaus
writes the result to a file of choice

args
    oracle
    param bounds
    otuput file

"""

import xlwings as xw
import matplotlib.pyplot as plt
import json
from save_state import get_state
from save_state import save_state
from descend import descend

# args
oracle = "oracle.xlsx"
params_config = "params.json"
out_file = "state1.json"


def optimize(oracle, params_config, out_file):

    # loads workbook
    wb = xw.Book(oracle)
    ws = wb.sheets[0]

    # vars
    max_iterations = 10
    x_points = []
    y_points = []

    # gets params
    params = get_state(oracle, params_config)

    # gets satisfaction cell reference
    bounds = json.load(open(params_config))
    sat_cell = bounds["Satisfaction"]

    # repeatedly descend until a plateu is reached
    for i in range(max_iterations):

        # descends
        grad = descend(oracle, params, params_config)

        # adds data point to progress graph
        x_points.append(i)
        y_points.append(ws.range(sat_cell).value)

        # checks for a plateau
        if i > 0 and grad == prev_grad:
            break

        # save previous grad
        prev_grad = grad

    # saves the result to file
    save_state(oracle, params_config, out_file)

    # shows progress plot
    plt.plot(x_points, y_points)
    plt.show()


if __name__ == "__main__":
    print("running optimize...")
    optimize(oracle, params_config, out_file)
