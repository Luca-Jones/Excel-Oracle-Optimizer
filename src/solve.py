"""
makes many randomly generated starting input vectors and applies
the descent to them until their output plateaus

Note: it is assumed that the oracle takes an n dimensional
input vector and outputs a number, namely the satisfaction

args
    oracle
    param bounds
    output file
    number of attempts

"""

import xlwings as xw
import matplotlib.pyplot as plt
import json
import randomcolor
import random
import time
from optimize import optimize

# args
oracle = "../in/oracle.xlsx"
params_config = "../in/params.json"
out_file = "../in/state1.json"
N = 25  # number of attempts


def solve(oracle, params_config, out_file, N):

    # vars
    best_vector = {}
    best_sat = 0

    # loads workbook
    wb = xw.Book(oracle)
    ws = wb.sheets[0]

    # initializes the plot
    fig, axs = plt.subplots(1, 1)

    # gets param bounds
    bounds = json.load(open(params_config))

    # get satisfaction cell reference
    sat_cell = bounds["Satisfaction"]

    for i in range(N):

        # generates random starting vector
        while True:

            params = {}

            for key, vals in bounds["Discrete"].items():
                params[key] = vals[random.randrange(0, len(vals))]
                ws.range(key).raw_value = params[key]

            for key, vals in bounds["Continuous"].items():
                params[key] = (
                    random.choice(list(range(0, int((vals[1] - vals[0]) / vals[2]))))
                    * vals[2]
                    + vals[0]
                )
                ws.range(key).raw_value = params[key]

            if ws.range(sat_cell).raw_value > 0:
                break

        # times this procedure
        start_time = time.time()

        # optimizes for this attempt
        x_points, y_points = optimize(oracle, params_config)

        # prints time taken
        print(f"\tTime to complete attempt {i + 1}: {time.time() - start_time} seconds")

        # adds data to a subplot
        axs.plot(
            x_points,
            y_points,
            color=randomcolor.RandomColor().generate()[0],
        )

        # checks if it is the best
        current_sat = ws.range(sat_cell).raw_value
        if current_sat > best_sat:
            best_sat = current_sat
            best_vector = params

    # saves the spreadsheet
    wb.save()
    wb.close()

    # save results to file
    with open(out_file, "w") as o:
        o.write(json.dumps(best_vector))

    # displays the plot
    plt.title(label="Satisfaction vs Iterations")
    plt.xlabel("Number of Iterations")
    plt.ylabel("Average Satisfaction")
    plt.savefig("Optimized_Satisfaction.png")
    plt.show()


if __name__ == "__main__":
    solve(oracle, params_config, out_file, N)
