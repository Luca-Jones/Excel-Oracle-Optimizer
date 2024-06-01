"""
We will assume that each of the parameters is roughly independent
Thus, we will use the gradient approach.
If it fails, we will use greedy algorithm :(

Things to edit for the specific application:
- state0.json
- boundaries.json
- output json files
- S (output reference)
- dx (discrete gradient step)
- file (spreadsheet name)

"""

import xlwings as xw
import matplotlib.pyplot as plt
import json

# load workbook
file = "RWH.xlsx"
wb = xw.Book(file)
ws = wb.sheets["Main"]

# input / output json
data = "s12.json"
params = json.load(open(data))
bounds = json.load(open("boundaries.json"))
parameter = "C29"
units = ""
original_value = params[parameter]
S = "K10"  # satisfaction cell reference
x_points = []
y_points = []

# load initial values
for param, val in params.items():
    ws.range(param).raw_value = val

# save the original satisfaction
original_sat = ws.range(S).value

if parameter in bounds["Discrete"]:

    for opt in bounds["Discrete"][parameter]:
        ws.range(parameter).raw_value = opt

        # plot point
        x_points.append(opt)
        y_points.append(ws.range(S).value)
        print((x_points, y_points))

    # plot bar chart
    plt.bar(x=x_points, height=y_points, width=0.8, align="center")


elif parameter in bounds["Continuous"]:
    vals = bounds["Continuous"][parameter]
    for val in range(vals[0], vals[1] + 1, vals[2]):
        ws.range(parameter).raw_value = val

        # plot point
        x_points.append(val)
        y_points.append(ws.range(S).value)

    # plot chart
    plt.plot(x_points, y_points)

plt.annotate(
    text=f"({original_value} {units}, {original_sat * 1000 // 1 / 10} %)",
    xy=(original_value, original_sat),
    xytext=(original_value, original_sat + 0.1),
    horizontalalignment="center",
    verticalalignment="top",
    arrowprops=dict(facecolor="black"),
)
plt.show()
