"""
We will assume that each of the parameters is roughly independent
Thus, we will use the gradient approach.
If it fails, we will use greedy algorithm :(

Things to edit for the specific application:
- state0.json
- boundaries.json
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

S = "K10"  # satisfaction cell reference
# dx = 0.2  # continuous value step

x_points = []
y_points = []

# define parameter boundaries
bounds = json.load(open("boundaries.json"))

# load initial parameters
params = json.load(open("state0.json"))
for param, val in params.items():
    ws.range(param).raw_value = val
grad = {}  # key = CELL, value = CELL_VALUE

# Show the Satisfaction
print(f"Initial Satisfaction: {ws.range(S).value * 1000 // 1 / 10}%")

# iteratively calculate and apply the gradient
for i in range(20):

    for param, val in params.items():

        # save the original satisfaction
        current_sat = ws.range(S).value

        if param in bounds["Discrete"]:

            # Calculate discrete component of the gradient
            best_sat = current_sat
            choice = val

            for opt in bounds["Discrete"][param]:

                ws.range(param).raw_value = opt
                if ws.range(S).value > best_sat:
                    choice = opt
                    best_sat = ws.range(S).value

            # reset value
            ws.range(param).raw_value = val

            # record the new value
            grad[param] = choice

        elif param in bounds["Continuous"]:

            # Calculate continuous gradient
            dx = bounds["Continuous"][param][2]
            df = 0
            best_sat = current_sat

            # compare +dx, 0, and -dx
            if val + dx <= bounds["Continuous"][param][1]:

                ws.range(param).raw_value = val + dx
                if ws.range(S).value > best_sat:
                    df = dx
                    best_sat = ws.range(S).value

            if val - dx >= bounds["Continuous"][param][0]:

                ws.range(param).raw_value = val - dx
                if ws.range(S).value > best_sat:
                    df = -1 * dx

            # reset param
            ws.range(param).raw_value = val

            # record the gradient
            grad[param] = val + df

    # Nudge parameters
    for param in params:
        params[param] = grad[param]
        ws.range(param).raw_value = grad[param]

    # Show the Satisfaction
    print(f"Iteration {i+1}: Satisfaction: {ws.range(S).value * 1000 // 1 / 10}%")

    # add data point to progress graph
    x_points.append(i)
    y_points.append(ws.range(S).value * 1000 // 1 / 10)

# Save the parameters in a json
with open("parameters.json", "w") as outfile:
    outfile.write(json.dumps(params))

# show progress plot
plt.plot(x_points, y_points)
plt.show()
