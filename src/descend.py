"""
performs the descent on a starting input vector until it plateaus

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
data_in = "s18.json"
data_out = "state1.json"
iterations = 5

S = "K10"  # satisfaction cell reference
# dx = 0.2  # continuous value step

x_points = []
y_points = []

# define parameter boundaries
bounds = json.load(open("boundaries.json"))

# load initial parameters
params = json.load(open(data_in))
grad = ["param", "val"]  # key = CELL, value = CELL_VALUE

for param, val in params.items():
    grad[0] = param
    grad[1] = val
    ws.range(param).raw_value = val

# Show the Satisfaction
print(f"Initial Satisfaction: {ws.range(S).value * 1000 // 1 / 10}%")

# iteratively calculate and apply the gradient
for i in range(iterations):

    # save the original satisfaction
    current_sat = ws.range(S).value
    best_sat = current_sat

    for param, val in params.items():

        if param in bounds["Discrete"]:

            for opt in bounds["Discrete"][param]:

                ws.range(param).raw_value = opt
                if ws.range(S).value > best_sat:
                    best_sat = ws.range(S).value
                    grad[0] = param
                    grad[1] = opt

            # reset value
            ws.range(param).raw_value = val

        elif param in bounds["Continuous"]:

            # Calculate continuous gradient
            dx = bounds["Continuous"][param][2]

            # compare +dx, 0, and -dx
            if val + dx <= bounds["Continuous"][param][1]:

                ws.range(param).raw_value = val + dx
                if ws.range(S).value > best_sat:
                    best_sat = ws.range(S).value
                    grad[0] = param
                    grad[1] = val + dx

            if val - dx >= bounds["Continuous"][param][0]:

                ws.range(param).raw_value = val - dx
                if ws.range(S).value > best_sat:
                    best_sat = ws.range(S).value
                    grad[0] = param
                    grad[1] = val - dx

            # reset param
            ws.range(param).raw_value = val

    # update ideal solution
    params[grad[0]] = grad[1]
    print(f"\t{(grad[0], grad[1])}")
    ws.range(grad[0]).raw_value = grad[1]

    # Show the Satisfaction
    print(f"Iteration {i+1}: Satisfaction: {ws.range(S).value * 1000 // 1 / 10}%")

    # add data point to progress graph
    x_points.append(i)
    y_points.append(ws.range(S).value * 1000 // 1 / 10)

# Save the parameters in a json
with open(data_out, "w") as outfile:
    outfile.write(json.dumps(params))

# show progress plot
plt.plot(x_points, y_points)
plt.show()
