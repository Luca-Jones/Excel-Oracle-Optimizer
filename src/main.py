"""
makes many randomly generated starting input vectors and applies
the descent to them until their output plateaus

Note: it is assumed that the oracle takes an n dimensional
input vector and outputs a number, namely the satisfaction

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
import randomcolor
import random
import time

N = 25  # number of crawlers
S = "K10"  # satisfaction cell reference
grad = ["param", "val"]  # key = CELL, value = CELL_VALUE
iterations = 40  # number of iterations per crawler
use_old = False
presets = [f"s{i}.json" for i in range(1, N + 1)]
# requirements = [[5, 0.7], [10, 0.75], [30, 0.8], [40, 0.85], [45, 0.9]]

# load workbook
file = "RWH.xlsx"
wb = xw.Book(file)
ws = wb.sheets["Main"]

# input / output json
data = [f"s{i}.json" for i in range(1, N + 1)]

# define parameter boundaries
bounds = json.load(open("boundaries.json"))

# initialize the progress plot figure
fig, axs = plt.subplots(1, 1)

# maximum value coordinates
max_sat = 0


def satisfaction():
    av_sat = 0
    station_cell = "C33"
    stations = ["Station 1 2014", "Station 1 2015", "Station 2 2015", "Station 3 2013"]
    for station in stations:
        ws.range(station_cell).raw_value = station
        av_sat += ws.range(S).value
    return av_sat / len(stations)


best_crawler = 0

# iterate over each crawler
for crawler in range(0, N):

    print(f"Crawler {crawler + 1}:")

    # load random initial parameters
    params = {}

    if use_old:
        params = json.load(open(presets[crawler]))

    while not use_old:
        for param, vals in bounds["Discrete"].items():
            params[param] = vals[random.randrange(0, len(vals))]
            ws.range(param).raw_value = params[param]

        for param, vals in bounds["Continuous"].items():
            params[param] = (
                random.choice(list(range(0, int((vals[1] - vals[0]) / vals[2]))))
                * vals[2]
                + vals[0]
            )

            ws.range(param).raw_value = params[param]
            grad[0] = param
            grad[1] = params[param]

        if satisfaction() > 0:
            break

    # track the progress of the crawler
    x_points = []
    y_points = []

    # iteratively calculate and apply the gradient
    for i in range(iterations):

        start_time = time.time()

        # save previous grad
        prev_grad = [grad[0], grad[1]]

        # save the original satisfaction
        current_sat = satisfaction()
        best_sat = current_sat

        for param, val in params.items():

            if param in bounds["Discrete"]:

                for opt in bounds["Discrete"][param]:

                    ws.range(param).raw_value = opt
                    current_sat = satisfaction()
                    if current_sat > best_sat:
                        best_sat = current_sat
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
                    current_sat = satisfaction()
                    if current_sat > best_sat:
                        best_sat = current_sat
                        grad[0] = param
                        grad[1] = val + dx

                if val - dx >= bounds["Continuous"][param][0]:

                    ws.range(param).raw_value = val - dx
                    current_sat = satisfaction()
                    if current_sat > best_sat:
                        best_sat = current_sat
                        grad[0] = param
                        grad[1] = val - dx

                # reset param
                ws.range(param).raw_value = val

        # update ideal solution
        params[grad[0]] = grad[1]
        ws.range(grad[0]).raw_value = grad[1]
        print(f"\t{(grad[0], grad[1])}")

        # Show the Satisfaction

        print(f"\tIteration {i+1}: Satisfaction: {best_sat * 1000 // 1 / 10}%")

        # add data point to progress graph
        x_points.append(i + 1)
        y_points.append(best_sat * 1000 // 1 / 10)

        if best_sat > max_sat:
            max_sat = best_sat
            best_crawler = crawler + 1

        # check to see if there's a plateau
        if grad[0] == prev_grad[0] and grad[1] == prev_grad[1]:
            x_points = x_points + list(range(i + 2, iterations + 1))
            y_points = y_points + [
                best_sat * 1000 // 1 / 10 for j in range(i + 2, iterations + 1)
            ]
            break

        # if i + 2 > 20 and satisfaction() < 0.60:
        #    break

        print(f"\tTime to complete: {time.time() - start_time} seconds")

    axs.plot(
        x_points,
        y_points,
        color=randomcolor.RandomColor().generate()[0],
    )
    # axs.annotate(text=f"s{crawler + 1}.json",xy=(x_points[-1], y_points[-1]),xytext=(x_points[-1] - 0.5, y_points[-1]),arrowprops=dict(),horizontalalignment="left",)

    # Save the parameters in a json
    with open(data[crawler], "w") as outfile:
        outfile.write(json.dumps(params))


# show progress plot
# plt.annotate(
#    text="Our Solution",
#    xy=(iterations, 1000 * max_sat // 1 / 10),
#    xytext=(iterations - 5, 50),
#    arrowprops=dict(),
# )
print(best_crawler)
with open("bestresult.txt", "w") as outfile:
    outfile.write(f"{best_crawler}")
plt.title(label="Satisfaction vs Iterations")
plt.xlabel("Number of Iterations")
plt.ylabel("Average Satisfaction")
plt.savefig("Optimized_Satisfaction.png")
plt.show()
