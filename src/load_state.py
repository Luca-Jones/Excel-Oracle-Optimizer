"""
loads a state json file into the oracle spreadsheet

"""

import xlwings as xw
import json

file = "RWH.xlsx"

# load workbook
wb = xw.Book(file)
ws = wb.sheets["Main"]


def satisfaction():
    av_sat = 0
    station_cell = "C33"
    stations = ["Station 1 2014", "Station 1 2015", "Station 2 2015", "Station 3 2013"]
    for station in stations:
        ws.range(station_cell).raw_value = station
        av_sat += ws.range("K10").value
    return av_sat / len(stations)


max_sat = 0
max_index = 0

for i in range(0):

    params = json.load(open(f"s{i}.json"))

    # set all parameters to their default values
    for param, val in params.items():
        ws.range(param).raw_value = val

    # retrieve the output value
    output = satisfaction()
    if output > max_sat:
        max_sat = output
        max_index = i
    print(output)


params = json.load(open("b10.json"))

# set all parameters to their default values
for param, val in params.items():
    ws.range(param).raw_value = val

# retrieve the output value
output = satisfaction()
print(output)
