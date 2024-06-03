from solve import solve

# oracle
print("Enter the location of the oracle (Excel Spreadsheet): ")
oracle = input()

try:
    open(oracle)
except:
    print("File does not exist")
    exit(0)

# bounds
print("Enter the location of the input boundaries json file: ")
params_config = input()

try:
    open(params_config)
except:
    print("File does not exist")
    exit(0)

# output file
print("Enter the name of the json file to save the results to: ")
out_file = input()

# attempts
print("Enter the number of attempts to optimize: ")

try:
    N = int(input())
except:
    print("Please enter a number greater than 0.")
    exit(0)

try:
    solve(oracle, params_config, out_file, N)
except:
    print(
        "Something went wrong. Please make sure that the arguments are entered correctly."
    )
    exit(0)

exit(0)
