"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Solves the full Security Constrained OPF and times it
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""

# Find out how to set the area or superarea
# Import timeit library for timing purposes
import timeit
# Import necessary libraries
from decimal import Decimal

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# All the cases to be tested, including only ones with cost curves
files = [ "B7SCOPF", "PSC_37Bus_SCOPF", "ACTIVSg200", "ACTIVSg500", "PSC_2000_DCOPF", "ACTIVSg2000", "ACTIVSg10k", "ACTIVSg25k", "ACTIVSg70k" ]

# Example of file path (%s is where each file will be inserted):
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\synthetic_case\\cases\\%s.pwb"

# Will solve powerflow for SC OPF
def solve_pwrflw():
	SimAuto.RunScriptCommand("SolveFullSCOPF(POWERFLOW);"),

# Will solve optimal powerflow for SC OPF
def solve_opf():
	SimAuto.RunScriptCommand("SolveFullSCOPF(OPF);")

# Main function that will run the timeit function on the case
def solve_full_scopf(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)

	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit(solve_pwrflw, number=5)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print("Powerflow:", timing, "sec")
	
	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit(solve_opf, number=1)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print("OPF:", timing, "sec")
	
# Complete task for each case
for file in files:
	print('\n' + file)
	solve_full_scopf(file_name % file)