"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Calculates the LODF values and times it
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""

# Find out what argument to use for CalculateLODF
# Import timeit library for timing purposes
import timeit
# Import necessary libraries
from decimal import Decimal
import random

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# All the cases to be tested
files = [ "WSCC 9 bus", "ACTIVSg200", "ACTIVSg500", "ACTIVSg2000", "ACTIVSg10k", "ACTIVSg25k", "ACTIVSg70k" ]

# Example of file path (%s is where each file will be inserted):
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\synthetic_case\\cases\\%s.pwb"

# Create a function that the timeit can call that will solve the power flow
def solve(near_bus, far_bus, circuit):
	SimAuto.RunScriptCommand('CalculateLODF([BRANCH "%s" "%s" "%s"], DC);' % (near_bus, far_bus, circuit))

# Main function that will run the timeit function on the case
def time_lodf(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)

	# Choose a random branch, there must be at least 1 in the case
	branches = SimAuto.GetParametersMultipleElement("Branch", ["BusNumFrom", "BusNumTo", "LineCircuit"], 0)[1]
	global near_bus, far_bus, circuit
	x = random.randint(0, len(branches[0])-1)
	near_bus = branches[0][x].strip()
	far_bus = branches[1][x].strip()
	circuit = branches[2][x].strip()
	
	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit('solve(near_bus, far_bus, circuit)', 'from __main__ import solve, near_bus, far_bus, circuit', number=1)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print(timing, "sec")

# Complete task for each case
for file in files:
	print('\n' + file)
	time_lodf(file_name % file)
	