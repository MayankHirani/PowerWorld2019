"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Calculates the TLR values and times it
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""

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

# Create a function that the timeit can call that will calculate the TLR values
def solve(near_bus_num, far_bus_num, circuit, area):
	SimAuto.RunScriptCommand('CalculateTLR([BRANCH "%s" "%s" "%s"], BUYER, [AREA %s]);' % (near_bus_num, far_bus_num, circuit, area))

# Main function that will run the timeit function on the case
def calculate_tlr(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)
	
	# Get a list of branches
	branches = SimAuto.GetParametersMultipleElement("Branch", ["BusNumFrom", "BusNumTo", "LineCircuit"], 0)[1]
	
	# Create variables as global variables so timeit can call them.
	# Choose an area to use (For demo, we choose random one)
	global near_bus, far_bus, circuit, area
	area = random.choice(SimAuto.ListOfDevices("AREA", '')[1][0])

	# Choose a branch (For demo, we choose random one)
	x = random.randint(0, len(branches[0])-1)
	near_bus = branches[0][x].strip()
	far_bus = branches[1][x].strip()
	circuit = branches[2][x].strip()

	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit('solve(near_bus, far_bus, circuit, area)', 'from __main__ import solve, near_bus, far_bus, circuit, area', number=1)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print(timing, "sec")
	
	# Close the case
	SimAuto.CloseCase()

# Complete task for each case
for file in files:
	print('\n' + file)
	calculate_tlr(file_name % file)