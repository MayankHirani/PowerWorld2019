"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Determines the length of the path to other objects
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

# Create a function that the timeit can call that will calculate the distance
# from start to end
def solve(bus1, bus2):
	SimAuto.RunScriptCommand("DetermineShortestPath([BUS %s], [BUS %s], Nodes, All, save_file);" % (bus1, bus2))

# Main function that will run the timeit function on the case
def time_calculation(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)

	# Choose a random bus for the starting place, and one for the end.
	# There must be at least 2 buses in the case
	global bus1, bus2
	buses = random.sample(SimAuto.ListOfDevices("BUS", '')[1][0], 2)

	bus1 = buses[0]
	bus2 = buses[1]

	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit('solve(bus1, bus2)', 'from __main__ import solve, bus1, bus2', number=10)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print(timing, "sec")
	
# Complete task for each case
for file in files:
	print('\n' + file)
	time_calculation(file_name % file)