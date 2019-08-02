"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Runs PV study
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
files = [ "B7SCOPF", "ACTIVSg2000", "ACTIVSg10k", "ACTIVSg25k", "ACTIVSg70k" ]

# Example of file path (%s is where each file will be inserted):
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\synthetic_case\\cases\\%s.pwb"

# Will solve the PV study
def solve_pv(source, sink):
	SimAuto.RunScriptCommand('PVRun([INJECTIONGROUP "%s"], [INJECTIONGROUP "%s"]);' % (source, sink))

# Main function that will run the timeit function on the case
def main(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)

	# Auto insert injection groups, there must be at least 2 to run command
	SimAuto.RunScriptCommand('InjectionGroupsAutoInsert;')
	
	# Create global variables for the source and sink and set each to a
	# random injection group
	global source, sink
	injection_groups = random.sample(SimAuto.ListOfDevices("INJECTIONGROUP", "")[1][0], 2)
	source = injection_groups[0]
	sink = injection_groups[1]

	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit('solve_pv(source, sink)', 'from __main__ import solve_pv, source, sink', number=1)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print("PV:", timing, "sec")
	
	
# Complete task for each case
for file in files:
	print('\n' + file)
	main(file_name % file)