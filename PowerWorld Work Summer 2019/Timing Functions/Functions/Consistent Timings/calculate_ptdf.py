"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Calculates the PTDF values and times it
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""
def CheckResultForError(SimAutoOutput, Message):
    if SimAutoOutput[0] != '':
        print('Error: ' + SimAutoOutput[0])
    else:
        print(Message)

# Import timeit library for timing purposes
import timeit
# Import necessary libraries
from decimal import Decimal
import random

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# All the cases to be tested
files = [ "B7SCOPF", "PSC_2000_DCOPF", "ACTIVSg2000", "ACTIVSg10k", "ACTIVSg25k", "ACTIVSg70k" ]

# Example of file path (%s is where each file will be inserted):
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\synthetic_case\\cases\\%s.pwb"

# Create a function that the timeit can call that will calculate the PTDF values
def solve(seller, buyer):
	# Calculate the PTDF, with the buyer and seller both being different areas
	SimAuto.RunScriptCommand("CalculatePTDF([AREA %s], [AREA %s], AC);" % (seller, buyer))

# Main function that will run the timeit function on the case
def calculate_ptdf(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)
	
	# Choose 2 random areas, one for the seller and one for the buyer.
	# Create them as global variables to be used as arguments in timeit
	global seller, buyer
	areas = random.sample(SimAuto.ListOfDevices("AREA", '')[1][0], 2)
	seller = areas[0]
	buyer = areas[1]

	# Time doing the script command for the number of times specified by number
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(2):
		x = timeit.timeit('solve(seller, buyer)', 'from __main__ import  solve, seller, buyer', number=100)
		timings.append(x)

	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timings[1])
	print(timing, "sec")
	
	# Close the case
	SimAuto.CloseCase()

# Complete task for each case
for file in files:
	print('\n' + file)
	calculate_ptdf(file_name % file)