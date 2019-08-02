"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Times the fastest creation of a new case
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""

# Import timeit library for timing purposes
import timeit
# Import necessary libraries
from decimal import Decimal

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Create a function that the timeit can call that will create a case
def create_case():
	SimAuto.RunScriptCommand("NewCase;")

# Main function that will run the timeit function on the case
def time_creation():

	# Time doing the script command for the number of times specified by number.
	# The first time can be inaccurate, so we take the second
	timings = []
	for num in range(10):
		x = timeit.timeit(create_case, number=1)
		timings.append(x)

	# Use Decimal to display the fastest timing in scientific notation
	# for readability
	timing = '%.4E' % Decimal(min(timings))
	print(timing, "sec")

# Run the function
time_creation()