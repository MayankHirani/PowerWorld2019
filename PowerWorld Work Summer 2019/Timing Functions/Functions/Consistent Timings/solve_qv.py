"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Description: Runs the QV study
Includes: Setup, OpenCase(), RunScriptCommand(), timeit, CloseCase()
"""

# Import timeit library for timing purposes
import timeit
# Import necessary libraries
from decimal import Decimal

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# All the cases to be tested
files = [ "WSCC 9 bus", "ACTIVSg200", "ACTIVSg500", "ACTIVSg2000", "ACTIVSg10k", "ACTIVSg25k", "ACTIVSg70k" ]

# Example of file path (%s is where each file will be inserted):
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\synthetic_case\\cases\\%s.pwb"

# Will solve the QV study
def solve_qv(save_file):

	SimAuto.RunScriptCommand("QVRun(%s, NO);" % save_file)

# Main function that will run the timeit function on the case
def main(file_name):

	# Open the case
	SimAuto.OpenCase(file_name)
	
	# A global variable of the file that QV results will be saved to
	global save_file
	save_file = 'c:/Users/mayank/Desktop/mayank/synthetic_case/cases/save_file.aux'

	# Time the QV solve and use the argument False to specify that we do
	# not want to save the results
	timing = timeit.timeit('solve_qv(save_file)', 'from __main__ import solve_qv, save_file', number=2)
	
	# Use Decimal to display the timing in scientific notation for readability
	timing = '%.4E' % Decimal(timing)
	print("QV:", timing, "sec")

# Complete task for each case
for file in files:
	print('\n' + file)
	main(file_name % file)