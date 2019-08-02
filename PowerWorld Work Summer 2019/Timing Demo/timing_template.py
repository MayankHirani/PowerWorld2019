"""
Template for timing a function for SimAuto in Python
Developed by PowerWorld 2019, created by Mayank Hirani
For an example of how to use this, see timing_template_example.py
"""


# Import the timeit library
import timeit

# Import json for storing the results
import json

# Import needed libraries for acquiring the case files
from os import listdir
from os.path import isfile, join

# Import the win32com.client to create a SimAuto object
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")


def main_func():

	# Add the path to your cases below.
	# Example: "C:\\Dir1\\Dir2\\Dir3\\case_dir"
	case_folder = ""

	# Create a list of all the PWB case files
	case_files = [file for file in listdir(case_folder) if (isfile(join(case_folder, file)) and (file.endswith('pwb') or file.endswith('PWB')))]

	# Paste the function that you wish to be timed below.
	def func():
		# Paste here


	# Create a dictionary of the fastest time for every case
	case_times = {}


	# p is the number of times you will time the function.
	# The higher p is, the more precision, but the longer the run time.
	p = 100

	for case in case_files:

		# Open the current case
		SimAuto.OpenCase(case_folder + '\\' + case)

		# Create a list of timings for this case
		timing_list = []

		# Time the function (for p number of times) and add the timings
		# to the case's list of timings
		for i in range(p):
			timing = timeit.timeit(func, number=1)
			timing_list.append(timing)

		# Add the fastest time to the dictionary of all cases
		case_times[case] = min(timing_list)

	# Open a file (or create one if none exists) and dump the fastest
	# times for all the cases in.
	with open('timing_results.txt', 'w') as outfile:
		json.dump(case_times, outfile)

main_func()