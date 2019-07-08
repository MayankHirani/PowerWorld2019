
# Example of how to use SimAuto methods in Python
# Developed by PowerWorld 2019, created by Mayank Hirani

# Includes: Setup, OpenCase(), GetParametersMultipleElement(), VARIANT,
# CreateIfNotFound, ChangeParametersMultipleElement(), WriteAuxFile(),
# CloseCase()

# Import the win32com library to connect to SimAuto
import win32com.client
from win32com.client import VARIANT
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Import pythoncom to use VARIANT and os.path to interact with files
# and directories
import pythoncom
import os.path

# This function is used to check if SimAuto commands are working properly
def CheckResultForError(SimAutoOutput, Message):
    if SimAutoOutput[0] != '':
        print('Error: ' + SimAutoOutput[0])
    else:
        print(Message)

# Example of file path:
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT_mayank.pwb"

# Open the case using the file path
CheckResultForError(SimAuto.OpenCase(file_name), "Opened Case")

# The fields that will be used when retrieving the parameters and values, 
# as well as changing values
parameters = [ "GenID", "BusNum", "GenMW" ]

# Retrieve the numbers of the buses and their corresponding GenMW
bus_numbers = [x.strip() for x in (SimAuto.GetParametersMultipleElement("Gen", parameters, 0)[1][1])]
gen_mw = [x.strip() for x in (SimAuto.GetParametersMultipleElement("Gen", parameters, 0)[1][2])]

# Print each bus number and its corresponding GenMW from the two lists
print('\nGenMW of each bus')
for x in range(len(bus_numbers)):
	print(bus_numbers[x], gen_mw[x])
print()

# VARIANT is a data type used for data processes, must be used here for the
# arrays of arrays.
# To change values, VARIANT must be used or values obtained from
# GetParameter(). For the second method, see function_example_3.
parameters = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_VARIANT, parameters)

# Make list of empty values that the new values will be inserted into
# List will look like: [ None, None, None... ]
values = [ None ] * len(bus_numbers)

# Create a list of the GenMW values as floats, and calculate the mean of that
# list
gen_mw_float = [ float(x) for x in gen_mw ]
average_gen_mw = sum(gen_mw_float) / len(gen_mw_float)

# Each value in the set of values is set to the list of values where  GenMW
# is the average
for x, bus_num in enumerate(bus_numbers):
	values[x] = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_VARIANT, ["1", bus_num, average_gen_mw])

# CreateIfNotFound will create a value if there is no value for a given field,
# this is a data attribute of the SimAuto object
SimAuto.CreateIfNotFound = True

# Use the values array to change all the GenMW of all the buses to 100 using
# the set of all the values
CheckResultForError(SimAuto.ChangeParametersMultipleElement("Gen", parameters, values), "Changed GenMW Values")

# Check that the values were changed by printing the bus numbers and their
# corresponding GenMW. Create a list of the bus numbers, and a list of the
# GenMW values
bus_numbers = [x.strip() for x in (SimAuto.GetParametersMultipleElement("Gen", parameters, 0)[1][1])]
gen_mw = [x.strip() for x in (SimAuto.GetParametersMultipleElement("Gen", parameters, 0)[1][2])]

# Print out each bus number and its corresponding GenMW
print('\nGenMW of each bus')
for x in range(len(bus_numbers)):
	print(bus_numbers[x], gen_mw[x])
print()

# Clear the AUX file prior to appending data to it, if the file exists
if os.path.isfile('c://Users/mayank/Desktop/mayank/B7FLAT Testing/written_aux.aux'):
	open('c://Users/mayank/Desktop/mayank/B7FLAT Testing/written_aux.aux', 'w').close()

# Create an AUX file with the generator data
CheckResultForError(SimAuto.WriteAuxFile("written_aux.aux", 0, "Gen", True, parameters), "Created AUX file")

# Close the B7FLAT case
CheckResultForError(SimAuto.CloseCase(),"Closed Case")