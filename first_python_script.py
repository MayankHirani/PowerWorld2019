
# Playing around with the commands Thomas gave, calling them directly in Python
# Getting values from a case, changing them, and calling minor functions on the case, then saving the current case as 'B7FLAT_mayank_updated.pwb'

# Import all needed libraries and add colors for fun
import sys
import importlib
import datetime
import os
CEND = '\033[0m'
YELLOW = '\033[33m'
GRAY = '\033[90m'
RED = '\033[91m'

# Get the reload function and connection to SimAuto
from importlib import reload
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Method from Thomas to print actions and errors
def CheckResultForError(SimAutoOutput, Message):
    if SimAutoOutput[0] != '':
        print('Error: ' + SimAutoOutput[0])
    else:
        print(GRAY + Message + CEND)

# Create a date and time and add it to the log
timestamp = str(datetime.datetime.now())
print('\n' + YELLOW + timestamp + CEND)
print()

# The case to be used is B7FLAT must be located here
case_file = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT_mayank.pwb"

# Open the case, clear the log, add the timestamp, and solve the powerflow (RunScriptCommand will run AUX commands)
CheckResultForError(SimAuto.OpenCase(case_file), "Opened Case")
CheckResultForError(SimAuto.RunScriptCommand("LogClear;"), "Cleared Log")
CheckResultForError(SimAuto.RunScriptCommand('LogAdd(%s)' % timestamp), "Added Timestamp")
CheckResultForError(SimAuto.RunScriptCommand("SolvePowerFlow(RECTNEWT);"), "Solved Power Flow")

# Contingency setup, run, and save
CheckResultForError(SimAuto.RunScriptCommand('LoadAux("c://Users//mayank//Desktop//mayank//B7FLAT Testing//autoinsert_options.aux");'), "Loaded Options")
CheckResultForError(SimAuto.RunScriptCommand('CTGAutoInsert;'), "Auto Inserted Options")
CheckResultForError(SimAuto.RunScriptCommand('CTGSolveAll;'), "Contigency Test Run")
CheckResultForError(SimAuto.RunScriptCommand('CTGProduceReport("first_script_contingency_report");'), "Created Contingency Report")

# Create a list of fields, enter  values for required information
parameters = [ 'BusName', 'BusNum', 'BusAngle' ]
bus_names = [ 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven' ]
bus_numbers = [ '1', '2', '3', '4', '5', '6', '7' ]
bus_angles = []

# We do not want to create a new bus if it cannot find one
SimAuto.CreateIfNotFound = False

# |GET FUNCTION| Find the angle for each bus using the GetParameters() function
def get_angles(parameters, bus_names, bus_numbers, bus_angles):

	values = [ '', '', '' ]
	print()
	print(YELLOW + 'Bus Angles' + CEND)
	for x in range(7):

		# Set the values of the required fields (Name and number) so it knows what bus you are referring to
		values[0] = bus_names[x]
		values[1] = bus_numbers[x]

		# Print out the angles
		angle = SimAuto.GetParameters("Bus", parameters, values)[1][2]
		print("Bus " + bus_names[x] + " angle: " + angle)

		# Remove the white space before the number
		if '-' in angle:
			bus_angles.append(angle[1:])
		else:
			bus_angles.append(angle[2:])

	# Print out all the angles in a list		
	print()
	print(bus_angles)
	print()

	return bus_angles


# |CHANGE SINGLE| Trying to change the angle of just the first bus to 5
def change_first_bus(parameters, values):
	
	# Reset values to just the first bus specified, as well as an angle
	values = [ 'One', '1', '5' ]

	# Change the values for the first bus
	SimAuto.ChangeParametersSingleElement("Bus", parameters, values)
	print(SimAuto.GetParameters("Bus", parameters, values))


# |CHANGE ALL| Goes through all the buses and changes each angle
def change_indv(parameters, bus_names, bus_numbers, bus_angles):

	# Reset the values
	values = [ '', '', '' ]
	print(YELLOW + 'New Bus Angles\t' + CEND + GRAY + '(angle * -1)' + CEND)

	for x in range(7):

		# For each one, the new set of values is [ bus_name, bus_number, inverse of bus angle ]
		values = [ bus_names[x], bus_numbers[x], str(float(bus_angles[x]) * -1) ]
		SimAuto.ChangeParametersSingleElement("Bus", parameters, values)
		
		# Print these angles now
		angle = SimAuto.GetParametersSingleElement("Bus", parameters, values)[1][2]
		print("Bus " + bus_names[x] + " angle: " + angle)

		# Remove white space before
		if '-' in angle:
			bus_angles[x] = angle[1:]
		else:
			bus_angles[x] = angle[2:]

	# Print out all the angles in a list
	print()
	print(bus_angles)
	print()

# |CHANGE ALL| Changes all the values in one function with MultipleElement
def change_all(parameters, bus_names, bus_numbers, bus_angles):

	values = [ bus_names, bus_numbers, bus_angles ]
	CheckResultForError(SimAuto.ChangeParametersMultipleElement("Bus", parameters, VALUES))

# Get the angles of all the buses, then change each one by calling the methods
print(GRAY + '\n-------------------------------------------' + CEND)
bus_angles = get_angles(parameters, bus_names, bus_numbers, bus_angles)
print(GRAY + '-------------------------------------------\n' + CEND)
change_indv(parameters, bus_names, bus_numbers, bus_angles)
print(GRAY + '-------------------------------------------\n' + CEND)

# Save and close the case after saving the log
CheckResultForError(SimAuto.RunScriptCommand('LogSave("first_python_script_log", NO);'), "Log Saved")
CheckResultForError(SimAuto.SaveCase(r"c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT_mayank_updated.pwb", "PWB", True), "Saved Case")
CheckResultForError(SimAuto.CloseCase(), "Closed Case")

# Ask the user if they want to keep the created files or delete them
user_input = ''
print(RED + "\nWould you like to save CONTINGENCY REPORT & LOG? ( Y > Yes | N > No )" + CEND)
while user_input not in ['y', 'Y', 'n', 'N']:
	print(">>>  ", end='')
	user_input = input()

# If yes, delete both the contingency report and the log
if user_input in ['n', 'N']:
	os.remove('c://Users//mayank//Desktop//mayank//B7FLAT Testing//first_python_script_log')
	os.remove('c://Users//mayank//Desktop//mayank//B7FLAT Testing//first_script_contingency_report')
	print(GRAY + '\nFiles Deleted\n' + CEND)
else:
	print(GRAY + '\nFiles Saved\n' + CEND)

