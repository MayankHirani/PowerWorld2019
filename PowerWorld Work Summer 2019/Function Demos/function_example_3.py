"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Script #3: 2 Methods to load the data for buses
Includes: Setup, OpenCase(), WriteAuxFile(), GetParametersMultipleElementRect(),
ChangeParametersMultipleElementsRect(), CloseCase()
"""

# Import the win32com library to connect to SimAuto
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Function used to check if SimAuto commands are working properly
def CheckResultForError(SimAutoOutput, Message):
    if SimAutoOutput[0] != '':
        print('Error: ' + SimAutoOutput[0])
    else:
        print(Message)

# Example of file path:
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing"

# The fields to be used
parameters = [ "BusNum", "AreaNum", "BusGenMVR" ]

# Open the case using the file path
CheckResultForError(SimAuto.OpenCase(file_name + "\\B7FLAT.pwb"), "Opened Case")

# METHOD 1: Write an AUX file with the data

SimAuto.RunScriptCommand('DeleteFile("%s\\B7FLAT.pwb");' % file_name)

# Create an AUX file with bus data
CheckResultForError(SimAuto.WriteAuxFile("bus_data.aux", 0, "Bus", False, parameters), "Created AUX file")

# METHOD 2: Get the parameters using GetParametersMultipleElementRect(), 
# which returns the values for each bus as a tuple of tuples

# Print the values for the parameters from the output of 
# GetParametersMultipleElementRect()
print(SimAuto.GetParametersMultipleElementRect("Bus", parameters, 0))
print()
values = SimAuto.GetParametersMultipleElementRect("Bus", parameters, 0)[1]
for value in values:
	print(value[0].strip(), value[1].strip(), value[2].strip())
print()

# CHANGING VALUES: Values have to be loaded from GetParameters or by using
# VARIANT. For VARIANT, see function_exmaple_2.
# Change the values using ChangeParametersMultipleElementRect()
CheckResultForError(SimAuto.ChangeParametersMultipleElementRect("Bus", parameters, values), "Changed Values")

# Close the B7FLAT case
CheckResultForError(SimAuto.CloseCase(), "Closed Case")