"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Script #5: Different SimAuto data attributes
Includes: Setup, OpenCase(), UIVisible, CurrentDir, CreateIfNotFound, ProcessID
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

# Example of file path for the B7FLAT case:
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT_mayank.pwb"

# Function that opens the case and displays "Opened Case" if successful, 
# otherwise displays SimAuto error
CheckResultForError(SimAuto.OpenCase(file_name), "Opened Case")

# Use the CurrentDir attribute to print the current directory
print(SimAuto.CurrentDir)

# CreateIfNotFound will create a value if there is no value for a given field,
# this is a data attribute of the SimAuto object
SimAuto.CreateIfNotFound = True

# Print the proccess ID of the current process with the ProcessID attributes
print("Process ID:", SimAuto.ProcessID)