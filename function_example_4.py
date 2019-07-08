
# Example of how to use SimAuto methods in Python
# Developed by PowerWorld 2019, created by Mayank Hirani

# Includes: Setup, OpenCaseType(), UIVisible, ListOfDevices(),
# GetSpecificFieldList(), CloseCase()

# Import the win32com library to connect to SimAuto
import win32com.client
from win32com.client import VARIANT
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Import pythoncom to use VARIANT and os.path to interact with files
# and directories
import pythoncom
import os.path

# Function used to check if SimAuto commands are working properly
def CheckResultForError(SimAutoOutput, Message):
    if SimAutoOutput[0] != '':
        print('Error: ' + SimAutoOutput[0])
    else:
        print(Message)

# If we want to see a visual representation, we use UIVisible. In this
# example, the B7FLAT case will be opened in PowerWorld Simulator
SimAuto.UIVisible = True

# Example of file path:
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT_mayank.pwb"

# Open the case using the file path and specify the case type
CheckResultForError(SimAuto.OpenCaseType(file_name, "PWB"), "Opened Case")

# Lists all buses using the ListOfDevices() method
print(SimAuto.ListOfDevices("Bus", '')[1][0])

# Lists the specific fields for GenMW using the GetSpecificFieldList() method
print(SimAuto.GetSpecificFieldList("Bus", [ "GenMW" ])[1][0])

# Lists the system metrics
print("System Metrics:", SimAuto.GetSystemMetrics()[1])

# Print the case header, which includes the description of the case
print("Case Header:", SimAuto.GetCaseHeader(file_name)[1])

# Print the proccess ID of the current process
print("Process ID:", SimAuto.ProcessID)

# Close the B7FLAT case
CheckResultForError(SimAuto.CloseCase(), "Closed Case")
