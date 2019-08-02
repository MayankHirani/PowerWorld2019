"""
Example of how to use SimAuto methods in Python
Developed by PowerWorld 2019, created by Mayank Hirani

Script #4: Other SimAuto functions to retrieve different data
Includes: Setup, OpenCaseType(), GetCaseHeader(), ListOfDevices(),
ListOfDevicesAsVariantStrings(), GetSpecificFieldList(), GetSystemMetrics(),
CloseCase()
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
file_name = "c:\\Users\\mayank\\Desktop\\mayank\\B7FLAT Testing\\B7FLAT.pwb"

# Open the case using the file path and specify the case type
CheckResultForError(SimAuto.OpenCaseType(file_name, "PWB"), "Opened Case")

# Print the case header, which includes the description of the case, using
# the GetCaseHeader() method
print("Case Header:", SimAuto.GetCaseHeader(file_name)[1])

# Lists all buses using the ListOfDevices() method
print("\nDevices:", SimAuto.ListOfDevices("Bus", '')[1][0])
print("Devices as strings:", SimAuto.ListOfDevicesAsVariantStrings("Bus", '')[1][0], "\n")

# Lists the specific fields for GenMW using the GetSpecificFieldList() method
print("Field List:", SimAuto.GetSpecificFieldList("Bus", [ "GenMW" ])[1][0], "\n")

# Displays the system metrics using the GetSystemMetrics() method
print("System Metrics:", SimAuto.GetSystemMetrics()[1])

# Close the B7FLAT case
CheckResultForError(SimAuto.CloseCase(), "Closed Case")
