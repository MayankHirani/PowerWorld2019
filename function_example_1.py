
# Example of how to use SimAuto methods in Python
# Developed by PowerWorld 2019, created by Mayank Hirani

# Includes: Setup, CurrentDir, OpenCase(), RunScriptCommand(), CloseCase()

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

# Below clears the log with the AUX file command 'LogClear;'
CheckResultForError(SimAuto.RunScriptCommand('LogClear;'), "Cleared Log")

# Below completes a Powerflow test with the AUX file command
# 'SolvePowerFlow();'
CheckResultForError(SimAuto.RunScriptCommand('SolvePowerFlow(RECTNEWT)'), "Solved Power Flow")

# Function that closes the B7FLAT case and displays "Closed Case" if
# successful, otherwise displays SimAuto error
CheckResultForError(SimAuto.CloseCase(), "Closed Case")


# A full list of the script commands that can be run using RunScriptCommand()
# can be found here: 
# https://www.powerworld.com/WebHelp/Content/Other_Documents/Auxiliary-File-Format.pdf