"""
Case Study done in Python to interact with SimAuto
Developed by PowerWorld 2019, created by Mayank Hirani
Description: Solves tbe contingency before and after scaling the load 
Based off of: https://www.powerworld.com/knowledge-base/comparing-contingency-results 
"""

# Import the win32com.client to create a SimAuto object
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Specify the file path to where MultCTG_B7FLAT.pwb is stored
# Example file path: 'c:\\Users\\User1\\Desktop\\Directory\\MultCTG_B7FLAT.pwb'
file_path = 'c:\\...\\MultCTG_B7FLAT.pwb'

# Specify the directory that AUX files are stored in (as well as this module)
# Example file path: 'c:\\Users\\User1\\Desktop\\Directory\\case_study_directory'
save_file_path = 'c:\\'

# Open the case file
SimAuto.OpenCase(file_path)

# Load the contingencies that will be solved
SimAuto.RunScriptCommand('LoadAux("%s\\case_study_2_contingencies.aux", YES);' % save_file_path)

# Enter run mode
SimAuto.RunScriptCommand('EnterMode(RUN);')

# Flag all buses for scaling
SimAuto.RunScriptCommand('SetData(Bus, [BusScale], ["Yes"], All);')

# The list of fields that will be included when data is saved for the contingencies
field_list = '[CTGLabel,OwnerName,OwnerName:1,CustomString:2,BusNum,BusNum:1,BusName,BusName:1,BusNomVolt:1,BusNomVolt:2,LineCircuit,OwnerName:2,OwnerName:3,LimViolCat,LimViolLimit,LimViolPct,LimViolValue:2,LimViolValue,LimViolPct:1]'

# The scales that the system load will be set to each time.
# First is 80% of base load, second is 100% of the load, third is 120% of the load
scale_values = { 'low_load':0.80, 'base_load':1.00, 'high_load':1.20 }

# This is the function we will pass each scaling value into.
# The function changes the load and solves contingencies and saves the data
def load_scaling(load_type):

	# Use the scale value that corresponds with the load change from the dictionary
	scale = scale_values[load_type]

	# Scale the system load to the specified amount
	SimAuto.RunScriptCommand('Scale(Load, Factor, [%s], Bus);' % scale)

	# Solve the power flow
	SimAuto.RunScriptCommand('SolvePowerFlow;')

	# Set as contingency reference
	SimAuto.RunScriptCommand('CTGSetAsReference;')

	# Solve the contingencies
	SimAuto.RunScriptCommand('CTGSolveAll;')

	# Save the data of the contingencies with the fields specified in field_list
	SimAuto.RunScriptCommand('SaveData("%s\\%s_ctg_results", AUX, ViolationCTG, %s, []);' % (save_file_path, load_type, field_list))

	# Reset the load back to base value (The scale function multiples the
	# current value by the scale value specified)
	scale = 1 / scale
	SimAuto.RunScriptCommand('Scale(Load, Factor, [%s], Bus);' % scale)


# Iterate through each scale factor and call the function on each one
for load_type in scale_values:
	load_scaling(load_type)