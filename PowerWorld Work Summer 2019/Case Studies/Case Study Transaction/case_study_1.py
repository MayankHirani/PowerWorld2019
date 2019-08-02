"""
Case Study done in Python to interact with SimAuto
Developed by PowerWorld 2019, created by Mayank Hirani
Description: Solves a contingency test before and after a PTDF transaction
Based off of: https://www.powerworld.com/knowledge-base/ptdf-and-contingency-analysis-of-a-new-transaction
"""

# Import the win32com.client to create a SimAuto object
import win32com.client
SimAuto = win32com.client.Dispatch("pwrworld.SimulatorAuto")

# Specify the file path to where B7OPF.pwb is
# Example file path: 'c:\\Users\\User1\\Desktop\\Directory\\B7OPF.pwb'
file_path = 'c:\\...\\B7OPF.pwb'

# Specify the directory that the saved AUX files are (and this file)
# Example file path: 'c:\\Users\\User1\\Desktop\\Directory\\case_study_directory'
save_file_path = 'c:\\'

# Open the case file
SimAuto.OpenCase(file_path)


# Solve the powerflow
SimAuto.RunScriptCommand('SolvePowerFlow;')

# Set the system as a base case for different flows
SimAuto.RunScriptCommand('DiffFlowSetAsBase;')

# Calculate the PTDF values between the seller and the buyer (in this case,
# AREA 2 is the seller, and AREA 1 is the buyer)
SimAuto.RunScriptCommand('CalculatePTDF([AREA 2], [AREA 1], AC);')

# Delete any currently saved data
SimAuto.RunScriptCommand('DeleteFile("%s\\study_1_ptdf_data");' % save_file_path)

# Save the PTDF data to a designated location (the same file that was
# cleared above)
SimAuto.RunScriptCommand('SaveData("%s\\study_1_ptdf_data", AUX, BRANCH, [BusNum, BusNum:1, AbsValPTDF],[]);' % (save_file_path))


# Load the contingencies in from the AUX file they are stored in
SimAuto.RunScriptCommand('LoadAux("%s\\case_study_1_contingencies.aux", YES);' % save_file_path)

# Solve the contingencies
SimAuto.RunScriptCommand('CTGSolve(Contingencies);')

# Create a report on the solved contingencies
SimAuto.RunScriptCommand('CTGProduceReport("%s\\BeforeTransactionCTGReport");' % save_file_path)


# Load in the MW transaction between Area 1 and Area 2
SimAuto.RunScriptCommand('LoadAux("%s\\case_study_1_transaction.aux", YES);' % save_file_path)


# Solve the power flow
SimAuto.RunScriptCommand('SolvePowerFlow;')

# Set the system as a base case for different flows
SimAuto.RunScriptCommand('DiffFlowMode(DIFFERENCE);')

# Calculate the PTDF values between the seller and the buyer once again
SimAuto.RunScriptCommand('CalculatePTDF([Area 2], [Area 1], AC);')

SimAuto.RunScriptCommand('SaveData("%s\\study_1_ptdf_data", AUX, BRANCH, [BusNum, BusNum:1, AbsValPTDF],[]);' % save_file_path)


# Solve the contingencies again
SimAuto.RunScriptCommand('CTGSolve(Contingencies);')

# Save the results to a new AUX file
SimAuto.RunScriptCommand('CTGProduceReport("%s\\AfterTransactionCTGReport");' % save_file_path)


# Close the case file
SimAuto.CloseCase()
