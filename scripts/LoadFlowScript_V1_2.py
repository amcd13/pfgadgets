import powerfactory as pf
import os
import pandas as pd
from pfgadgets import LoadFlow, GetData

"""
Version: 1.2

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script executes a load flow calculation and exports the
             results to file containing csv

Inputs: terminalSetName - Name of the set containing calculation relevant terminal/s
        lineSetName - Name of the set containing calculation relevant line/s
        transformerSetName - Name of the set containing calculation relevant transformer/s
        terminalAttributes - List of variable/s to export for terminal object/s
        lineAttributes - List of variable/s to export for line object/s
        transformerAttributes - List of variables/s to export for transformer object/s
        terminalHeadings - List of string/s to replace terminal dataframe headings
        lineHeadings - List of string/s to replace line dataframe headings
        transformerHeadings - List of string/s to replace transformer dataframe headings
"""

# Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

######### INPUTS ############################################

# Define set names
terminalSetName = 'Terminals'
lineSetName = 'Lines'
transformerSetName = 'Transformers'

# Define object attributes to collect from load flow results
terminalAttributes = ['e:uknom', 'm:u', 'm:Ul', 'Ir']
lineAttributes = ['Inom', 'c:loading']
transformerAttributes = ['Snom', 'c:loading', 'n:u:bushv', 'n:u:buslv', 'c:nntap']

# Define headings for results
terminalHeadings = ['Name', 'Rated Voltage (kV)', 'Voltage (pu)', 'Voltage (kV)', 'Rated Current (kA)']
lineHeadings = ['Name', 'Rated Current (kA)', 'Loading (%)']
transformerHeadings = ['Name', 'Rated Power (MVA)', 'Loading (%)', 'HV Voltage (kV)', 'LV Voltage (kV)', 'Tap Position']

# Define operation scenarios to perform load flow analysis on
operationScenarios = ['Maximum Normal Load', 'Peak Load']

#############################################################

# Initialise dataframes to store results
allTerminalResults = pd.DataFrame()
allLineResults = pd.DataFrame()
allTransformerResults = pd.DataFrame()

# Get script path and make results directory
scriptPath = os.getcwd()
resultsPath = scriptPath + '\\LoadFlowResults\\'

if not os.path.exists(resultsPath):
    os.mkdir(resultsPath)

# Iterate trough operation scenarios
for operationScenario in operationScenarios:
    # Execute LoadFlow command
    LoadFlow(method=0, autoTap=0, feederScaling=0, opScen=operationScenario)

    # Collect objects data
    terminalDataFrame = GetData(terminalSetName, terminalAttributes, resultHeadings=terminalHeadings).result
    lineDataFrame = GetData(lineSetName, lineAttributes, resultHeadings=lineHeadings).result
    transformerDataFrame = GetData(transformerSetName, transformerAttributes, resultHeadings=transformerHeadings).result

    # Insert column with operation scenario name
    terminalDataFrame.insert(0, 'Scenario', [operationScenario] * terminalDataFrame.shape[0])
    lineDataFrame.insert(0, 'Scenario', [operationScenario] * lineDataFrame.shape[0])
    transformerDataFrame.insert(0, 'Scenario', [operationScenario] * transformerDataFrame.shape[0])

    # Append scenario results to master data frames
    allTerminalResults = pd.concat([allTerminalResults, terminalDataFrame])
    allLineResults = pd.concat([allLineResults, lineDataFrame])
    allTransformerResults = pd.concat([allTransformerResults, transformerDataFrame])

# Write results to xlsx workbook in separate sheets
with pd.ExcelWriter(resultsPath + 'LoadFlowResults.xlsx') as writer:
    allTerminalResults.to_excel(writer, sheet_name='Terminal', index=False)
    allLineResults.to_excel(writer, sheet_name='Line', index=False)
    allTransformerResults.to_excel(writer, sheet_name='Transformer', index=False)
