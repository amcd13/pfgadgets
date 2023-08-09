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

#############################################################
# INPUTS

# Define set names
terminal_set_name = 'Terminals'
line_set_name = 'Lines'
transformer_set_name = 'Transformers'

# Define object attributes to collect from load flow results
terminal_attributes = ['e:uknom', 'm:u', 'm:Ul', 'Ir']
line_attributes = ['Inom', 'c:loading']
transformer_attributes = ['Snom', 'c:loading', 'n:u:bushv', 'n:u:buslv', 'c:nntap']

# Define headings for results
terminal_headings = ['Name', 'Rated Voltage (kV)', 'Voltage (pu)', 'Voltage (kV)', 'Rated Current (kA)']
line_headings = ['Name', 'Rated Current (kA)', 'Loading (%)']
transformer_headings = ['Name', 'Rated Power (MVA)', 'Loading (%)',
                        'HV Voltage (kV)', 'LV Voltage (kV)', 'Tap Position']

# Define operation scenarios to perform load flow analysis on
operation_scenarios = ['Maximum Normal Load', 'Peak Load']

#############################################################

# Initialise dataframes to store results
all_terminal_results = pd.DataFrame()
all_line_results = pd.DataFrame()
all_transformer_results = pd.DataFrame()

# Get script path and make results directory
script_path = os.getcwd()
results_path = script_path + '\\LoadFlowResults\\'

if not os.path.exists(results_path):
    os.mkdir(results_path)

# Iterate trough operation scenarios
for operation_scenario in operation_scenarios:
    # Execute LoadFlow command
    LoadFlow(method=0, auto_tap=0, feeder_scaling=0, op_scen=operation_scenario)

    # Collect objects data
    terminal_df = GetData(terminal_set_name, terminal_attributes, result_headings=terminal_headings).result
    line_df = GetData(line_set_name, line_attributes, result_headings=line_headings).result
    transformer_df = GetData(transformer_set_name, transformer_attributes, result_headings=transformer_headings).result

    # Insert column with operation scenario name
    terminal_df.insert(0, 'Scenario', [operation_scenario] * terminal_df.shape[0])
    line_df.insert(0, 'Scenario', [operation_scenario] * line_df.shape[0])
    transformer_df.insert(0, 'Scenario', [operation_scenario] * transformer_df.shape[0])

    # Append scenario results to master data frames
    all_terminal_results = pd.concat([all_terminal_results, terminal_df])
    all_line_results = pd.concat([all_line_results, line_df])
    all_transformer_results = pd.concat([all_transformer_results, transformer_df])

# Write results to xlsx workbook in separate sheets
with pd.ExcelWriter(results_path + 'LoadFlowResults.xlsx') as writer:
    all_terminal_results.to_excel(writer, sheet_name='Terminal', index=False)
    all_line_results.to_excel(writer, sheet_name='Line', index=False)
    all_transformer_results.to_excel(writer, sheet_name='Transformer', index=False)
    app.PrintPlain('Load flow results written to: %s' % (results_path + 'LoadFlowResults.xlsx'))
