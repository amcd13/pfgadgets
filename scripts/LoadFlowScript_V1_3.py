import powerfactory as pf
import os
import pandas as pd
from pfgadgets import LoadFlow, GetData
from DataProcessing_V1_0 import busbar_util, busbar_voltage, cable_util, cb_util, tx_util

"""
Version: 1.3

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script executes a load flow calculation and exports the
             results to excel workbook

Inputs: term_set_name - Name of the set containing calculation relevant terminal/s
        line_set_name - Name of the set containing calculation relevant line/s
        tx_set_name - Name of the set containing calculation relevant transformer/s
        cb_set_name - Name of the set containing calculation relevant circuit breaker/s
        term_dict - Dictionary with list of attribute/s & headings to export for terminal object/s
        line_dict - Dictionary with list of attribute/s & headings to export for line object/s
        tx_dict - Dictionary with list of attribute/s & headings to export for transformer object/s
        cb_dict - Dictionary with list of attribute/s & headings to export for circuit breaker object/s
"""

# Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

#############################################################
# INPUTS

# Define set names
term_set_name = 'Terminals'
line_set_name = 'Lines'
tx_set_name = 'Transformers'
cb_set_name = 'Circuit Breakers'

# Define object attributes to collect from load flow results
term_dict = {'attributes': ['loc_name',
                            'e:uknom',
                            'm:u',
                            'm:Ul',
                            'e:vmax',
                            'e:vmin',
                            'm:Sout',
                            'Ir',
                            'Ithlim',
                            'Tkr',
                            'Iplim'],
               'headings': ['Name',
                            'Rated Voltage (kV)',
                            'Voltage (pu)',
                            'Voltage (kV)',
                            'Upper Voltage Limit (pu)',
                            'Lower Voltage Limit (pu)',
                            'Outgoing Apparent Power (kVA)',
                            'Rated Current (kA)',
                            'Rated Short-Time Thermal Current (kA)',
                            'Rated Short-Circuit Duration (s)',
                            'Short-Circuit Peak Current (kA)']}
line_dict = {'attributes': ['loc_name',
                            'Inom',
                            'c:loading',
                            'e:dline',
                            'm:I:bus2'],
               'headings': ['Name',
                            'Rated Current (kA)',
                            'Loading (%)',
                            'Line Length (m)',
                            'Maximum Current (kA)']}
tx_dict = {'attributes': ['loc_name',
                          't:utrn_h',
                          't:utrn_l',
                          'Snom',
                          'm:I:buslv',
                          'c:loading',
                          'n:u:bushv',
                          'n:u:buslv',
                          'c:nntap'],
             'headings': ['Name',
                          'Rated Primary Voltage (kV)',
                          'Rated Secondary Voltage (kV)',
                          'Rated Power (MVA)',
                          'Current Secondary (kA)',
                          'Loading (%)',
                          'HV Voltage (kV)',
                          'LV Voltage (kV)',
                          'Tap Position']}
cb_dict = {'attributes': ['loc_name',
                          'Inom',
                          'm:I:bus1',
                          'm:I:bus2',
                          'm:Brkload:bus1',
                          'm:Brkload:bus2'],
             'headings': ['Name',
                          'Rated Current (kA)',
                          'Current Bus1 (kA)',
                          'Current Bus2 (kA)',
                          'Circuit Breaker Loading Bus 1 (%)',
                          'Circuit Breaker Loading Bus 2 (%)']}

# Define operation scenarios to perform load flow analysis on
operation_scenarios = ['Maximum Normal Load', 'Peak Load']

#############################################################

# Initialise dataframes to store results
all_term_results = pd.DataFrame()
all_line_results = pd.DataFrame()
all_tx_results = pd.DataFrame()
all_cb_results = pd.DataFrame()

# Get script path and make results directory
script_path = os.getcwd()
results_path = script_path + '\\LoadFlowResults\\'

if not os.path.exists(results_path):
    os.mkdir(results_path)

# Iterate trough operation scenarios
for operation_scenario in operation_scenarios:
    # Execute LoadFlow command
    LoadFlow(method=0, auto_tap=0, feeder_scaling=0).ex(op_scen=operation_scenario)

    # Collect objects data
    term_df = GetData(term_set_name, term_dict['attributes']).result
    line_df = GetData(line_set_name, line_dict['attributes']).result
    tx_df = GetData(tx_set_name, tx_dict['attributes']).result
    cb_df = GetData(cb_set_name, cb_dict['attributes']).result

    # Insert column with operation scenario name
    term_df.insert(0, 'Scenario', [operation_scenario] * term_df.shape[0])
    line_df.insert(0, 'Scenario', [operation_scenario] * line_df.shape[0])
    tx_df.insert(0, 'Scenario', [operation_scenario] * tx_df.shape[0])
    cb_df.insert(0, 'Scenario', [operation_scenario] * cb_df.shape[0])

    # Append scenario results to master data frames
    all_term_results = pd.concat([all_term_results, term_df])
    all_line_results = pd.concat([all_line_results, line_df])
    all_tx_results = pd.concat([all_tx_results, tx_df])
    all_cb_results = pd.concat([all_cb_results, cb_df])

# Processing Data into report tables
busbar_utilisation = busbar_util(all_term_results)
busbar_voltage_level = busbar_voltage(all_term_results)
cable_utilisation = cable_util(all_line_results)
tx_utilisation = tx_util(all_tx_results)
cb_utilisation = cb_util(all_cb_results)

# Write results to xlsx workbook in separate sheets
with pd.ExcelWriter(results_path + 'LoadFlowResults.xlsx') as writer:
    all_term_results.to_excel(writer, sheet_name='Terminal', index=False)
    all_line_results.to_excel(writer, sheet_name='Line', index=False)
    all_tx_results.to_excel(writer, sheet_name='Transformer', index=False)
    all_cb_results.to_excel(writer, sheet_name='Circuit Breaker', index=False)

    busbar_utilisation.to_excel(writer, sheet_name='BusbarUtilisation', index=False)
    busbar_voltage_level.to_excel(writer, sheet_name='BusbarVoltageLevel', index=False)
    cable_utilisation.to_excel(writer, sheet_name='CableUtilisation', index=False)
    cb_utilisation.to_excel(writer, sheet_name='CircuitBreakerUtilisation', index=False)
    tx_utilisation.to_excel(writer, sheet_name='TransformerUtilisation', index=False)

    app.PrintPlain('Load flow results written to: %s' % (results_path + 'LoadFlowResults.xlsx'))
