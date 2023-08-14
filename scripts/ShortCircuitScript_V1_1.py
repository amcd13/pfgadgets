import powerfactory as pf
import os
import pandas as pd
from pfgadgets import ShortCircuit, GetData
from DataProcessing_V1_0 import fault_level_summary

"""
Version: 1.1

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script executes a short-circuit calculation and exports
             the results to csv files

Inputs: fault_types - The types of faults to run
        maximum - The name of the maximum fault operation scenario
        minimum - The name of the minimum fault operation scenario
        sc_set_name - The name of the set containing the objects to run the short-circuit calculation on
        sc3p_attributes - A list of variables to extract during 3-phase faults
        scgf_attributes - A list of variables to extract during ground faults
"""

# Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

#############################################################
# INPUTS

# Define fault types to run
fault_types = ['3psc', 'spgf']
fault_cases = ['Max', 'Min']

# Define maximum and minimum operation scenarios
maximum = ['Maximum Normal Load']
minimum = ['Peak Load']

# Setup set names
sc_set_name = 'Short-Circuit Set'

# Setup object attributes to collect from load flow results
sc3p_attributes = [
    'loc_name',
    'e:uknom',
    'm:Ikss',
    'm:Skss',
    'm:Ith']
scgf_attributes = [
    'loc_name',
    'e:uknom',
    'm:Ikss:A',
    'm:Skss:A',
    'e:uknom']

#############################################################

# Initialise dataframes to store results
max_3psc_data = pd.DataFrame()
min_3psc_data = pd.DataFrame()
max_spgf_data = pd.DataFrame()
min_spgf_data = pd.DataFrame()

# Define file path for export
file_path = os.getcwd() + '\\ShortCircuitResults\\'

if not os.path.exists(file_path):
    os.mkdir(file_path)

operation_scenarios = maximum + minimum

# Iterate through operation scenarios
for operation_scenario in operation_scenarios:
    # Iterate through fault types
    for fault_type in fault_types:
        #Iterate through fault cases:
        for fault_case in fault_cases:
            #Change sc command inputs based on active operation scenario
            if fault_case == 'Max':
                calculate = 0
            elif fault_case == 'Min':
                calculate = 1

            if fault_type == '3psc':
                fault_attributes = sc3p_attributes
            elif fault_type == 'spgf':
                fault_attributes = scgf_attributes

            # Execute ShortCircuit maximum command
            ShortCircuit(sc_set_name, fault_type=fault_type, calculate=calculate, set_select=True, op_scen=operation_scenario)

            # Get short circuit results
            sc_data = GetData(sc_set_name, fault_attributes).result
            sc_data.insert(0, 'Scenario', [operation_scenario] * sc_data.shape[0])

            if fault_case == 'Max' and fault_type == '3psc':
                max_3psc_data = pd.concat([max_3psc_data, sc_data])
            elif fault_case == 'Min' and fault_type == '3psc':
                min_3psc_data = pd.concat([min_3psc_data, sc_data])
            elif fault_case == 'Max' and fault_type == 'spgf':
                max_spgf_data = pd.concat([max_spgf_data, sc_data])
            elif fault_case == 'Min' and fault_type == 'spgf':
                min_spgf_data = pd.concat([min_spgf_data, sc_data])

# Processing Data into report tables
fl_summary = fault_level_summary(max_3psc_data, min_3psc_data, max_spgf_data, min_spgf_data)

# Write results to xlsx workbook in separate sheets
with pd.ExcelWriter(file_path + 'FaultLevelResults.xlsx') as writer:
    max_3psc_data.to_excel(writer, sheet_name='Max_3pSCData', index=False)
    min_3psc_data.to_excel(writer, sheet_name='Min_3pSCData', index=False)
    max_spgf_data.to_excel(writer, sheet_name='Max_1pSCData', index=False)
    min_spgf_data.to_excel(writer, sheet_name='Min_1pSCData', index=False)
    fl_summary.to_excel(writer, sheet_name='FaultLevelSummary', index=False)
