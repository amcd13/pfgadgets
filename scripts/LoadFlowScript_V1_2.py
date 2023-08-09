import powerfactory as pf
import os
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

# Define fault types to run
fault_types = ['3psc', 'spgf']

# Define maximum and minimum operation scenarios
maximum = ['Operation Scenario']
minimum = ['Operation Scenario']

# Setup set names
sc_set_name = 'Short-Circuit Set'

# Setup object attributes to collect from load flow results
sc3p_attributes = ['m:Ikss', 'm:Skss']
scgf_attributes = ['m:Ikss:A', 'm:Skss:A']

#############################################################

# Define file path for export
file_path = os.getcwd() + '\\ShortCircuitResults'

# Iterate through fault types
for fault_type in fault_types:
    # Iterate through maximum and minimum fault cases
    for calculate in range(0, 2):
        # Change fault calculation based on maximum or minimum case
        if calculate == 0:
            # Execute ShortCircuit maximum command
            ShortCircuit(sc_set_name, fault_type=fault_type, calculate=calculate, set_select=True, op_scen=maximum)
            # Define file name for export
            file_name = 'ShortCircuitData_%s_%s' % (fault_type, 'Max')
        else:
            # Execute ShortCircuit minimum command
            ShortCircuit(sc_set_name, fault_type=fault_type, calculate=calculate, set_select=True, op_scen=minimum)
            # Define file name for export
            file_name = 'ShortCircuitData_%s_%s' % (fault_type, 'Min')

        # Collect data based on fault type
        if fault_type == '3psc':
            # Collect objects data
            sc_data = GetData(sc_set_name, sc3p_attributes)
        elif fault_type == 'spgf':
            # Collect objects data
            sc_data = GetData(sc_set_name, scgf_attributes)

        # Export data
        sc_data.export(file_name=file_name, file_path=file_path, replace=True)
