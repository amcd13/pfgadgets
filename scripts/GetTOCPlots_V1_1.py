import powerfactory as pf
import os
from pfgadgets import ShortCircuit, GetPlot, GetObject

"""
Version: 1.1

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script sets up a TOC plot page with relevant devices, runs fault cases and exports the plot

Inputs: fault_types -List of types of faults to run during short-circuit calculation
        maximum - Name of operation scenario used for maximum fault case
        minimum - Name of operation scenario used for minimum fault case
        plot_dict - List of dictionaries containing the terminal to run short-circuit calculation on and associated TOC plot
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
maximum = 'Maximum Normal Load'
minimum = 'Maximum Normal Load'

# Define dictionary containing terminals to run short-circuit and associated devices to display in TOC plot
plot_dict = [{'terminal': ['A0S01-SB-NF-001'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NF-001']},
             {'terminal': ['A0S01-TX-FB-001 HV Bushing'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NB-001']},
             {'terminal': ['A0S01-SB-NB-001'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NB-001']},
             {'terminal': ['A0S02-TX-FB-001 HV Bushing'], 'plot': ['P0S01-SB-NF-001_A0S02-SB-NB-001']},
             {'terminal': ['A0S02-SB-NB-001'], 'plot': ['P0S01-SB-NF-001_A0S02-SB-NB-001']},
             {'terminal': ['A0S01-TX-FB-002-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-003-RMU']},
             {'terminal': ['A0S01-TX-FB-003-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-003-RMU']},
             {'terminal': ['A0S01-TX-FB-004-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-004-RMU']}]

#############################################################


path = os.getcwd() + '/TOCPlots'

# Iterate through length of device dictionary
for i in range(len(plot_dict)):
    j = len(plot_dict) - i
    # Iterate through fault types
    for fault_type in fault_types:
        # Iterate through maximum and minimum fault cases
        for calculate in range(0, 2):
            # Get plot, plot title and plot settings
            plot_name = plot_dict[i]['plot'][0]
            plot = GetPlot(plot_name, page_type='SetVipage')
            plot_title = plot.title
            plot_settings = plot.plot.GetContents('*.VisOcplot')[0]
            oc_plot_settings = plot_settings.GetContents('*.SetOcplt')[0]

            # Define lists to store terminal and associated device objects defined in plotDict
            terminal_list = []
            terminal_names = plot_dict[i]['terminal']

            # Grab PF objects from names in plotDict and store in list
            for terminalName in terminal_names:
                terminal = GetObject(terminalName).obj
                terminal_list.append(terminal)

            # Change fault calculation based on maximum or minimum case
            if calculate == 0:
                # Execute ShortCircuit maximum command
                ShortCircuit(terminal.loc_name, fault_type=fault_type, calculate=calculate, set_select=None,
                             op_scen=maximum)
                # Define file name for export
                file_name = 'ShortCircuitData_%s_%s' % (fault_type, 'Max')
                max_or_min = 'Maximum'
            else:
                # Execute ShortCircuit minimum command
                ShortCircuit(terminal.loc_name, fault_type=fault_type, calculate=calculate, set_select=None,
                             op_scen=minimum)
                # Define file name for export
                file_name = 'ShortCircuitData_%s_%s' % (fault_type, 'Min')
                max_or_min = 'Minimum'

            if fault_type == '3psc':
                oc_plot_settings.SetAttribute('ishow', 1)
            elif fault_type == 'spgf':
                oc_plot_settings.SetAttribute('ishow', 3)

            # Adjust title strings
            plot_title.sub1z = terminal_list[0].loc_name
            plot_title.sub2z = fault_type
            plot_title.sub3z = max_or_min

            app.PrintPlain('\nTerminals remaining: %s' % j)

            # Define file name and export plot
            file_name = '%s_%s_%s' % (terminal_list[0].loc_name, fault_type, max_or_min)
            plot.export(file_path=path, file_name=file_name, replace=True)
