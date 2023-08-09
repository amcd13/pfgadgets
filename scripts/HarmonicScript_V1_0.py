import powerfactory as pf
import os
from pfgadgets import HarmonicLoadFlow, FrequencySweep, GetPlot, GetObject, GetData

"""
Version: 1.0

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script runs both harmonic load flow and frequency sweep calculations and exports associated plots and data

Inputs: external_grid_name - Name of the external grid element to modify
        distortion_plot_name - Name of the harmonic distortion plot page
        frequency_plot_name - Name of the frequency sweep plot page
        terminal_set_name - Name of the set containing refrences to terminals to grab data from after harmonic load flow
        terminal_attributes - List of attributes to export after harmonic load flow
"""

# Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

############################################################
# INPUTS

# Define name of external grid element
external_grid_name = 'External Grid'

# Define names of plot pages for harmonic distortion and frequency sweep
distortion_plot_name = 'Harmonic distortion'
frequency_plot_name = 'Frequency sweep'

# Define name of set and attributes to export data from
terminal_set_name = 'Harmonic Set'
terminal_attributes = ['m:THD']

############################################################

# Get external grid, plot pages and title objects
external_grid = GetObject(external_grid_name).obj
distortion_plot = GetPlot(distortion_plot_name, page_type='GrpPage')
frequency_plot = GetPlot(frequency_plot_name, page_type='GrpPage')
title = distortion_plot.title

# Clear title strings
title.sub1z = ''
title.sub2z = ''
title.sub3z = ''

# Define path for export
path = os.getcwd() + '/HarmonicResults'

# Define list of max or min scenario
max_or_min = ['Maximum', 'Minimum']

# Iterate through max and min scenarios
for i in range(2):
    # Set external grid to max or min for harmonic purposes
    external_grid.SetAttribute('cusedhrm', i)

    # Define file names for plot pages export
    harmonic_file_name = 'HarmonicDistortion_%s' % (max_or_min[i])
    frequency_file_name = 'FrequencySweep_%s' % (max_or_min[i])

    # Execute harmonic load flow
    HarmonicLoadFlow()

    # Get harmonic load flow data and export
    harmonic_data = GetData(terminal_set_name, terminal_attributes)
    harmonic_data.export(file_path=path, file_name=harmonic_file_name, replace=True)

    # Alter title block and export harmonic distortion plot
    title.sub1z = max_or_min[i] + ' Scenario'
    distortion_plot.export(file_path=path, file_name=harmonic_file_name, replace=True, frame=1)

    # Execute frequency sweep
    FrequencySweep()

    # Export frequency sweep plot
    frequency_plot.export(file_path=path, file_name=frequency_file_name, replace=True, frame=1)
