import powerfactory as pf
import os
import pandas as pd
from pfgadgets import HarmonicLoadFlow, FrequencySweep, GetPlot, GetObject, GetData

"""
Version: 1.0

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script runs both harmonic load flow and frequency sweep calculations and exports associated plots and data

Inputs: externalGridName - Name of the external grid element to modify
        distortionPlotName - Name of the harmonic distortion plot page
        frequencyPlotName - Name of the frequency sweep plot page
        terminalSetName - Name of the set containing refrences to terminals to grab data from after harmonic load flow
        terminalAttributes - List of attributes to export after harmonic load flow
"""

#Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

########### INPUTS #########################################

#Define name of external grid element
externalGridName = 'External Grid'

#Define names of plot pages for harmonic distortion and frequency sweep
distortionPlotName = 'Harmonic distortion'
frequencyPlotName = 'Frequency sweep'

#Define name of set and attributes to export data from
terminalSetName = 'Harmonic Set'
terminalAttributes = ['m:THD']

############################################################

#Get external grid, plot pages and title objects
externalGrid = GetObject(externalGridName).obj
distortionPlot = GetPlot(distortionPlotName, pageType='GrpPage')
frequencyPlot = GetPlot(frequencyPlotName, pageType='GrpPage')
title = distortionPlot.title

#Clear title strings
title.sub1z = ''
title.sub2z = ''
title.sub3z = ''

#Define path for export
path = os.getcwd() + '/HarmonicResults'

#Define list of max or min scenario
maxOrMin = ['Maximum', 'Minimum']

#Iterrate through max and min scenarios
for i in range(2):
    #Set external grid to max or min for harmonic purposes
    externalGrid.SetAttribute('cusedhrm', i)

    #Define file names for plot pages export
    harmonicFileName = 'HarmonicDistortion_%s' % (maxOrMin[i])
    frequencyFileName = 'FrequencySweep_%s' % (maxOrMin[i])

    #Execute harmonic load flow
    HarmonicLoadFlow()

    #Get harmonic load flow data and export
    harmonicData = GetData(terminalSetName, terminalAttributes)
    harmonicData.export(filePath=path, fileName=harmonicFileName, replace=True)

    #Alter title block and export harmonic distortion plot
    title.sub1z = maxOrMin[i] + ' Scenario'
    distortionPlot.export(filePath=path, fileName=harmonicFileName, replace=True, frame=1)

    #Execute frequency sweep 
    FrequencySweep()

    #Export frequency sweep plot
    frequencyPlot.export(filePath=path, fileName=frequencyFileName, replace=True, frame=1)
