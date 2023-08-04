import powerfactory as pf
import os
import pandas as pd
from pfgadgets import ShortCircuit, GetPlot, GetObject

"""
Version: 1.1

Author: Andrew McDermott

Dependencies: Ppfgadgets module

Description: This script sets up a TOC plot page with relevant devices, runs fault cases and exports the plot

Inputs: faultTypes -List of types of faults to run during short-circuit calculation
        maximum - Name of operation scenario used for maximum fault case
        minimum - Name of operation scenario used for minimum fault case
        deviceDict - List of dictionaries containing the terminal to run short-circuit calculation on and associated TOC plot
"""

#Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

######### INPUTS ############################################

#Define fault types to run
faultTypes = ['3psc','spgf']

#Define maximum and minimum operation scenarios
maximum = ['Operation Scenario']
minimum = ['Operation Scenario']

#Define dictionary contatining terminals to run short-circuit and associated devices to display in TOC plot
deviceDict = [{'terminal': ['A0S01-SB-NF-001'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NF-001']},
              {'terminal': ['A0S01-TX-FB-001 HV Bushing'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NB-001']},
              {'terminal': ['A0S01-SB-NB-001'], 'plot': ['P0S01-SB-NF-001_A0S01-SB-NB-001']},
              {'terminal': ['A0S02-TX-FB-001 HV Bushing'], 'plot': ['P0S01-SB-NF-001_A0S02-SB-NB-001']},
              {'terminal': ['A0S02-SB-NB-001'], 'plot': ['P0S01-SB-NF-001_A0S02-SB-NB-001']},
              {'terminal': ['A0S01-TX-FB-002-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-003-RMU']},
              {'terminal': ['A0S01-TX-FB-003-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-003-RMU']},
              {'terminal': ['A0S01-TX-FB-004-RMU'], 'plot': ['P0S01-SB-NF-001_A0S01-TX-FB-004-RMU']}]

#############################################################



path = os.getcwd() + '/TOCPlots'

#Iterate through length of device dictrionary
for i in range(len(deviceDict)):
    j = len(deviceDict) - i
    app.PrintPlain('\nTerminals remaining: %s' % (j))
    #Iterate through fault types
    for faultType in faultTypes:
        #Iterate through maximum and minimum fault cases
        for calculate in range(0,2):
            #Get plot, plot title and plot settings
            plotName = deviceDict[i]['plot'][0]
            plot = GetPlot(plotName, pageType='SetVipage')
            plotTitle = plot.title
            plotSettings = plot.plot.GetContents('*.VisOcplot')[0]
            OCPLotSettings = plotSettings.GetContents('*.SetOcplt')[0]
            
            #Define lists to store terminal and associated device objects defined in deviceDict
            terminalList = []
            terminalNames = deviceDict[i]['terminal']

            #Grab PF objects from names in deviceDict and store in list
            for terminalName in terminalNames:
                terminal = GetObject(terminalName).obj
                terminalList.append(terminal)
            
            #Change fault calculation based on maximum or minimum case
            if calculate == 0:
                #Execute ShortCircuit maximum command
                ShortCircuit(terminal.loc_name, faultType=faultType, calculate=calculate, setSelect=None, opScen=maximum)
                #Define file name for export
                fileName = 'ShortCircuitData_%s_%s' % (faultType, 'Max')
                maxOrMin = 'Maximum'
            else:
                #Execute ShortCircuit minimum command
                ShortCircuit(terminal.loc_name, faultType=faultType, calculate=calculate, setSelect=None, opScen=minimum)
                #Define file name for export
                fileName = 'ShortCircuitData_%s_%s' % (faultType, 'Min')
                maxOrMin = 'Minimum'

            if faultType == '3psc':
                OCPLotSettings.SetAttribute('ishow', 1)
            elif faultType == 'spgf':
                OCPLotSettings.SetAttribute('ishow', 3)

            #Adjust title strings
            plotTitle.sub1z = terminalList[0].loc_name
            plotTitle.sub2z = faultType
            plotTitle.sub3z = maxOrMin

            #Define file name and export plot
            fileName = '%s_%s_%s' % (terminalList[0].loc_name, faultType, maxOrMin)
            plot.export(filePath=path, fileName=fileName, replace=True)
