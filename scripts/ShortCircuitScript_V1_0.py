import powerfactory as pf
import os
import pandas as pd
from pfgadgets import ShortCircuit, GetData

"""
Version: 1.0

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script excutes a short-circuit calculation and exports
             the results to csv files

Inputs: faultTypes - The types of faults to run
        maximum - The name of the maximum fault operation scenario
        minimum - The name of the minimum fault operation scenario
        scSetName - The name of the set containing the objects to run the short-circuit calcualtion on
        sc3pscAttributes - A list of variables to extract during 3-phase faults
        scSpgfAttributes - A list of variables to extract during ground faults
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

#Setup set names
scSetName = 'Short-Circuit Set'

#Setup object attributes to collect from load flow results
sc3pscAttributes = ['m:Ikss','m:Skss']
scSpgfAttributes = ['m:Ikss:A','m:Skss:A']

#############################################################

#Define file path for export
filePath = os.getcwd() + '\ShortCircuitResults'

#Iterate through fault types
for faultType in faultTypes:
    #Iterate through maximum and minimum fault cases
    for calculate in range(0,2):
        #Change fault calculation based on maximum or minimum case
        if calculate == 0:
            #Execute ShortCircuit maximum command
            ShortCircuit(scSetName, faultType=faultType, calculate=calculate, setSelect=True, opScen=maximum)
            #Define file name for export
            fileName = 'ShortCircuitData_%s_%s' % (faultType, 'Max')
        else:
            #Execute ShortCircuit minimum command
            ShortCircuit(scSetName, faultType=faultType, calculate=calculate, setSelect=True, opScen=minimum)
            #Define file name for export
            fileName = 'ShortCircuitData_%s_%s' % (faultType, 'Min')

        #Collect data based on fault type
        if faultType == '3psc':
            #Collect objects data
            scData = GetData(scSetName, sc3pscAttributes)
        elif faultType == 'spgf':
            #Collect objects data
            scData = GetData(scSetName, scSpgfAttributes)

        #Export data    
        scData.export(fileName=fileName, filePath=filePath, replace=True)

