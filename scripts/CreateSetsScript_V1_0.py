import powerfactory as pf
import os
import pandas as pd
from PFClasses_V1_1 import createSet

"""
Version: 1.0

Author: Andrew McDermott

Dependencies: PFClasses module

Description: This script creates sets intended to store references to
             objects relevant to load flow and short-circuit calculations

Inputs: generalSetNames - List of names of general sets to create
        scSetNames - List of name of short-circuit sets to create
"""

#Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

########### INPUTS #########################################

#Define set names
generalSetNames = ['Terminals', 'Lines', 'Transformers', 'Harmonic Set']
scSetNames = ['Short-Circuit Set']

############################################################

#Create sets
for generalSetName in generalSetNames:
    createSet(generalSetName)
for scSetName in scSetNames:
    createSet(scSetName, setType=1)




