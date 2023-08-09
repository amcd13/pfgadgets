import powerfactory as pf
from pfgadgets import CreateSet

"""
Version: 1.0

Author: Andrew McDermott

Dependencies: pfgadgets module

Description: This script creates sets intended to store references to
             objects relevant to load flow and short-circuit calculations

Inputs: general_set_names - List of names of general sets to create
        sc_set_names - List of name of short-circuit sets to create
"""

# Get PowerFactory application
app = pf.GetApplication()
app.EchoOff()
app.ClearOutputWindow()

############################################################
# INPUTS

# Define set names
general_set_names = ['Terminals', 'Lines', 'Transformers', 'Harmonic Set']
sc_set_names = ['Short-Circuit Set']

############################################################

# Create sets
for general_set_name in general_set_names:
    CreateSet(general_set_name)
for sc_set_name in sc_set_names:
    CreateSet(sc_set_name, set_type=1)
