import powerfactory as pf
from pfgadgets import CreateSet

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
