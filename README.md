# pfgadgets package

Description: This package is designed to assist in the use of python within the PowerFactory environment

## Classes List:
- CreateSet
- FrequencySweep
- GetData
- GetObject
- GetPlot
- GetResult
- HarmonicLoadFlow
- LoadFlow
- ShortCircuit

## Class Description
### pfgadgets.CreateSet
*__CreateSet(set_name, set_type=5, sc_name=None)__*

*Description*: This class is designed to create a set

*Parameters*: 
- set_name - Name of set to create
- set_type - Defines the type of set to create. Refer PF for types
- sc_name - Defines the name of the study case to create set in. If None then will create in active study case

*Attributes*: 
- .set - Returns the created set object
---
### pfgadgets.FrequencySweep
*__FrequencySweep(method=0, init=0, start=50, stop=2500, results=None)__*

*Description*: This class is designed to perform a frequency sweep calculation

*Parameters*:
- method - defines the frequency sweep method (0 - balanced, 1 - unbalanced)
- init - defines whether to initialise with load flow
- start - defines the starting frequency
- stop - defines the ending frequency
- results - defines which results variable to save results to
---
### pfgadgets.GetData
*__GetData(set_name, attribute_list, result_headings=None)__*

*Description*: This class is designed to collect data from a set and offers export to csv option

*Parameters*: 
- set_name - Name of the set to extract data from
- attribute_list - List of attributes to get data from
- result_headings - List of strings for data frame headings

*Attributes*: 
- .obj - Returns object/s that data was collected from
- .name - Returns name of object/s that data was collected from
- .attribute - Returns list of attributes that was collected
- .result - Returns the results data frame that holds the extracted data
---
### pfgadgets.GetData.export
*__pfgadgets.GetData.export(file_name='results', file_path=None, replace=False)__*

*Description*: Exports the collected data to a csv

*Parameters*: 
- file_name - Name of file that data will export to. Should not include file extension.
- file_path - Defines directory where csv will export to. If None then export to script directory
- replace - Defines whether to replace existing files or not. (False - do not replace existing, True - replace existing)
---
### pfgadgets.GetObject
*__pfgadgets.GetObject(object_name)__*

*Description*: This class is designed to collect a given object

*Parameters*: 
- object_name - Name of object to collect

*Attributes*: 
- .obj - Returns the collected object
- .name - Returns the name of the collected object
---
### pfgadgets.GetPlot
*__pfgadgets.GetPlot(plot_name, page_type='SetVipage')__*

*Description*: This class is designed to collect a plot page and offers an export option

*Parameters*: 
- plot_name - Name of plot to be collected
- page_type - Defined the type of page to be collected
	
*Attributes*: 
- .plot - Returns to plot objects that was collected
- .title - Returns the title object that may be included in a plot page
---
### pfgadgets.GetPlot.export
*__pfgadgets.GetPlot.export(file_type='wmf', file_path=None, frame = 0, file_name=None, replace=False)__*

*Description*: Exports the plot page to the desired type and location

*Parameters*: 
- file_type - Defines the extension of the exported plot page file
- file_path - Defines directory where csv will export to. If None then export to script directory
- frame - Defines whether to include a page frame or not (0 - no frame, 1 - include frame)
- file_name - Name of file that data will export to. Should not include file extension.
- replace - Defines whether to replace existing files or not. (False - do not replace existing, True - replace existing)
---
### pfgadgets.GetResult
*__pfgadgets.GetResult(results_file, file_path=None, file_name='result', replace=None)__*

*Description*: This class is designed to export a result object

*Parameters*:
- results_file - Name of the results object to grab from PowerFactory
- file_path - Defines directory where csv will export to. If None then export to script directory
- file_name - Name of file that data will export to. Should not include file extension.
- replace - Defines whether to replace existing files or not. (False - do not replace existing, True - replace existing)
---
### pfgadgets.HarmonicLoadFlow
*__pfgadgets.HarmonicLoadFlow(method=0)__*

*Description*: This class is designed to perform a harmonic load flow calculation

*Parameters*: 
- method - defines the harmonic load flow method (0 - balanced, 1 - unbalanced)
---
### pfgadgets.LoadFlow
*__pfgadgets.LoadFlow(method=0, auto_tap=0, feeder_scaling=0, op_scen=None)__*

*Description*: This class is designed to perform a load flow calculation

*Parameters*: 
- method - defines the load flow method (0 - balanced, 1 - unbalanced)
- auto_tap - defines whether to enable automatic tap changing (0 - off, 1 - on)
- feeder_scaling - defines whether to enable feeder load scaling (0 - off, 1 - on)
- op_scen - defines the operation scenario to activate before calculation (if None then no operation scenario will be activated)
---
### pfgadgets.ShortCircuit
*__pfgadgets.ShortCircuit(object_name, faultType='3psc', calculate=0, set_select=None, op_scen=None)__*

*Description*: This class is designed to perform a short-circuit calculation

*Parameters*: 
- object_name - Name of object to perform short-circuit calculation on. If setSelect!=None then a set can be passed instead of an object
- fault_type - Defines the type of fault to use during short-circuit calculation
- calculate - Defines whether to use maximum or minimum fault calculation (0 - maximum, 1 - minimum)
- set_select - Defines whether a set or single object is used for calculation (None - single object, != None - set)
- op_scen - defines the operation scenario to activate before calculation (if None then no operation scenario will be activated)

*Attributes*: 
- .obj - Returns object/s where short-circuit calculation was performed
- .name - Returns name/s of objects where short-circuit calculation was performed
