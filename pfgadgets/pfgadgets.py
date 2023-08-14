import powerfactory as pf
import os
import pandas as pd

app = pf.GetApplication()


# Define class to make sets
class CreateSet:
    def __init__(self, set_name, set_type=5, sc_name=None, replace=True):
        # Determine if set should be built in active study case or specified by study case name
        if sc_name is None:
            # Get active study case
            study_case = app.GetActiveStudyCase()
        else:
            # Get specified study case
            study_cases = app.GetProjectFolder('study').GetContents('*.IntCase', 1)

            try:
                study_case = [k for k in study_cases if k.loc_name == sc_name][0]
            except IndexError:
                raise Exception('Could not find study case with name: %s' % sc_name)

        #Get all sets from study case
        all_sets = study_case.GetContents('*.SetSelect', 1)

        # Remove sets if duplicate
        for set in all_sets:
            set_elm_name = set.loc_name
            if set_elm_name == set_name:
                if replace:
                    set.Delete()
                else:
                    app.PrintWarn('Set with name %s already exists' % set_name)
            else:
                pass

        new_set = study_case.CreateObject('SetSelect', set_name)
        new_set.iused = set_type

        self.set = new_set


# Define class for performing frequency sweep calculation
class FrequencySweep:
    def __init__(self, method=0, init=0, start=50, stop=2500, results=None):
        # Get study case
        study_case = app.GetActiveStudyCase()

        # Get frequency sweep command
        f_sweep = app.GetFromStudyCase('ComFsweep')

        # If result input is given change to results input
        if results is not None:
            result_files = study_case.GetContents('*.ElmRes', 1)
            for resultFile in result_files:
                if resultFile.loc_name == results:
                    f_sweep.p_resvar = resultFile

        # Initialise frequency sweep
        f_sweep.iopt_net = method
        f_sweep.ildfinit = init
        f_sweep.fstart = start
        f_sweep.fstop = stop

        # Execite frequency sweep
        f_sweep.Execute()
        app.PrintPlain('Frequency sweep calculation successfully executed')


# Define class for collecting data after performing calculation
class GetData:
    def __init__(self, set_name, attribute_list, result_headings=None):
        # Get study case and set
        study_case = app.GetActiveStudyCase()

        try:
            general_set = study_case.GetContents('%s' % set_name)[0]
        except IndexError:
            raise Exception('Cannot find set with name: %s' % set_name)

        # Define set object and name lists
        obj_dict = []
        name_dict = []
        type_dict = []

        # Collect objects/names from set
        for reference in general_set.GetContents():
            element_object = reference.obj_id
            element_name = element_object.loc_name
            element_type = element_object.typ_id
            obj_dict.append(element_object)
            name_dict.append(element_name)
            type_dict.append(element_type)

        # Define column headings of results frame based on input
        if result_headings is None:
            headings = attribute_list
        else:
            headings = result_headings

        # Initialize dataframe to store results
        results_frame = pd.DataFrame(columns=headings)

        # Collect results from set based on attribute list
        for i in range(len(obj_dict)):
            obj = obj_dict[i]
            obj_type = type_dict[i]
            results = []

            # Iterate through attributes
            for attribute in attribute_list:
                # Try grab object attribute
                try:
                    result = obj.GetAttribute(attribute)
                # Except can't find object attribute
                except AttributeError:
                    # Check if object has type
                    if obj_type is not None:
                        # Try grab attribute from object type
                        try:
                            result = obj_type.GetAttribute(attribute)
                        # Can't find attribute so throw error and exit
                        except AttributeError:
                            raise Exception('Cannot retrieve object attribute: %s' % attribute)

                    # If object has no type replace result with NaN and print warning
                    else:
                        result = float("NaN")
                        app.PrintWarn('%s does not have attribute %s' % (obj, attribute))
                results.append(result)
            results_frame.loc[len(results_frame)] = results

        # Define class methods
        self.obj = obj_dict
        self.name = name_dict
        self.attribute = attribute_list
        self.result = results_frame

    # Define class method to export results
    def export(self, file_name='results', file_path=None, replace=False):
        # Define file path to save results based on if filePath input given
        if file_path is None:
            path = os.getcwd() + ('\\%s' % file_name) + '.csv'
        else:
            path = file_path + ('\\%s' % file_name) + '.csv'

        # Make file path if it doesn't exist
        if file_path is not None:
            if not os.path.exists(file_path):
                os.makedirs(file_path)

        # If replace is False then edit file name to include (i) suffix
        if not replace:
            i = 1
            # Define file path to save results based on if filePath input given
            while os.path.isfile(file_path):
                if file_path is None:
                    path = os.getcwd() + ('\\%s' % file_name) + ('(%s)' % i) + '.csv'
                    i += 1
                else:
                    path = file_path + ('\\%s' % file_name) + ('(%s)' % i) + '.csv'
                    i += 1

        # Execute export of results to csv
        self.result.to_csv(path_or_buf=path)
        app.PrintPlain("Exported results to: %s" % path)


# Define class for acquiring element
class GetObject:
    def __init__(self, object_name):
        # Get network data folder
        network_data = app.GetProjectFolder('netdat')

        # Get element from network data folder
        element_list = network_data.GetContents(object_name, 1)

        # Check if element_list is empty
        if element_list:
            element = element_list[0]
        # Throw error if element_list is empty
        else:
            raise Exception('Could not retrieve element: %s' % object_name)

        # Throw warning if element_list contains more than one element
        if len(element_list) > 1:
            app.PrintWarn('More than one element found with name: %s' % object_name)

        self.obj = element
        self.name = element.loc_name


# Define class for exporting plots
class GetPlot:
    def __init__(self, plot_name, page_type='SetVipage'):
        set_desktop = app.GetGraphicsBoard()  # Get graphics board
        plot_pages = set_desktop.GetContents('*.%s' % page_type)  # Get all plots

        # Grab list of plots with plot_name
        plot_list = [plot for plot in plot_pages if plot.loc_name == plot_name]

        # Check if plot was found, else throw error
        if plot_list:
            plot = plot_list[0]
        else:
            raise Exception('No plot found with name: %s' % plot_name)

        # Check if more than one plot found and throw warning
        if len(plot_list) > 1:
            app.PrintWarn('More than one element found with name: %s' % plot_name)

        self.plot = plot
        self.title = set_desktop.GetContents('*.SetTitm')[0]

    def export(self, file_type='wmf', file_path=None, frame=0, file_name=None, replace=False, auto_scale=True):
        wr = app.GetFromStudyCase('ComWr')  # Get write command
        script_file_path = os.getcwd()  # Get file path of script

        # Show plot then scale axis/rebuild
        plot = self.plot
        plot.Show()

        # Scale plot depending on scale input
        if scale:
            plot.DoAutoScaleX()
            plot.DoAutoScaleY()
        else:
            pass
        
        app.Rebuild(2)

        # Setup file name for export to either default (plot name) or custom
        if file_name is None:
            name = plot.loc_name
        else:
            name = file_name

        # Set export path to either default (script directorty) or custom
        if file_path is None:
            path = script_file_path + '\\' + name + '.' + file_type
        else:
            path = file_path + '\\' + name + '.' + file_type

        # If replace is false then change file name to include (i) suffix
        if not replace:
            i = 1
            while os.path.isfile(path):
                if file_path is None:
                    path = script_file_path + '\\' + name + ('(%s)' % i) + '.' + file_type
                    i += 1
                else:
                    path = file_path + '\\' + name + ('(%s)' % i) + '.' + file_type

        # Initialise write command
        wr.iopt_rd = file_type
        wr.drawPageFrame = frame
        wr.f = path

        # Perform write command
        wr.Execute()
        app.PrintPlain("Exported '%s' plot to: %s" % (plot.loc_name, path))


# Define class for retrieving data from results file
class GetResult:
    def __init__(self, results_file, file_path=None, file_name='result', replace=None):
        # Get result export command
        res = app.GetFromStudyCase('ComRes')
        study_case = app.GetActiveStudyCase()

        # Retrieve result objects
        result_object_list = study_case.GetContents(results_file, 1)

        # Get result object if list is not empty
        if result_object_list:
            result_object = result_object_list[0]
        else:
            raise Exception('Cannot retrieve result file with name: %s' % results_file)

        # Check is more than one object was retrieved and trow warning
        if len(result_object_list) > 1:
            app.PrintWarn('More than one element found with name: %s' % results_file)

        # Define file path to save results based on if filePath input given
        if file_path is None:
            path = os.getcwd() + ('\\%s' % file_name) + '.csv'
        else:
            path = file_path + ('\\%s' % file_name) + '.csv'

        if file_path is not None:
            if not os.path.exists(file_path):
                os.makedirs(file_path)

        # If replace is False then edit file name to include (i) suffix
        if not replace:
            i = 1
            # Define file path to save results based on if filePath input given
            while os.path.isfile(file_path):
                if file_path is None:
                    path = os.getcwd() + ('\\%s' % file_name) + ('(%s)' % i) + '.csv'
                    i += 1
                else:
                    path = file_path + ('\\%s' % file_name) + ('(%s)' % i) + '.csv'
                    i += 1

        res.pResult = result_object
        res.iopt_exp = 6
        res.f_name = path


# Define class for performing harmonic load flow calculation
class HarmonicLoadFlow:
    def __init__(self, method=0):
        # Get harmonic load flow command and initialise
        hlf = app.GetFromStudyCase('ComHldf')
        hlf.iopt_net = method

        # Execute harmonic load flow
        hlf.Execute()
        app.PrintPlain('Harmonic load flow successfully executed')


# Define class for performing load flow calaculation
class LoadFlow:
    def __init__(self, method=0, auto_tap=0, feeder_scaling=0, op_scen=None):
        # Get load flow command
        lf = app.GetFromStudyCase('ComLdf')

        # Change operation scenario to either active case or opScen input
        if op_scen is not None:
            op_scens = app.GetProjectFolder('scen').GetContents('*.IntScenario', 1)
            relevant_op_scen = [k for k in op_scens if k.loc_name == op_scen]

            # Check if operation scenario was found
            if relevant_op_scen:
                relevant_op_scen[0].Activate()
            else:
                raise Exception('Cannot find operation scenario: %s' % op_scen)

        # Initialise load flow
        lf.iopt_net = method
        lf.iopt_at = auto_tap
        lf.iopt_fls = feeder_scaling

        # Execute load flow
        lf.Execute()
        app.PrintPlain('Load flow calculation successfully executed')


# Define class for performing short circuit calculation
class ShortCircuit:
    def __init__(self, object_name, fault_type='3psc', calculate=0, set_select=None, op_scen=None):
        # Get study case
        studyCase = app.GetActiveStudyCase()

        # Change operation scenario to either active case or opScen input
        if op_scen is not None:
            op_scens = app.GetProjectFolder('scen').GetContents('*.IntScenario', 1)
            relevant_op_scen = [k for k in op_scens if k.loc_name == op_scen]

            # Check if operation scenario was found
            if relevant_op_scen:
                relevant_op_scen[0].Activate()
            else:
                raise Exception('Cannot find operation scenario: %s' % op_scen)

        # Clean up short circuit objects in study case
        shc = studyCase.GetContents('*.SetTitm')
        title_old = studyCase.GetContents('*.SetTitm')
        if not (shc or title_old):
            pass
        else:
            for i in shc:
                i.Delete()
            for t in title_old:
                t.Delete()

        # Get short circuit command and initialise
        shc = app.GetFromStudyCase('ComShc')
        shc.iopt_asc = 0
        shc.iopt_allbus = 0
        shc.iopt_mde = 1
        shc.iopt_shc = fault_type
        shc.iopt_cur = calculate

        # Collect set objects if setSelect=1 else single object
        if set_select is not None:
            # Define terminal object and name lists
            obj_dict = []
            name_dict = []

            # Get terminal set
            general_set = studyCase.GetContents('%s' % object_name)[0]

            # Collect objects/names from set
            for reference in general_set.GetContents():
                elementObject = reference.obj_id
                obj_dict.append(elementObject)
                elementName = elementObject.loc_name
                name_dict.append(elementName)
            shc.shcobj = general_set
        else:
            network_data = app.GetProjectFolder('netdat')
            try:
                obj = network_data.GetContents(object_name, 1)[0]
            except IndexError:
                raise Exception('Unable to find element: %s' % object_name)
            shc.shcobj = obj

        # Execute short circuit command
        shc.Execute()

        if set_select is not None:
            app.PrintPlain('%s short-circuit calculation successfully executed @ %s' % (fault_type, name_dict))
            # Define class methods
            self.obj = obj_dict
            self.name = name_dict
        else:
            app.PrintPlain('%s short-circuit calculation successfully executed @ %s' % (fault_type, object_name))
            # Define class methods
            self.obj = obj
            self.name = object_name
