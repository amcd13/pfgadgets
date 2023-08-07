import powerfactory as pf
import os
import pandas as pd

app = pf.GetApplication()

#Define class to make sets
class CreateSet:
    def __init__(self, setName, *args, setType=5, studyCaseName=None):
        #Determine if set should be built in active study case or specified by study case name
        if studyCaseName==None:
            #Get active study case
            studyCase = app.GetActiveStudyCase()

            #Remove sets if duplicate
            allSets = studyCase.GetContents('*.SetSelect')
            for set in allSets:
                setElementName = set.loc_name
                if setElementName==setName:
                    set.Delete()
                else:
                    pass

            newSet = studyCase.CreateObject('SetSelect',setName)
            newSet.iused = setType

            self.set = newSet

#Define class for performing frequency sweep calculation
class FrequencySweep:
    def __init__(self, *args, method=0, init=0, start=50, stop=2500, results=None):
        #Get study case
        studyCase = app.GetActiveStudyCase()

        #Get frequency sweep command
        fsweep = app.GetFromStudyCase('ComFsweep')

        #If result input is given change to results input
        if results != None:
            resultFiles = studyCase.GetContents('*.ElmRes', 1)
            for resultFile in resultFiles:
                if resultFile.loc_name == results:
                    fsweep.p_resvar = resultFile

        #Initialise frequency sweep
        fsweep.iopt_net = method
        fsweep.ildfinit = init
        fsweep.fstart = start
        fsweep.fstop = stop

        #Execite frequency sweep
        fsweep.Execute()
        app.PrintPlain('Frequency sweep calculation successfully executed')

#Define class for collecting data after performing calculation
class GetData:
    def __init__(self, setName, attributeList, *args, resultHeadings=None):
        #Get study case and set
        studyCase = app.GetActiveStudyCase()
        generalSet = studyCase.GetContents('%s' % (setName))[0]

        #Define set object and name lists
        objDict = []
        nameDict = []
        typeDict = []

        #Collect objects/names from set
        for reference in generalSet.GetContents():
            elementObject = reference.obj_id
            elementName = elementObject.loc_name
            elementType = elementObject.typ_id
            objDict.append(elementObject)
            nameDict.append(elementName)
            typeDict.append(elementType)

        #Define column headings of results frame based on input
        if resultHeadings == None:
            headings = ['Name'] + attributeList
        else:
            headings = resultHeadings

        #Innitialise dataframe to store results
        resultsFrame = pd.DataFrame(columns = headings)

        #Collect results from set based on attirbute list
        for i in range(len(objDict)):
            obj = objDict[i]
            objType = typeDict[i]
            results = []
            results.append(obj.loc_name)
            for attribute in attributeList:
                try:
                    result = obj.GetAttribute(attribute)
                except AttributeError:
                    if not objType == None:
                        result = objType.GetAttribute(attribute)
                    else:
                        result = float("NaN")
                results.append(result)
            resultsFrame.loc[len(resultsFrame)] = results

        #Define class methods
        self.obj = objDict
        self.name = nameDict
        self.attribute =  attributeList
        self.result = resultsFrame

    #Define class method to export results
    def export(self, fileName = 'results', filePath =  None, replace=False):
        #Define file path to save results based on if filePath input given
        if filePath == None:
            path = os.getcwd() + ('\%s' % (fileName)) + '.csv'
        else:
            path = filePath + ('\%s' % (fileName)) + '.csv'

        if filePath != None:
            if not os.path.exists(filePath):
                os.makedirs(filePath)

        #If replace is False then edit file name to include (i) suffix
        if not replace:
            i = 1
            #Define file path to save results based on if filePath input given
            while os.path.isfile(filePath):
                if filePath == None:
                    path = os.getcwd() + ('\%s' % (fileName)) + ('(%s)' % (i)) + '.csv'
                    i += 1
                else:
                    path = filePath + ('\%s' % (fileName)) + ('(%s)' % (i)) + '.csv'
                    i += 1


        #Execute export of results to csv
        (self.result).to_csv(path_or_buf = path)
        app.PrintPlain("Exported results to: %s" % (path))

#Define class for aquiring element
class GetObject:
    def __init__(self, objectName):
        #Get network data folder
        networkData = app.GetProjectFolder('netdat')

        #Get element from network data folder
        element = networkData.GetContents(objectName, 1)[0]

        self.obj = element
        self.name = element.loc_name

#Define class for exporting plots
class GetPlot:
    def __init__(self, plotName, pageType='SetVipage'):
        setDesktop = app.GetGraphicsBoard()                         #Get graphics board
        plotPages = setDesktop.GetContents('*.%s' % (pageType))     #Get all plots

        #Iterate through plots and export
        for plot in plotPages:
            if plot.loc_name == plotName:
                self.plot = plot

        self.title = setDesktop.GetContents('*.SetTitm')[0]

    def export(self, *args, fileType='wmf', filePath=None, frame = 0, fileName=None, replace=False):
        wr = app.GetFromStudyCase('ComWr')                          #Get write command
        script = app.GetCurrentScript()                             #Get current working script
        scriptFilePath = os.getcwd()                                #Get file path of script

        #Show plot then scale axis/rebuild
        plot = self.plot
        plot.Show()
        plot.DoAutoScaleX()
        plot.DoAutoScaleY()
        app.Rebuild(2)

        #Setup file name for export to either default (plot name) or custom
        if fileName == None:
            name = plot.loc_name
        else:
            name = fileName

        #Set export path to either default (script directorty) or custom
        if filePath == None:
            path = scriptFilePath + '\\' + name + '.' + fileType
        else:
            path = filePath + '\\' + name + '.' + fileType

        #If replace is false then change file name to include (i) suffix
        if not replace:
            i = 1
            while os.path.isfile(path):
                if filePath == None:
                    path = scriptFilePath + '\\' + name + ('(%s)' % (i)) + '.' + fileType
                    i += 1
                else:
                    path = filePath + '\\' + name + ('(%s)' % (i)) + '.' + fileType

        #Initialise write command
        wr.iopt_rd = fileType
        wr.drawPageFrame = frame
        wr.f = path

        #Perform write command
        wr.Execute()
        app.PrintPlain("Exported '%s' plot to: %s" % (plot.loc_name, path))

#Define class for retriving data from results file
class GetResult:
    def __init__(self, resultsFile, *args, filePath=None, fileName='result', replace=None):
        #Get result export command
        res = app.GetFromStudyCase('ComRes')
        studyCase = app.GetActiveStudyCase()
        resultObject = studyCase.GetContents(resultsFile)[0]

        #Define file path to save results based on if filePath input given
        if filePath == None:
            path = os.getcwd() + ('\%s' % (fileName)) + '.csv'
        else:
            path = filePath + ('\%s' % (fileName)) + '.csv'

        if filePath != None:
            if not os.path.exists(filePath):
                os.makedirs(filePath)

        #If replace is False then edit file name to include (i) suffix
        if not replace:
            i = 1
            #Define file path to save results based on if filePath input given
            while os.path.isfile(filePath):
                if filePath == None:
                    path = os.getcwd() + ('\%s' % (fileName)) + ('(%s)' % (i)) + '.csv'
                    i += 1
                else:
                    path = filePath + ('\%s' % (fileName)) + ('(%s)' % (i)) + '.csv'
                    i += 1

        res.pResult = resultObject
        res.iopt_exp = 6
        res.f_name = path

#Define class for performing harmonic load flow calculation
class HarmonicLoadFlow:
    def __init__(self, *args, method=0):
        #Get harmonic load flow command and initialise
        hlf = app.GetFromStudyCase('ComHldf')
        hlf.iopt_net = method

        #Execute harmonic load flow
        hlf.Execute()
        app.PrintPlain('Harmonic load flow successfully executed')

#Define class for performing load flow calaculation
class LoadFlow:
    def __init__(self, *args, method=0, autoTap=0, feederScaling=0, opScen=None):
        #Get load flow command
        lf = app.GetFromStudyCase('ComLdf')

        #Change operation scenario to either active case or opScen input
        if opScen != None:
            operationScenarios = app.GetProjectFolder('scen').GetContents('*.IntScenario', 1)
            for operationScenario in operationScenarios:
                if operationScenario.loc_name == opScen:
                    relevantOperationScenario = operationScenario
                    relevantOperationScenario.Activate()

        #Initialise load flow
        lf.iopt_net = method
        lf.iopt_at = autoTap
        lf.iopt_fls = feederScaling

        #Execute load flow
        lf.Execute()
        app.PrintPlain('Load flow calculation successfully executed')

#Define class for performing short circuit calculation
class ShortCircuit:
    def __init__(self, objectName, faultType='3psc', calculate=0, setSelect=None, opScen=None):
        #Get study case
        studyCase = app.GetActiveStudyCase()

        #Change operation scenario to either active case or opScen input
        if opScen != None:
            operationScenarios = app.GetProjectFolder('scen').GetContents('*.IntScenario', 1)
            for operationScenario in operationScenarios:
                if operationScenario.loc_name == opScen:
                    relevantOperationScenario = operationScenario
                    relevantOperationScenario.Activate()

        #Clean up short cirucit objects in study case
        shc = studyCase.GetContents('*.SetTitm')
        titleOld = studyCase.GetContents('*.SetTitm')
        if not (shc or titleOld):
            pass
        else:
            for i in shc:
                i.Delete()
            for t in titleOld:
                t.Delete()

        #Get short circuit command and initialise
        shc = app.GetFromStudyCase('ComShc')
        shc.iopt_asc = 0
        shc.iopt_allbus = 0
        shc.iopt_mde = 1
        shc.iopt_shc = faultType
        shc.iopt_cur = calculate

        #Collect set objects if setSelect=1 else single object
        if setSelect != None:
            #Define terminal object and name lists
            objDict = []
            nameDict = []

            #Get terminal set
            generalSet = studyCase.GetContents('%s' % (objectName))[0]

            #Collect objects/names from set
            for reference in generalSet.GetContents():
                elementObject = reference.obj_id
                objDict.append(elementObject)
                elementName = elementObject.loc_name
                nameDict.append(elementName)
            shc.shcobj = generalSet
        else:
            networkData = app.GetProjectFolder('netdat')
            obj = networkData.GetContents(objectName, 1)[0]
            shc.shcobj = obj

        #Execute short circuit command
        shc.Execute()
        if setSelect != None:
            app.PrintPlain('%s short-circuit calculation successfully executed @ %s' % (faultType,nameDict))
            #Define class methods
            self.obj = objDict
            self.name = nameDict
        else:
            app.PrintPlain('%s short-circuit calculation successfully executed @ %s' % (faultType, objectName))
            #Define class methods
            self.obj = obj
            self.name = objectName












