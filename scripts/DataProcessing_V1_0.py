import copy
import pandas as pd
import math

def busbar_util (term_data):
    df = term_data

    # Initialise dataframes to store results
    busbar_util_frame = pd.DataFrame()

    # Organising Results
    busbar_util_frame['Tag Name'] = df['loc_name']
    busbar_util_frame['Rated Voltage (kV)'] = round(df['e:uknom'],2)
    busbar_util_frame['Rated Continuous Current (A)'] = round(df['Ir']*1000,0)
    busbar_util_frame['Maximum Current (A)'] = round(df['m:Sout']/(df['m:Ul']*math.sqrt(3)), 0)
    busbar_util_frame['Utilisation (%)'] = round(busbar_util_frame['Maximum Current (A)']/busbar_util_frame['Rated Continuous Current (A)']*100, 1)
    busbar_util_frame['Scenario'] = df['Scenario']

    return busbar_util_frame

def busbar_voltage (term_data):
    df = term_data

    # Initialise dataframes to store results
    busbar_voltage_frame = pd.DataFrame()

    # Organising Results
    busbar_voltage_frame['Tag Name'] = df['loc_name']
    busbar_voltage_frame['Rated Voltage (kV)'] = round(df['e:uknom'],2)
    busbar_voltage_frame['Voltage Level (kV)'] = round(df['m:Ul'], 3)
    busbar_voltage_frame['Voltage Limit'] = '+' + (((df['e:vmax'])-1)*100).astype(int).astype(str) + '%, ' + (((df['e:vmin'])-1)*100).astype(int).astype(str) + '%'
    condition = (df['m:u'] < df['e:vmax']) & (df['m:u'] > df['e:vmin'])
    busbar_voltage_frame.loc[condition, 'Assessment'] = 'Pass'
    busbar_voltage_frame.loc[~condition, 'Assessment'] = 'Fail'
    busbar_voltage_frame['Scenario'] = df['Scenario']

    return busbar_voltage_frame

def cable_util(line_data):
    df = line_data

    # Initialise dataframes to store results
    cable_util_frame = pd.DataFrame()

    # Organising Results
    cable_util_frame['Circuit ID'] = df['loc_name']
    cable_util_frame['Line Length (m)'] = round(df['e:dline']*1000)
    cable_util_frame['Rated Continuous Current (kA)'] = round(df['Inom'], 3)
    cable_util_frame['Maximum Current (kA)'] = round(df['m:I:bus2'], 3)
    cable_util_frame['Utilisation (%)'] = round(df['c:loading'], 1)
    cable_util_frame['Scenario'] = df['Scenario']

    return    cable_util_frame

def cb_util(cb_data):
    df = cb_data

    # Initialise dataframes to store results
    cb_util_frame  = pd.DataFrame()

    # Organising Results
    cb_util_frame['Tag Name'] = df['loc_name']
    cb_util_frame['Rated Current (kA)'] = round(df['Inom'], 3)
    cb_util_frame['Current (kA)'] = round(df['m:I:bus2'], 3)
    cb_util_frame['Circuit Breaker Loading (%)'] = round(df['m:Brkload:bus2']/df['Inom'], 1)
    cb_util_frame['Scenario'] = df['Scenario']

    return cb_util_frame

def tx_util(tx_data):
    
    df = tx_data

    # Initialise dataframes to store results
    tx_util_frame = pd.DataFrame()

    # Organising Results
    tx_util_frame['Tag Name'] = df['loc_name']
    tx_util_frame['Nominal L-L Voltage (kV)'] = round(df['t:utrn_h']).astype(str)+'/'+round(df['t:utrn_l'],2).astype(str)
    tx_util_frame['Rated Power (MVA)'] = df['Snom']
    tx_util_frame['Nominal Current (LV-Side) (kA)'] = round(df['Snom']/(df['t:utrn_l']*math.sqrt(3)), 3)
    tx_util_frame['Current Magnitude (LV-Side) (kA)'] = round(df['m:I:buslv'], 3)
    tx_util_frame['Utilisation (%)'] = round(df['c:loading'], 1)
    tx_util_frame['Scenario'] = df['Scenario']

    return tx_util_frame

def FaultLevelSummaryTable(_3psc_Max, _3psc_Min, _spgf_Max, _spgf_Min):

    # Initialise dataframes to store results
    FaultLevelSummary = pd.DataFrame()
    Max_3p = pd.DataFrame()
    Min_3p = pd.DataFrame()
    Max_1p = pd.DataFrame()
    Min_1p = pd.DataFrame()

    Max_3p = copy.deepcopy(_3psc_Max)
    Min_3p = copy.deepcopy(_3psc_Min)
    Max_1p = copy.deepcopy(_spgf_Max)
    Min_1p = copy.deepcopy(_spgf_Min)

    # Adding to each Data Frame for the merging process below 
    Max_3p['3P (Ikss) Max (kA)'] = Max_3p['m:Ikss']
    Min_3p['3P (Ikss) Min (kA)'] = Min_3p['m:Ikss']
    Max_1p['1PG (Ikss) Max (kA)'] = Max_1p['m:Ikss:A']
    Min_1p['1PG (Ikss) Min (kA)'] = Min_1p['m:Ikss:A']

    # Creating a Unique Filter Variable (This is needed when multipal scenarios are considered)
    Max_3p['Filter'] = Max_3p['Name']+Max_3p['Scenario']
    Min_3p['Filter'] = Min_3p['Name']+Min_3p['Scenario']
    Max_1p['Filter'] = Max_1p['Name']+Max_1p['Scenario']
    Min_1p['Filter'] = Min_1p['Name']+Min_1p['Scenario']

    # Organising Results
    FaultLevelSummary['Filter'] =  Max_3p['Filter']
    FaultLevelSummary['Name'] = Max_3p['Name']
    FaultLevelSummary['Nominal Voltage (kV)'] = Max_3p['e:uknom'].round(2)
    FaultLevelSummary['3P (Ikss) Max (kA)'] = Max_3p['3P (Ikss) Max (kA)']
    FaultLevelSummary = pd.merge(FaultLevelSummary, Min_3p[['Filter', '3P (Ikss) Min (kA)']], on='Filter', how='inner').round(2)
    FaultLevelSummary = pd.merge(FaultLevelSummary, Max_1p[['Filter', '1PG (Ikss) Max (kA)']], on='Filter', how='inner').round(2)
    FaultLevelSummary = pd.merge(FaultLevelSummary, Min_1p[['Filter', '1PG (Ikss) Min (kA)']], on='Filter', how='inner').round(2)
    FaultLevelSummary = pd.merge(FaultLevelSummary, Max_3p[['Filter', 'Scenario']], on='Filter', how='inner')
    
    #Deleting the 'Filter' column
    FaultLevelSummary.drop('Filter', axis=1, inplace=True) 

    return FaultLevelSummary

def SwitchboardShortCircuitThermalResultsTable(_Max3PSCData):
    df = _Max3PSCData

    # Initialise dataframes to store results
    SwitchboardShortCircuitThermalResults = pd.DataFrame()

    # Organising Results
    SwitchboardShortCircuitThermalResults['Tag Name'] = df['Name']
    condition = df['Ithlim'] > 0
    SwitchboardShortCircuitThermalResults.loc[condition,'Rated Short-Time Thermal Current(kA/Sec)'] = df['Ithlim'].astype(str)+'kA/'+df['Tkr'].astype(str)+'s'
    SwitchboardShortCircuitThermalResults.loc[~condition,'Rated Short-Time Thermal Current(kA/Sec)'] = 'NA'
    SwitchboardShortCircuitThermalResults['Thermal Equivalent Short-Circuit Current (kA)'] = df['m:Ith'].round(2)
    SwitchboardShortCircuitThermalResults['Rated Peak Withstand Current (kA)'] = df['Iplim']
    SwitchboardShortCircuitThermalResults['Peak Short Circuit Current (kA)'] = df['m:Ip'].round(2)
    SwitchboardShortCircuitThermalResults['3P Max (Ikss)']=df['m:Ikss'].round(2)
    

    condition = (df['Ithlim'] > df['m:Ith']) & (df['Iplim'] > df['m:Ip']) & (df['Ithlim'] > 0)
    SwitchboardShortCircuitThermalResults.loc[condition, 'Assessment'] = 'Pass'
    SwitchboardShortCircuitThermalResults.loc[~condition, 'Assessment'] = 'Fail'
    condition = df['Ithlim'] > 0
    SwitchboardShortCircuitThermalResults.loc[~condition, 'Assessment'] = 'Not Assessed'

    SwitchboardShortCircuitThermalResults['Scenario'] = df['Scenario']
    
    return SwitchboardShortCircuitThermalResults

