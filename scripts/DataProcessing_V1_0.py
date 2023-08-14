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
    tx_util_frame['Nominal L-L Voltage (kV)'] = round(df['t:utrn_h']).astype(str)+'/'+round(df['t:utrn_l'], 2).astype(str)
    tx_util_frame['Rated Power (MVA)'] = df['Snom']
    tx_util_frame['Nominal Current (LV-Side) (kA)'] = round(df['Snom']/(df['t:utrn_l']*math.sqrt(3)), 3)
    tx_util_frame['Current Magnitude (LV-Side) (kA)'] = round(df['m:I:buslv'], 3)
    tx_util_frame['Utilisation (%)'] = round(df['c:loading'], 1)
    tx_util_frame['Scenario'] = df['Scenario']

    return tx_util_frame

def fault_level_summary(max_3psc, min_3psc, max_spgf, min_spgf):

    # Initialise dataframes to store results
    fl_summary_frame = pd.DataFrame()
    max_3p = pd.DataFrame()
    min_3p = pd.DataFrame()
    max_1p = pd.DataFrame()
    min_1p = pd.DataFrame()

    max_3p = copy.deepcopy(max_3psc)
    min_3p = copy.deepcopy(min_3psc)
    max_1p = copy.deepcopy(max_spgf)
    min_1p = copy.deepcopy(min_spgf)

    # Adding to each Data Frame for the merging process below 
    max_3p['3P (Ikss) Max (kA)'] = max_3psc['m:Ikss']
    min_3p['3P (Ikss) Min (kA)'] = min_3psc['m:Ikss']
    max_1p['1PG (Ikss) Max (kA)'] = max_spgf['m:Ikss:A']
    min_1p['1PG (Ikss) Min (kA)'] = min_spgf['m:Ikss:A']

    # Creating a Unique Filter Variable (This is needed when multipal scenarios are considered)
    max_3p['Filter'] = max_3p['loc_name']+max_3p['Scenario']
    min_3p['Filter'] = min_3p['loc_name']+min_3p['Scenario']
    max_1p['Filter'] = max_1p['loc_name']+max_1p['Scenario']
    min_1p['Filter'] = min_1p['loc_name']+min_1p['Scenario']

    # Organising Results
    fl_summary_frame['Filter'] =  max_3p['Filter']
    fl_summary_frame['Name'] = max_3p['loc_name']
    fl_summary_frame['Nominal Voltage (kV)'] = round(max_3p['e:uknom'], 3)
    fl_summary_frame['3P (Ikss) Max (kA)'] = max_3p['3P (Ikss) Max (kA)']
    fl_summary_frame = round(pd.merge(fl_summary_frame, min_3p[['Filter', '3P (Ikss) Min (kA)']], on='Filter', how='inner'), 3)
    fl_summary_frame = round(pd.merge(fl_summary_frame, max_1p[['Filter', '1PG (Ikss) Max (kA)']], on='Filter', how='inner'), 3)
    fl_summary_frame = round(pd.merge(fl_summary_frame, min_1p[['Filter', '1PG (Ikss) Min (kA)']], on='Filter', how='inner'), 3)
    fl_summary_frame = round(pd.merge(fl_summary_frame, max_3p[['Filter', 'Scenario']], on='Filter', how='inner'), 3)
    
    #Deleting the 'Filter' column
    fl_summary_frame.drop('Filter', axis=1, inplace=True)

    return fl_summary_frame

def sb_sc_thermal(max_3psc_data):
    df = max_3psc_data

    # Initialise dataframes to store results
    sb_sc_thermal_frame = pd.DataFrame()

    # Organising Results
    sb_sc_thermal_frame['Tag Name'] = df['Name']
    condition = df['Ithlim'] > 0
    sb_sc_thermal_frame.loc[condition,'Rated Short-Time Thermal Current(kA/Sec)'] = df['Ithlim'].astype(str)+'kA/'+df['Tkr'].astype(str)+'s'
    sb_sc_thermal_frame.loc[~condition,'Rated Short-Time Thermal Current(kA/Sec)'] = 'NA'
    sb_sc_thermal_frame['Thermal Equivalent Short-Circuit Current (kA)'] = df['m:Ith'].round(2)
    sb_sc_thermal_frame['Rated Peak Withstand Current (kA)'] = df['Iplim']
    sb_sc_thermal_frame['Peak Short Circuit Current (kA)'] = df['m:Ip'].round(2)
    sb_sc_thermal_frame['3P Max (Ikss)']=df['m:Ikss'].round(2)
    

    condition = (df['Ithlim'] > df['m:Ith']) & (df['Iplim'] > df['m:Ip']) & (df['Ithlim'] > 0)
    sb_sc_thermal_frame.loc[condition, 'Assessment'] = 'Pass'
    sb_sc_thermal_frame.loc[~condition, 'Assessment'] = 'Fail'
    condition = df['Ithlim'] > 0
    sb_sc_thermal_frame.loc[~condition, 'Assessment'] = 'Not Assessed'

    sb_sc_thermal_frame['Scenario'] = df['Scenario']
    
    return sb_sc_thermal_frame
