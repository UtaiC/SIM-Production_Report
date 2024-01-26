######## Library ##########################################
from ast import Str
from email.mime import image
from email.policy import default
from operator import concat
from os import renames
from re import A
from select import select
from ssl import Options
from tkinter import Menu
from turtle import width
import streamlit as st
# from streamlit_option_menu import option_menu
import pandas as pd
from PIL import Image
import numpy as np
import os
from datetime import datetime, timedelta
import glob
from datetime import datetime, timedelta
import datetime
import calendar
############## Logo and Format ############################
Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=700)
#######################################################
st.markdown(
    """
    <div style="display: flex; justify-content: center;">
        <h1>Stock Movement Analysis 2024</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
####################### Format Data Frame ################################
def format_dataframe_columns(df):
    formatted_df = df.copy()  # Create a copy of the DataFrame
    for column in formatted_df.columns:
        if formatted_df[column].dtype == 'float64':  # Check if column has a numeric type
            formatted_df[column] = formatted_df[column].apply(lambda x: '{:,.0f}'.format(x))
    return formatted_df
#######################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
################################################################
################# MENU ####################################
Minput = st.sidebar.selectbox('Input-Month',['2024-01','2024-02','2024-03','2024-04','2024-05','2024-06',
                                             '2024-07','2024-08','2024-09','2024-10','2024-11','2024-12'])
Winput = st.sidebar.selectbox('Input-Week', [1,2,3,4,5])
Customer=st.sidebar.selectbox('Input-Customer', ['Valeo','TBKK','Electrolux','Homexpert','Koshin'])
Process = st.sidebar.selectbox('Input-Process',['DC','FN','SB','T5','MC','QC','Movement'])

#######################################################
def generate_weeks(year):
    start_date = datetime(year, 1, 1)
    start_date += timedelta(weeks=1)
    end_date = datetime(year, 12, 31)
    current_date = start_date
    weeks = []
    week_number = 1  # Start week number from 40

    while current_date <= end_date:
        week_start = current_date.strftime('%Y-%m-%d')
        week_end = (current_date + timedelta(days=6)).strftime('%Y-%m-%d')
        weeks.append((week_number, f"{week_start} - {week_end}"))
        current_date += timedelta(days=7)
        week_number += 1

    return weeks

def get_week_range_for_month(year, month):
    first_day = datetime.date(year, month, 1)
    last_day = datetime.date(year, month, calendar.monthrange(year, month)[1])

    start_week = first_day.isocalendar()[1]
    end_week = last_day.isocalendar()[1]

    if first_day.isocalendar()[0] < year:
        start_week = 1
    if last_day.isocalendar()[0] > year:
        end_week = 53

    return start_week, end_week
##################### New Month +1 #####################################
Minput = Minput
year, month = map(int, Minput.split('-'))
start_week, end_week = get_week_range_for_month(year, month)
########################## Read File #####################
st.write('---')
############## Read Files ##############
Cust=pd.read_excel('Customer.xlsx')
# Function to load dataframes from the Excel file
@st.cache_data
def load_dataframes(sheet_names):
    # url = "https://docs.google.com/spreadsheets/d/1QBYKcKy5feKnWfth4L1XfnbUx9EFjg_L/export?format=xlsx"
    file="Production-2024.xlsx"
    dataframes = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(file, header=7, engine='openpyxl', sheet_name=sheet_name)
        dataframes[sheet_name] = df
    return dataframes
all_sheet_names = [str(week) for week in range(start_week,end_week+1)]
dataframes = load_dataframes(all_sheet_names)
dataframes = {int(week): df for week, df in dataframes.items()}
# ############## Function Month #######################################
#     dataframes = {}
#     for sheet_name in sheet_names:
#         df = pd.read_excel(file, header=7, engine='openpyxl', sheet_name=sheet_name)
#         dataframes[sheet_name] = df
#     return dataframes
############## Range Month #################################
# all_sheet_names = [str(week) for week in range(start_week,end_week+1)]
# dataframes = load_dataframes(all_sheet_names)
# dataframes = {int(week): df for week, df in dataframes.items()}
#################### Weekly Data ##########################
DataMerges=dataframes[Winput]
EndCheck=dataframes[Winput+1]
EndCheck=EndCheck[['Part no.','Beginning Balance','Beginning Balance.1','Beginning Balance.2','Beginning Balance.3',
                  'Beginning Balance.4','Beginning Balance.5','Beginning Balance.6']]
EndCheck=EndCheck.rename(columns={'Part no.':'Part_No',
                                  'Beginning Balance':'DCEnd',
                                  'Beginning Balance.1':'FNEnd-1',
                                  'Beginning Balance.2':'FNEnd-2',
                                  'Beginning Balance.3':'SBEnd',
                                  'Beginning Balance.4':'T5End',
                                  'Beginning Balance.5':'MCEnd',
                                  'Beginning Balance.6':'QCEnd'})
EndCheck['FNEnd']=EndCheck['FNEnd-1']+EndCheck['FNEnd-2']
################ Maergs Sheet ##################################
# sheet_numbers_to_merge = range(start_week, end_week+1) 
# dfs_to_merge = [dataframes[num] for num in sheet_numbers_to_merge]
# # Merge the selected dataframes into a single dataframe
# merged_df = pd.concat(dfs_to_merge, ignore_index=True)  # Concatenate and reset index
# DataMerges=merged_df
# DataMerges=DataMerges.fillna(0)
DataMerges=DataMerges.rename(columns={'Part no.':'Part_No','Itemes':'Part_No'})
values_to_exclude = ['nan', 'TBKK','Itemes','KOSHIN', 'ELECTROLUX', 'HOME EXPERT','Home Expert',0]
mask = ~DataMerges['Part_No'].isin(values_to_exclude)
DataMerges = DataMerges[mask]
# DataMerges
DataMerges['Part_No']=DataMerges['Part_No'].astype(str)
DataMerges=pd.merge(DataMerges,EndCheck,on='Part_No',how='left')
DataMerges=pd.merge(DataMerges,Cust,on='Part_No',how='outer')
DataMerges['Customer']=DataMerges['Customer'].fillna('NoN')
DataMerges=DataMerges[DataMerges['Customer'].str.contains(Customer)]
# DataMerges=DataMerges[~DataMerges['Machine'].str.contains('KAYAMA')]  
# DataMerges  
##################### DC ##################################
DCData=DataMerges
DCData.columns = [str(col) for col in DCData.columns]
boolean_mask = DCData.columns.str.startswith('2024')
Date_columns = DCData.loc[:, boolean_mask]
boolean_mask = DCData.columns.str.endswith(':00')
DC_Date_columns = DCData.loc[:, boolean_mask]
# Convert column names to strings explicitly
DC_Date_columns.columns = DC_Date_columns.columns.astype(str) 
DCData['Beginning Balance'] = pd.to_numeric(DCData['Beginning Balance'], errors='coerce')
DCData['Total'] = pd.to_numeric(DCData['Total'], errors='coerce')
DCData['DCEnd'] = pd.to_numeric(DCData['DCEnd'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=DCData['Total']-DCData['Beginning Balance']
DCData['Prod-Pcs'] = SUMPcs
DCData=DCData[['Part_No','Prod-Pcs','Total','Beginning Balance','DCEnd']+ DC_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in DCData
selected_columns = ['Part_No','Prod-Pcs','Total','Beginning Balance','DCEnd'] + DC_Date_columns.columns.tolist()
DCData = DCData[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance': 'first',
    'DCEnd': 'first',
    'Prod-Pcs': 'sum',
    'Total': 'sum'
}
# Iterate through columns in DC_Date_columns and add them to aggregation_functions
for col in DC_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on DCData
DCData = DCData.groupby('Part_No').agg(aggregation_functions).reset_index()
######### Display #########################
DCData['Prod-Pcs']=DCData['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
SUMBeg=DCData['Beginning Balance'].sum()
DCData['DC-Move']=(DCData['Beginning Balance']+DCData['Prod-Pcs'])-DCData['DCEnd']
DCPcs=DCData['Prod-Pcs'].sum()
SUMTT=SUMBeg+DCPcs
SUMEnd=DCData['DCEnd'].sum()
DCMove=SUMTT-SUMEnd
if Process=='DC':  
    st.write(' DC analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    DCData=DCData.fillna(0)
    DCData
    
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('Production',round(DCPcs),'Pcs')
    formatted_display('Total Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(DCMove),'Pcs')
##################### FN ##################################
FNData=DataMerges
FNData.columns = [str(col) for col in FNData.columns]
boolean_mask = FNData.columns.str.startswith('2024')
Date_columns = FNData.loc[:, boolean_mask]
boolean_mask = FNData.columns.str.endswith(':00.1')|FNData.columns.str.endswith(':00.2')
FN_Date_columns = FNData.loc[:, boolean_mask]
# Convert column names to strings explicitly
FN_Date_columns.columns = FN_Date_columns.columns.astype(str) 
FNData['Beginning Balance.1'] = pd.to_numeric(FNData['Beginning Balance.1'], errors='coerce')
FNData['Beginning Balance.2'] = pd.to_numeric(FNData['Beginning Balance.2'], errors='coerce')
FNData['Total 1'] = pd.to_numeric(FNData['Total 1'], errors='coerce')
FNData['Total 2'] = pd.to_numeric(FNData['Total 2'], errors='coerce')
FNData['FNEnd'] = pd.to_numeric(FNData['FNEnd'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=(FNData['Total 1']+FNData['Total 2'])-(FNData['Beginning Balance.1']+FNData['Beginning Balance.2'])
FNData['Prod-Pcs'] = SUMPcs
FNData['FN-Move']=(FNData['Total 1']+FNData['Total 2'])-FNData['FNEnd']
# FNData[['Part_No','FN-Move']]
FNData=FNData[['Part_No','Prod-Pcs','Total 1','Total 2','Beginning Balance.1','Beginning Balance.2','FNEnd',]+ FN_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in FNData
selected_columns = ['Part_No','Prod-Pcs','Total 1','Total 2','Beginning Balance.1','Beginning Balance.2','FNEnd',] + FN_Date_columns.columns.tolist()
FNData = FNData[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance.1': 'first',
    'Beginning Balance.2': 'first',
    'FNEnd': 'first',
    'Prod-Pcs': 'sum',
    'Total 1': 'sum',
    'Total 2': 'sum'
}
# Iterate through columns in FN_Date_columns and add them to aggregation_functions
for col in FN_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on FNData
FNData = FNData.groupby('Part_No').agg(aggregation_functions).reset_index()
######### Display #########################
SUMBeg=(FNData['Beginning Balance.1'].sum()+FNData['Beginning Balance.2'].sum())
FNData['Prod-Pcs']=FNData['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
SUMEnd=FNData['FNEnd'].sum()
FNPcs=FNData['Prod-Pcs'].sum()
SUMTT=SUMBeg+FNPcs
FNMove=SUMTT-SUMEnd

if Process=='FN':  
    st.write(' FN analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    FNData=FNData.fillna(0)
    FNData
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('Production',round(FNPcs),'Pcs')
    formatted_display('Total Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(FNMove),'Pcs')
##################### SB ##################################
SBData=DataMerges
SBData.columns = [str(col) for col in SBData.columns]
boolean_mask = SBData.columns.str.startswith('2024')
Date_columns = SBData.loc[:, boolean_mask]
boolean_mask = SBData.columns.str.endswith(':00.3')
SB_Date_columns = SBData.loc[:, boolean_mask]
# Convert column names to strings explicitly
SB_Date_columns.columns = SB_Date_columns.columns.astype(str) 
SBData['Beginning Balance.3'] = pd.to_numeric(SBData['Beginning Balance.3'], errors='coerce')
SBData['Total.1'] = pd.to_numeric(SBData['Total.1'], errors='coerce')
SBData['SBEnd'] = pd.to_numeric(SBData['SBEnd'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=SBData['Total.1']-SBData['Beginning Balance.3']
SBData['Prod-Pcs'] = SUMPcs
SBData=SBData[['Part_No','Prod-Pcs','Total.1','Beginning Balance.3','SBEnd','Type']+ SB_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in SBData
selected_columns = ['Part_No','Prod-Pcs','Total.1','Beginning Balance.3','SBEnd','Type'] + SB_Date_columns.columns.tolist()
SBData = SBData[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance.3': 'first',
    'SBEnd': 'first',
    'Type': 'first',
    'Prod-Pcs': 'sum',
    'Total.1': 'sum'
}
# Iterate through columns in SB_Date_columns and add them to aggregation_functions
for col in SB_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on SBData
SBData = SBData.groupby('Part_No').agg(aggregation_functions).reset_index()
    ######### Display #########################
SUMBeg=SBData['Beginning Balance.3'].sum()
SBData['Prod-Pcs']=SBData['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
SBData['SB-TT-RM']=SBData['Total.1'].where(SBData['Type'].str.contains('RM'))
SBData['SB-RM-Move']=SBData['SB-TT-RM']-SBData['SBEnd']
SBData['SB-TT-MC']=SBData['Total.1'].where(SBData['Type'].str.contains('MC'))
SBData['SB-MC-Move']=SBData['SB-TT-MC']-SBData['SBEnd']
SBData['SB-RM-Move']=SBData['SB-TT-RM']-SBData['SBEnd']
SBData['SB-MC-Move']=SBData['SB-MC-Move'].fillna(0)
SBData['SB-RM-Move']=SBData['SB-RM-Move'].fillna(0)
SBMCSUM=SBData['SB-MC-Move'].sum()
SBRMSUM=SBData['SB-RM-Move'].sum()
# SBData[['Part_No','SB-MC-Move']]
# SBData[['Part_No','SB-RM-Move']]
SBPcs=SBData['Prod-Pcs'].sum()
SUMEnd=SBData['SBEnd'].sum()
SUMTT=SUMBeg+SBPcs
SBMove=SUMTT-SUMEnd


if Process=='SB': 
    st.write(' SB analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    SBData=SBData.fillna(0)
    SBData
    
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('Production',round(SBPcs),'Pcs')
    formatted_display('Total.1 Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(SBMove),'Pcs')
##################### T5 ##################################
T5Data=DataMerges
T5Data.columns = [str(col) for col in T5Data.columns]
boolean_mask = T5Data.columns.str.startswith('2024')
Date_columns = T5Data.loc[:, boolean_mask]
boolean_mask = T5Data.columns.str.endswith(':00.4')
T5_Date_columns = T5Data.loc[:, boolean_mask]
# Convert column names to strings explicitly
T5_Date_columns.columns = T5_Date_columns.columns.astype(str) 
T5Data['Beginning Balance.4'] = pd.to_numeric(T5Data['Beginning Balance.4'], errors='coerce')
T5Data['Total.2'] = pd.to_numeric(T5Data['Total.2'], errors='coerce')
T5Data['MC OP1 / T5'] = pd.to_numeric(T5Data['MC OP1 / T5'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=T5Data['Total.2']-T5Data['Beginning Balance.4']
T5Data['Prod-Pcs'] = SUMPcs
T5Data=T5Data[['Part_No','Prod-Pcs','Total.2','Beginning Balance.4','MC OP1 / T5']+ T5_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in T5Data
selected_columns = ['Part_No','Prod-Pcs','Total.2','Beginning Balance.4','MC OP1 / T5'] + T5_Date_columns.columns.tolist()
T5Data = T5Data[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance.4': 'first',
    'MC OP1 / T5': 'first',
    'Prod-Pcs': 'sum',
    'Total.2': 'sum'
}
# Iterate through columns in T5_Date_columns and add them to aggregation_functions
for col in T5_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on T5Data
T5Data = T5Data.groupby('Part_No').agg(aggregation_functions).reset_index()
    ######### Display #########################
SUMBeg=T5Data['Beginning Balance.4'].sum()
T5Data['Prod-Pcs']=T5Data['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
T5Pcs=T5Data['Prod-Pcs'].sum()
SUMEnd=T5Data['MC OP1 / T5'].sum()
SUMTT=SUMBeg+T5Pcs
T5Move=SUMTT-SUMEnd
if Process=='T5':  
    st.write(' T5 analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    T5Data=T5Data.fillna(0)
    T5Data
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('Production',round(T5Pcs),'Pcs')
    formatted_display('Total.2 Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(T5Move),'Pcs')
##################### MC ##################################
MCData=DataMerges
MCData.columns = [str(col) for col in MCData.columns]
boolean_mask = MCData.columns.str.startswith('2024')
Date_columns = MCData.loc[:, boolean_mask]
boolean_mask = MCData.columns.str.endswith(':00.5')
MC_Date_columns = MCData.loc[:, boolean_mask]
# Convert column names to strings explicitly
MC_Date_columns.columns = MC_Date_columns.columns.astype(str) 
MCData['Beginning Balance.5'] = pd.to_numeric(MCData['Beginning Balance.5'], errors='coerce')
MCData['Total.3'] = pd.to_numeric(MCData['Total.3'], errors='coerce')
MCData['MCEnd'] = pd.to_numeric(MCData['MCEnd'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=MCData['Total.3']-MCData['Beginning Balance.5']
MCData['Prod-Pcs'] = SUMPcs
MCData=MCData[['Part_No','Prod-Pcs','Total.3','Beginning Balance.5','MCEnd']+ MC_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in MCData
selected_columns = ['Part_No','Prod-Pcs','Total.3','Beginning Balance.5','MCEnd'] + MC_Date_columns.columns.tolist()
MCData = MCData[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance.5': 'first',
    'MCEnd': 'first',
    'Prod-Pcs': 'sum',
    'Total.3': 'sum'
}
# Iterate through columns in MC_Date_columns and add them to aggregation_functions
for col in MC_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on MCData
MCData = MCData.groupby('Part_No').agg(aggregation_functions).reset_index()
    ######### Display #########################
SUMBeg=MCData['Beginning Balance.5'].sum()
MCData['Prod-Pcs']=MCData['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
MCPcs=MCData['Prod-Pcs'].sum()
SUMTT=SUMBeg+MCPcs
SUMEnd=MCData['MCEnd'].sum()
MCMove=SUMTT-SUMEnd
if Process=='MC':  
    st.write(' MC analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    MCData=MCData.fillna(0)
    MCData
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('Production',round(MCPcs),'Pcs')
    formatted_display('Total.3 Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(MCMove),'Pcs')
##################### QC ##################################
QCData=DataMerges
QCData.columns = [str(col) for col in QCData.columns]
boolean_mask = QCData.columns.str.startswith('2024')
Date_columns = QCData.loc[:, boolean_mask]
boolean_mask = QCData.columns.str.endswith(':00.6')
QC_Date_columns = QCData.loc[:, boolean_mask]
# Convert column names to strings explicitly
QC_Date_columns.columns = QC_Date_columns.columns.astype(str) 
QCData['Beginning Balance.6'] = pd.to_numeric(QCData['Beginning Balance.6'], errors='coerce')
QCData['Total.4'] = pd.to_numeric(QCData['Total.4'], errors='coerce')
QCData['QCEnd'] = pd.to_numeric(QCData['QCEnd'], errors='coerce')
Date_columns = Date_columns.apply(pd.to_numeric, errors='coerce')
##########################
SUMPcs=QCData['Total.4']-QCData['Beginning Balance.6']
QCData['Prod-Pcs'] = SUMPcs
QCData=QCData[['Part_No','Prod-Pcs','Total.4','Beginning Balance.6','QCEnd','Type','ACT.7']+ QC_Date_columns.columns.tolist()]
###############################
# Selecting specific columns in QCData
selected_columns = ['Part_No','Prod-Pcs','Total.4','Beginning Balance.6','QCEnd','Type','ACT.7'] + QC_Date_columns.columns.tolist()
QCData = QCData[selected_columns]
# Define the aggregation functions for each column
aggregation_functions = {
    'Beginning Balance.6': 'first',
    'QCEnd': 'first',
    'Prod-Pcs': 'sum',
    'Total.4': 'sum',
    'Type':'first',
    'ACT.7':'sum'
}
# Iterate through columns in QC_Date_columns and add them to aggregation_functions
for col in QC_Date_columns.columns:
    aggregation_functions[col] = 'sum'
# Perform groupby aggregation on QCData
QCData = QCData.groupby('Part_No').agg(aggregation_functions).reset_index()
    ######### Display #########################
SUMBeg=QCData['Beginning Balance.6'].sum()
QCData['Prod-Pcs']=QCData['Prod-Pcs'].apply(pd.to_numeric, errors='coerce')
# QCData['Prod-Pcs']=QCData['Prod-Pcs'].fillna(0)
QCPcs=QCData['Prod-Pcs'].sum()
SUMTT=QCData['Total.4'].sum()
SUMEnd=QCData['QCEnd'].sum()
QCMove=SUMTT-SUMEnd
QCRM=QCData[QCData['Type'].str.contains('RM')]
# QCRM[['Part_No','Prod-Pcs']]
QCRM=QCRM['Prod-Pcs'].sum()
QCMC=QCData[QCData['Type'].str.contains('MC')]
# QCMC[['Part_No','Prod-Pcs']]
QCMC=QCMC['Prod-Pcs'].sum()
if Process=='QC':  
    st.write(' QC analysis @', Minput)
    st.write('Week Data: Week@',Winput)
    QCData=QCData.fillna(0)
    QCData
    formatted_display('Begining Stock',round(SUMBeg),'Pcs')
    formatted_display('MC-QC Pcs',round(QCMC),'Pcs')
    formatted_display('RM-QC Pcs',round(QCRM),'Pcs')
    formatted_display('TT-Production',round(QCPcs),'Pcs')
    formatted_display('Total.4 Stock',round(SUMTT),'Pcs')
    formatted_display('Ending Stock',round(SUMEnd),'Pcs')
    formatted_display('Movement Stock',round(QCMove),'Pcs')
################## Gap Movement ##########################
if Process=='Movement':
    st.subheader('Movement Gap Analysis')
    st.write('Weekly Data Summarize @week',Winput)
    # st.write('---')
############### DC #########################
    FNProd=FNData[['Part_No','Prod-Pcs']]
    DCMove=DCData[['Part_No','DC-Move']]
    DCGapCheck=pd.merge(FNProd,DCMove,on='Part_No',how='left')
    DCGapCheck['BF-Gap']=DCGapCheck['Prod-Pcs']-DCGapCheck['DC-Move']
    # st.write('---')
    st.write('Gap DC Details')
    DCGapCheck=DCGapCheck[(DCGapCheck['DC-Move']!=0) | (DCGapCheck['Prod-Pcs']!=0)].reset_index(drop=True)
    DCGapCheck.index=DCGapCheck.index+1
    ###################
    formatted_df = format_dataframe_columns(DCGapCheck)
    st.dataframe(formatted_df)
    ####################
    DCSUM=DCGapCheck['BF-Gap'].sum()
    DCSUM
    ########################
    DCMove=DCGapCheck['DC-Move'].sum()
    BFGap=DCGapCheck['BF-Gap'].sum()
    ######################## Table-BF #########################################
    st.write('BF Gap/Loss')
    DCSUMDATA = {
        'Operation': ['DC Movement','FN Production','DC to FN GAP/loss'],
        'Type': ['Pcs','Pcs','Pcs'],
        'Quantity': [DCMove,FNPcs,BFGap]}
    DCSUMDATA  = pd.DataFrame(DCSUMDATA)
    DCSUMDATA .index += 1
    DCSUMDATA['Quantity'] = DCSUMDATA['Quantity'].map('{:,.0f}'.format)
    st.table(DCSUMDATA )
    ################# To Excel ############################
    name='DC-Movement @ Week-'
    week=str(Winput)
    filename=name+week+'.xlsx'
    path=r'C:\Users\utaie\Desktop\Costing\Movement-2023\Weekly-Movement-2024\\'
    DCGapCheck.to_excel(path+filename)
################# FN ##################
    FNData['FN-Move']=(FNData['Total 1']+FNData['Total 2'])-FNData['FNEnd']
    SBProd=SBData[['Part_No','Prod-Pcs']]
    FNMove=FNData[['Part_No','FN-Move']]
    FNGapCheck=pd.merge(SBProd,FNMove,on='Part_No',how='left')
    FNGapCheck['FN-Gap']=FNGapCheck['Prod-Pcs']-FNGapCheck['FN-Move']
    st.write('---')
    st.write('Gap FN Details')
    FNGapCheck=FNGapCheck[(FNGapCheck['FN-Move']!=0) | (FNGapCheck['Prod-Pcs']!=0)].reset_index(drop=True)
    FNGapCheck=FNGapCheck.fillna(0)
    FNGapCheck.index=FNGapCheck.index+1
    ###################
    formatted_df = format_dataframe_columns(FNGapCheck)
    st.dataframe(formatted_df)
    ####################
    FNSUM=FNGapCheck['FN-Gap'].sum()
    FNSUM
    #######################
    FNMove=FNGapCheck['FN-Move'].sum()
    FNGap=FNSUM
    ###################### Teble FN ###############
    st.write('BS Gap/Loss')
    FNSUMDATA = {
        'Operation': ['FN Movement','SB Production','FN to SB GAP/loss'],
        'Type': ['Pcs','Pcs','Pcs'],
        'Quantity': [FNMove,SBPcs,FNGap]}
    FNSUMDATA  = pd.DataFrame(FNSUMDATA)
    FNSUMDATA .index += 1
    FNSUMDATA['Quantity'] = FNSUMDATA['Quantity'].map('{:,.0f}'.format)
    st.table(FNSUMDATA )
    st.write('BM Gap/Loss')
    
    ################# To Excel ############################
    name='FN-Movement @ Week-'
    week=str(Winput)
    filename=name+week+'.xlsx'
    path=r'C:\Users\utaie\Desktop\Costing\Movement-2023\Weekly-Movement-2024\\'
    FNGapCheck.to_excel(path+filename)
################ SB ###################################
    SBData['SB-Move']=(SBData['Total.1']-SBData['SBEnd'])
    MCData['MC-Move']=(MCData['Total.3']-MCData['MCEnd'])
    QCRM=QCData[QCData['Type'].str.contains('RM')]
    QCMC=QCData[QCData['Type'].str.contains('MC')]
    MCProd=MCData[['Part_No','Prod-Pcs']]
    QCRM=QCRM[['Part_No','Prod-Pcs']]
    QCRM.rename(columns={'Prod-Pcs':'QCRM-Pcs'},inplace=True)
    SBMove=SBData[['Part_No','SB-Move']]
    SBGapCheck=pd.merge(MCProd,SBMove,on='Part_No',how='outer')
    SBGapCheck=pd.merge(QCRM,SBGapCheck,on='Part_No',how='outer')
    SBGapCheck=pd.merge(Cust,SBGapCheck,on='Part_No',how='outer')
    SBGapCheck['MC']=SBGapCheck['MC'].fillna('NoN')
    SBMoveT5=SBGapCheck[SBGapCheck['MC'].str.contains('KAYAMA')]
    SBMoveT5=SBMoveT5[['Part_No','SB-Move']]
    # SBMoveT5
    SBMoveT5.rename(columns={'SB-Move':'SB-Move-T5'},inplace=True)
    SBMoveRM=SBGapCheck[SBGapCheck['Type'].str.contains('RM')]
    SBMoveRM=SBMoveRM[['Part_No','SB-Move']]
    SBMoveRM.rename(columns={'SB-Move':'SB-Move-RM'},inplace=True)
    SBMoveRM=SBMoveRM.fillna(0)
    SBGapCheck['MC']=SBGapCheck['MC'].fillna('NoN')
    SBMoveMC=SBGapCheck[SBGapCheck['Type'].str.contains('MC')].where(SBGapCheck['MC']!=('KAYAMA'))
    SBMoveMC=SBMoveMC[['Part_No','SB-Move']]
    SBMoveMC.rename(columns={'SB-Move':'SB-Move-MC'},inplace=True)
    SBMoveMC=SBMoveMC.fillna(0)
    T5Prod=T5Data[['Part_No','Prod-Pcs']]
    T5Prod.rename(columns={'Prod-Pcs':'Prod-T5-Pcs'},inplace=True)
    SBGapCheck=pd.merge(SBMoveMC,SBGapCheck,on='Part_No',how='right')
    SBGapCheck=pd.merge(SBMoveRM,SBGapCheck,on='Part_No',how='right')
    SBGapCheck=pd.merge(SBMoveT5,SBGapCheck,on='Part_No',how='right')
    SBGapCheck=pd.merge(T5Prod,SBGapCheck,on='Part_No',how='right')
    SBGapCheck=SBGapCheck.fillna(0)
    #SUMALL['QC-Over-MC']=SUMALL['QC-Over-Pcs'].where(SUMALL['Part-Type'].str.contains('MC'))
    SBGapCheck['SB-Gap-T5']=SBGapCheck['Prod-T5-Pcs']-SBGapCheck['SB-Move-T5']
    SBGapCheck['SB-Gap-MC']=SBGapCheck['Prod-Pcs'].where(SBGapCheck['MC']!=('KAYAMA'))-SBGapCheck['SB-Move-MC']
    SBGapCheck['SB-Gap-RM']=SBGapCheck['QCRM-Pcs']-SBGapCheck['SB-Move-RM']
    st.write('Gap SB Details')
    # SBGapCheck=SBGapCheck[SBGapCheck['SB-Move']>0]
    SBGapCheck=SBGapCheck.fillna(0)
    SBGapCheck=SBGapCheck[(SBGapCheck['SB-Move']!=0) |(SBGapCheck['Prod-Pcs']!=0)|(SBGapCheck['QCRM-Pcs']!=0)].reset_index(drop=True)
    SBGapCheck.rename(columns={'Prod-Pcs':'Prod-MC-Pcs'},inplace=True)
    SBGapCheck['SB-TT-Gap']=SBGapCheck['SB-Gap-MC']+SBGapCheck['SB-Gap-RM']
    SBGapCheck.index=SBGapCheck.index+1
    
    SBGapCheck=SBGapCheck[['Part_No','SB-Move-T5','SB-Move-MC','SB-Move-RM','Prod-T5-Pcs','Prod-MC-Pcs','QCRM-Pcs','SB-Gap-T5','SB-Gap-MC','SB-Gap-RM','SB-TT-Gap']]
    
    ###################
    formatted_df = format_dataframe_columns(SBGapCheck)
    st.dataframe(formatted_df)
    ####################
    SBSUM=SBGapCheck['SB-TT-Gap'].sum()
    SBSUM
    # SBGapCheck
    ####################################
    SBMove=(SBGapCheck['SB-Move-T5'].sum())+(SBGapCheck['SB-Move-MC'].sum())+(SBGapCheck['SB-Move-RM'].sum())
    SBGap=SBSUM
    ############### Teble BS #################
    st.write('BS Summariz')
    SBSUMDATA = {
        'Operation': ['SB Movement','T5 Production','MC Production','QC Production','SB GAP/loss'],
        'Type': ['Pcs','Pcs','Pcs','Pcs','Pcs'],
        'Quantity': [SBMove,T5Pcs,MCPcs,QCPcs,SBGap]}
    SBSUMDATA  = pd.DataFrame(SBSUMDATA)
    SBSUMDATA .index += 1
    SBSUMDATA['Quantity'] = SBSUMDATA['Quantity'].map('{:,.0f}'.format)
    st.table(SBSUMDATA )
    st.write('BM Gap/Loss')
    
    ################# To Excel ############################
    name='SB-Movement @ Week-'
    week=str(Winput)
    filename=name+week+'.xlsx'
    path=r'C:\Users\utaie\Desktop\Costing\Movement-2023\Weekly-Movement-2024\\'
    SBGapCheck.to_excel(path+filename)
################# MC ####################    
    QCMC=QCMC[['Part_No','Prod-Pcs']]
    QCMC.rename(columns={'Prod-Pcs':'QCMC-Pcs'},inplace=True)
    MCMove=MCData[['Part_No','MC-Move']]
    MCGapCheck=pd.merge(QCMC,MCMove,on='Part_No',how='outer')
    MCGapCheck=MCGapCheck.fillna(0)
    MCGapCheck=MCGapCheck[(MCGapCheck['MC-Move']!=0)|(MCGapCheck['QCMC-Pcs']!=0)].reset_index(drop=True)
    MCGapCheck['MC-Gap']=MCGapCheck['QCMC-Pcs']-MCGapCheck['MC-Move']
    st.write('Gap MC Details')
    MCGapCheck.index=MCGapCheck.index+1
    ###################
    formatted_df = format_dataframe_columns(MCGapCheck)
    st.dataframe(formatted_df)
    ###################
    MCSUM=MCGapCheck['MC-Gap'].sum()
    MCSUM
    ##########################
    MCMove=MCGapCheck['MC-Move'].sum()
    ###################### Teble MC ###############
    st.write('FG0 Summariz')
    MCSUMDATA = {
        'Operation': ['MC Movement','QC Production','MC GAP/loss'],
        'Type': ['Pcs','Pcs','Pcs'],
        'Quantity': [MCMove,QCPcs,MCSUM]}
    MCSUMDATA  = pd.DataFrame(MCSUMDATA)
    MCSUMDATA .index += 1
    MCSUMDATA['Quantity'] = MCSUMDATA['Quantity'].map('{:,.0f}'.format)
    st.table(MCSUMDATA )
    ################# To Excel ############################
    name='MC-Movement @ Week-'
    week=str(Winput)
    filename=name+week+'.xlsx'
    path=r'C:\Users\utaie\Desktop\Costing\Movement-2023\Weekly-Movement-2024\\'
    MCGapCheck.to_excel(path+filename)
################# QC #####################################
    QCData['QC-Move']=QCData['Total.4']-QCData['QCEnd']   
    Sales=QCData[['Part_No','ACT.7']]
    Sales.rename(columns={'ACT.7':'Sales-Pcs'},inplace=True)
    QCMove=QCData[['Part_No','QC-Move']]
    QCGapCheck=pd.merge(Sales,QCMove,on='Part_No',how='outer')
    QCGapCheck=QCGapCheck.fillna(0)
    QCGapCheck=QCGapCheck[(QCGapCheck['QC-Move']!=0)|(QCGapCheck['Sales-Pcs']!=0)].reset_index(drop=True)
    QCGapCheck['QC-Gap']=QCGapCheck['Sales-Pcs']-QCGapCheck['QC-Move']
    st.write('Gap QC Details')
    # QCGapCheck
    QCGapCheck.index=QCGapCheck.index+1
    QCGapCheck
    QCSUM=QCGapCheck['QC-Gap'].sum()
    QCSUM
    ########################
    QCMove=QCGapCheck['QC-Move'].sum()
    ################## Table BF ####################################
    st.write('FG Summariz')
    SalePcs=QCData['ACT.7'].sum()
    QCGap=SalePcs-QCMove
    QCSUMDATA = {
        'Operation': ['QC Movement','Sales-Pcs','FG GAP/loss'],
        'Type': ['Pcs','Pcs','Pcs'],
        'Quantity': [QCMove,SalePcs,QCGap]}
    QCSUMDATA  = pd.DataFrame(QCSUMDATA)
    QCSUMDATA .index += 1
    QCSUMDATA['Quantity'] = QCSUMDATA['Quantity'].map('{:,.0f}'.format)
    st.table(QCSUMDATA )
    st.write('---')
    ################# To Excel ############################
    name='QC-Movement @ Week-'
    week=str(Winput)
    filename=name+week+'.xlsx'
    path=r'C:\Users\utaie\Desktop\Costing\Movement-2023\Weekly-Movement-2024\\'
    QCGapCheck.to_excel(path+filename)
