import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import altair as alt

#st.set_page_config(layout="wide")

Logo=Image.open('SIM-Logo.jpeg')
st.image(Logo,width=500)
st.header('**Production Report**')
db=pd.read_excel('Database.xlsx','DB')
#########################
DC=pd.read_excel('DC-Data-Nov-20.xlsx',index_col=1,header=3)
DC.drop('Unnamed: 0',axis='columns', inplace=True)
DC=DC.fillna(0)
#########################
FN=pd.read_excel('FN-Data-Nov-20.xlsx',index_col=1,header=3)
FN.drop('Unnamed: 0',axis='columns', inplace=True)
FN=FN.fillna(0)
#########################
MC=pd.read_excel('MC-Data-Nov-20.xlsx',index_col=1,header=3)
MC.drop('Unnamed: 0',axis='columns', inplace=True)
MC=MC.fillna(0)
#########################
QC=pd.read_excel('QC-Data-Nov-20.xlsx')
QC=QC.fillna(0)
###################################

if st.checkbox('Production DC report'):
    st.subheader('Production DC report')
    DC

DC['Pct']=((DC['NG-Pcs'].sum()/DC['DC-Pcs'].sum())*100)
st.subheader('DC Summarize')
DCsum=DC[['DC-Pcs','BF-Pcs','NG-Pcs']].sum()
DCsum
DCpct=DC['Pct']

st.subheader('NG%')
DCpct=DCpct.mean()
st.warning(DCpct)

DCpph=DC['DC-Pcs']/(DC['Work-Hr']+DC['OT-Hr'])
DCpph=DCpph.groupby('PartNo').mean()

if st.checkbox('DC Performance Pcs/Hrs'):
    st.subheader('DC Performance Pcs/Hrs')
    st.table(DCpph)
    st.bar_chart(DCpph)
st.subheader('DC SUM Performance Pcs/Hrs')
DCpphm=DCpph.mean()
st.success(DCpphm)
DC['Pcsphr']=DC['DC-Pcs']/(DC['Work-Hr']+DC['OT-Hr'])
DChdmc=DC[['HDMC','Pcsphr']].groupby('HDMC').mean()

if st.checkbox('HDMC SUM Performance Pcs/Hrs'):
    st.subheader('HDMC SUM Performance Pcs/Hrs')
    st.bar_chart(DChdmc)

#############################

if st.checkbox('Production FN report'):
    st.subheader('Production FN report')
    FN
FN['Pct']=(FN['NG-Pcs'].sum()/(FN['BM-Pcs']+FN['FG0-Pcs']).sum())*100
st.subheader('FN Summarize')
FNsum=FN[['BM-Pcs','FG0-Pcs','NG-Pcs']].sum()
FNsum
FNpct=FN['Pct'].mean()
st.subheader('NG%')
st.warning(FNpct)

FNpph=(FN['BM-Pcs']+FN['FG0-Pcs'])/(FN['Work-Hr']+FN['OT-Hr'])
FNpph=FNpph.groupby('PartNo').mean()

if st.checkbox('FN Performance Pcs/Hrs'):
    st.subheader('FN Performance Pcs/Hrs')
    st.table(FNpph)
    st.bar_chart(FNpph)
st.subheader('FN SUM Performance Pcs/Hrs')
FNpphm=FNpph.mean()
st.success(FNpphm)
#############################

if st.checkbox('Production MC report'):
    st.subheader('Production MC report')
    MC
MC['Pct']=(MC['NG-Pcs'].sum()/MC['MC-Pcs'].sum())*100
st.subheader('MC Summarize')
MCsum=MC[['MC-Pcs','M-FG0','NG-Pcs']].sum()
MCsum
MCpct=MC['Pct'].mean()
st.subheader('NG%')
st.warning(MCpct)

MCpph=MC['MC-Pcs']/(MC['Work-Hr']+MC['OT-Hr'])
MCpph=MCpph.groupby('PartNo').mean()

if st.checkbox('MC Performance Pcs/Hrs'):
    st.subheader('MC Performance Pcs/Hrs')
    st.table(MCpph)
    st.bar_chart(MCpph)


MCpphm=MCpph.mean()
st.success(MCpphm)
#############################

if st.checkbox('Production QC report'):
    st.subheader('Production QC report')
    QC

QC['Pct']=(QC['TT-NG-Pcs'].sum()/QC['Sorted-Pcs'].sum())*100
st.subheader('QC Summarize')

QCsum=QC[['Sorted-Pcs','FG1-Pcs','TT-NG-Pcs']].sum()
QCsum
QCpct=QC['Pct'].mean()
st.subheader('NG%')
st.warning(QCpct)
