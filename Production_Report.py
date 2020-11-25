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
DC=pd.read_excel('DC-Data.xlsx',index_col=1,header=3)
DC.drop('Unnamed: 0',axis='columns', inplace=True)
DC=DC.fillna(0)
#########################
FN=pd.read_excel('FN-Data.xlsx',index_col=1,header=3)
FN.drop('Unnamed: 0',axis='columns', inplace=True)
FN=FN.fillna(0)
#########################
MC=pd.read_excel('MC-Data.xlsx',index_col=1,header=3)
MC.drop('Unnamed: 0',axis='columns', inplace=True)
MC=MC.fillna(0)
#########################
QC=pd.read_excel('QC-Data.xlsx')
QC=QC.fillna(0)
###################################
st.subheader('Production DC report')
DC
DC['Pct']=(DC['NG-Pcs']/DC['DC-Pcs'])*100
st.subheader('DC Summarize')
DCsum=DC[['DC-Pcs','BF-Pcs','NG-Pcs']].sum()
DCsum
DCpct=DC['Pct'].mean()
st.subheader('NG%')
st.warning(DCpct)
#############################
st.subheader('Production FN report')
FN
FN['Pct']=(FN['NG-Pcs']/(FN['BM-Pcs']+FN['FG0-Pcs']))*100
st.subheader('FN Summarize')
FNsum=FN[['BM-Pcs','FG0-Pcs','NG-Pcs']].sum()
FNsum
FNpct=FN['Pct'].mean()
st.subheader('NG%')
st.warning(FNpct)
#############################
st.subheader('Production MC report')
MC
MC['Pct']=(MC['NG-Pcs']/(MC['M-FG0']+MC['MC-Pcs']))*100
st.subheader('MC Summarize')
MCsum=MC[['MC-Pcs','M-FG0','NG-Pcs']].sum()
MCsum
MCpct=MC['Pct'].mean()
st.subheader('NG%')
st.warning(MCpct)
#############################
st.subheader('Production QC report')
QC

QC['Pct']=(QC['TT-NG-Pcs']/QC['Sorted-Pcs'])*100
st.subheader('QC Summarize')

QCsum=QC[['Sorted-Pcs','FG1-Pcs','TT-NG-Pcs']].sum()
QCsum
QCpct=QC['Pct'].mean()
st.subheader('NG%')
st.warning(QCpct)