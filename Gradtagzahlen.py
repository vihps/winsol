'''
Created on 12.11.2017

@author: info_000
'''
#import requests
#import csv
import pandas as pd
#import numpy as np
from datetime import timedelta,date,datetime
#import IPython
from pandas.tests.io.parser import skiprows
#nur hier Eingabe
monat="1712" 
pfad='D:'+r'\Winsol'+r'\Excel'
datei=pfad+r'\E'+str(monat)+'.xls'
apfad='D:'+r'\GTZ'
adatei=apfad+r'\gtz'+str(monat)+'.xls'
#print(datei)
#Raumtemperatur
rt=20
#Heizgrenztemperatur
hg=15
#Spaltenname als string
gradtag="G"+str(rt)+"/"+str(hg)
hgtag='HGT'+str(hg)

dat = pd.read_excel(open(datei,'rb'), sheetname=0, usecols=[0,12],header=None, skiprows=1,names=['Zeitpunkt','ta'])

dat.Zeitpunkt = [datetime.strptime(str(i),"%Y-%m-%d %H:%M:%S") for i in dat.Zeitpunkt]

dates= [Zeitpunkt.date() for Zeitpunkt in dat.Zeitpunkt]

#use set() to obtain a unique list of dates
uniqueDates = list(set(dates)) 
uniqueDates.sort() 

countsAndDays      = [dates.count(ud) for ud in uniqueDates]
idx4Dates = [[date[0] for date in enumerate(dates) if date[1] == i] for i in uniqueDates]
averages = [dat.ta[idx4Dates[i[0]]].mean() for i in enumerate(uniqueDates)]

#Gradtagszahlen
gtzliste=[]
for i in enumerate(uniqueDates):
    if averages[i[0]]<hg:
        gtz1=[rt- averages[i[0]]]    
    else:
        gtz1=[0]  
    gtzliste.extend(gtz1)  
#Heizgradtage 
hgtliste=[]
for i in enumerate(uniqueDates):
    if averages[i[0]]<hg:
        hgt=[hg- averages[i[0]]]    
    else:
        hgt=[0]  
    hgtliste.extend(hgt)   
#create new DataFrame
tagDF  = pd.DataFrame(data =[[countsAndDays[i],averages[i],gtzliste[i],hgtliste[i]] for i in range(uniqueDates.__len__())],columns = ["Messwerte","AverageTemp",gradtag,hgtag],index = uniqueDates)

tagDF.to_excel(adatei)

