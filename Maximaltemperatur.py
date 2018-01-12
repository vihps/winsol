'''
Created on 07.01.2018

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
monat="1701" 
pfad='D:'+r'\Winsol'+r'\Excel'
datei=pfad+r'\E'+str(monat)+'.xls'
apfad='D:'+r'\Temp'
adatei=apfad+r'\Temp'+str(monat)+'.xls'
#print(datei)
#Raumtemperatur
rt=20
#Heizgrenztemperatur
hg=15
#Spaltenname als string
gradtag="G"+str(rt)+"/"+str(hg)
hgtag='HGT'+str(hg)

dat = pd.read_excel(open(datei,'rb'), sheetname=0, usecols=[0,1,2,38],header=None, skiprows=1,names=['Zeitpunkt','Koll','PuO','HKVL'])

dat.Zeitpunkt = [datetime.strptime(str(i),"%Y-%m-%d %H:%M:%S") for i in dat.Zeitpunkt]

dates= [Zeitpunkt.date() for Zeitpunkt in dat.Zeitpunkt]

#use set() to obtain a unique list of dates
uniqueDates = list(set(dates)) 
uniqueDates.sort() 

countsAndDays      = [dates.count(ud) for ud in uniqueDates]
idx4Dates = [[date[0] for date in enumerate(dates) if date[1] == i] for i in uniqueDates]
averages = [dat.HKVL[idx4Dates[i[0]]].mean() for i in enumerate(uniqueDates)]
averages_k = [dat.Koll[idx4Dates[i[0]]].mean() for i in enumerate(uniqueDates)]
averages_p = [dat.PuO[idx4Dates[i[0]]].mean() for i in enumerate(uniqueDates)]
#Gradtagszahlen
max_h = [dat.HKVL[idx4Dates[i[0]]].max() for i in enumerate(uniqueDates)]
max_k = [dat.Koll[idx4Dates[i[0]]].max() for i in enumerate(uniqueDates)]
max_p = [dat.PuO[idx4Dates[i[0]]].max() for i in enumerate(uniqueDates)]
#create new DataFrame
tagDF  = pd.DataFrame(data =[[countsAndDays[i],averages[i],max_h[i],averages_k[i],max_k[i],averages_p[i],max_p[i]] for i in range(uniqueDates.__len__())],columns = ["Messwerte","HKmittel","HKmax","Kollmittel","Kollmax","PuOmittel","PuOmax"],index = uniqueDates)
#pd.options.display.float_format = '${:,.2f}'.format
tagDF.to_excel(adatei,float_format='%.1f')

