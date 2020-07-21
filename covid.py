#to automate: use Windows Task Scheduler

import json
import requests
import xlsxwriter
import pandas as pd
from datetime import date

#log date script run
today = date.today()
datelog = today.strftime("%y-%m-%d")
filename = "{}CovidData.xlsx".format(datelog)

#get today's json object
url = 'http://www.dph.illinois.gov/sitefiles/COVID19/ResurgenceMetrics.json'
resp = requests.get(url)
data = resp.json()

#extract date, different shape from rest of data
datestamp = data.pop('LastUpdateDate')

#and anything  where regionID = 8
testPositivity = pd.json_normalize(data['testPositivity'])
hospitalAvailability = pd.json_normalize(data['hospitalAvailability'])
cliAdmissions = pd.json_normalize(data['cliAdmissions'])

regionPos = testPositivity[testPositivity["regionID"] == 8]
regionHosp = hospitalAvailability[hospitalAvailability["regionID"] == 8]
regionCli = cliAdmissions[cliAdmissions["regionID"] == 8]

#write panda dataframe to workbook
writer = pd.ExcelWriter(filename, engine="xlsxwriter")
regionPos.to_excel(writer, sheet_name="Positivity Rate")
regionHosp.to_excel(writer, sheet_name="Hospital Availability")
regionCli.to_excel(writer, sheet_name="Clinical Admissions")

#get the xlsxwriter objects 
posSheet = writer.sheets["Positivity Rate"]
hospSheet = writer.sheets["Hospital Availability"]
cliSheet = writer.sheets["Clinical Admissions"]
#and store them in an array
worksheets = [posSheet, hospSheet, cliSheet]

#set columns wide enough to be readable & hide Panda keys
for x in worksheets:
    x.set_column("A:A", 5, None, {'hidden': 1})
    x.set_column("B:E", 25)
    
#refinements:
#   customize column labels
#   visualization to show trends

#close writer and save
writer.save()
    

