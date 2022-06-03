from openpyxl import Workbook
import pandas as pd
import glob
import numpy as np
from sqlalchemy import engine_from_config
import xlsxwriter as xlsw

#Customized Pandas's options
pd.options.display.float_format = '{:,.2f}'.format


#Configs
filesPath = 'reportes/**/*.'
fileExtensions = ['csv']
listOfFiles = []
rowsCompiler = []
filters= ['Email','First Name', 'Last Name', 'Department', 'Content', 'Status']
tabsToCreate = ['Content','Status']
tabName = 'General'

#Print warnings
print('Press CTRL+C to exit\n\n')
print('Reading files, please wait...\n')

for extension in fileExtensions:
    listOfFiles.extend(glob.glob(filesPath + extension, recursive=True))

print("Loading files in memory, this can be late a few minutes, don't worry..\n")

for file in listOfFiles:
    rowsCompiler.append(pd.read_csv(file, usecols=filters))
    df_report = pd.concat(rowsCompiler)

cols = ['First Name', 'Last Name']
df_report['Full Name'] = df_report[cols].apply(lambda row: ' '.join(row.values.astype(str)), axis=1)
df_report['Full Namme'] = df_report['Full Name'].str.title()
df_report['Department'] = df_report['Department'].str.upper()
df_report['Status'] = df_report['Status'].str.title()
df_report['Status'] = df_report['Status'].replace('In_Progress', 'In Progress')
df_report['Status'] = df_report['Status'].replace('Not_Started', 'Not Started')
df_report.drop('First Name', inplace=True, axis=1)
df_report.drop('Last Name', inplace=True, axis=1)

df_report = df_report.reindex(columns=['Status', 'Full Name', 'Department', 'Content', 'Email'])

df2 = df_report
#Export to XLSX
resultSaveAs = 'resultado_reporte.xlsx'
writer = pd.ExcelWriter((resultSaveAs), engine='xlsxwriter')


df_report.set_index(['Department', 'Full Name','Email','Content'], inplace=True)

df_report.to_excel(writer, tabName)

criterias = ['Passed', 'In Progress', 'Not Started']
for criteria in criterias:
    
    filters = df2['Status'] == criteria
    df_filtered = df2.loc[filters]
    df_filtered.to_excel(writer, criteria)
    print(df_filtered.head())


writer.save()
