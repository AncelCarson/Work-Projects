#pylint: disable = invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 15/12/2020
# Update Date: 4/2/2025
# DrawingLogMaintenance.py: Will categorize all drawings in excel workbook
#       by type

#Libraries
import os
import glob
import pandas as pd
import openpyxl as pyxl
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#Variables
dfs = []
# inputWorkbook = 'Drawing Log.xlsx'
# outputWorkbook = 'Drawings by Unit.xlsx'
inputWorkbook = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawing Log.xlsx'
outputWorkbook = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawings by Unit.xlsx'
worksheets = ['ACCS','ACCM','CCAR','CGWR','FWCD','NWCR','SWU','WCCU']
writer = pd.ExcelWriter(outputWorkbook)

#Functions
" Main Finction "
def main():
   dfIn = pd.read_excel(inputWorkbook)
   dfOut = pd.DataFrame(columns = ['Unit','Tonage','SubUnit','Circuit','DISH',\
                                    'HGIJ','LQLN','SUCT'])
   print(dfIn)
   for index, row in dfIn.iterrows():
      dType = row["DrawingType"]
      newRow = {'Unit':row["ParentUnit"],'Tonage':row["Tonage"],'SubUnit':row["SubUnit"],'Circuit':row["Circuit"],
               dType:row["DrawingNumber"]}
      newRow[dType] = hyperlink(newRow[dType])
      if row["Active"] == 1:
         dfOut = pd.concat([dfOut, pd.DataFrame([newRow],columns = ['Unit','Tonage','SubUnit','Circuit',dType])], ignore_index = True)
   dfOut.sort_values(by=['Unit','Tonage','SubUnit','Circuit'],inplace=True)
   dfOut = dfOut.fillna("")
   for sheet in worksheets:
      dfs.append(dfOut.query('Unit == @sheet'))
   for num in range(len(worksheets)):
      del dfs[num]['Unit']
      dfp = dfs[num].groupby(by=['Tonage','SubUnit','Circuit']).sum()
      print('\n',worksheets[num])
      print(dfp)
      if len(dfs[num]) == 0:
         continue
      dfp.to_excel(writer, sheet_name=worksheets[num])
   writer.close()

" Grabs hyperlink of the first file matching the drawing name "
def hyperlink(drawing):
   fileLocation = f'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\\'+drawing+'*'
   if len(glob.glob(fileLocation)) > 0:
      fileLocation = glob.glob(fileLocation)[0]
      drawing = '=HYPERLINK("{}", "{}")'.format(fileLocation, drawing)
   return drawing

" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")
