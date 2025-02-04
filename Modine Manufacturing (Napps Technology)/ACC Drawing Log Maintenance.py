#pylint: disable = invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/6/2022
# Update Date: 4/2/2025
# ACCDrawingLogMaintenance.py: Will categorize all drawings in excel workbook
#       by type

#Libraries
import os
import sys
import glob
import pandas as pd
import openpyxl as pyxl
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
#pylint: enable=wrong-import-position

#Variables
dfs = []
# inputWorkbook = 'Drawing Log.xlsx'
# outputWorkbook = 'Drawings by Unit.xlsx'
inputWorkbook = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Drawing Logs\ACC Drawing Log.xlsx'
outputWorkbook = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Drawing Logs\ACC Drawings by Size.xlsx'
worksheets = ['10 Ton','15 Ton','20 Ton','25 Ton','30 Ton','40 Ton','50 Ton','55 Ton','60 Ton','70 Ton','80 Ton']
writer = pd.ExcelWriter(outputWorkbook)

#Functions
" Main Finction "
def main():
   dfIn = pd.read_excel(inputWorkbook)
   dfOut = pd.DataFrame(columns = ['Tonage','Application','Steps','Evap','HR','Access',\
                                   'Circuit','DISH','HGIJ','LQLN','SUCT'])
   print(dfIn)
   loader = Loader("Formatting Data...", "Data Formatted", 0.1).start()
   for index, row in dfIn.iterrows():
      dType = row[9]
      newRow = {'Tonage':row[1],'Application':row[2],'Steps':row[3],'Evap':row[4],
                'HR':row[5],'Access':row[6],'Circuit':row[7], dType:row[0]}
      # newRow[dType] = hyperlink(newRow[dType])
      if row[8] == 1:
         dfOut = pd.concat([dfOut, pd.DataFrame([newRow],columns = {'Tonage','Application','Steps','Evap','HR','Access','Circuit',dType})], ignore_index = True)
   loader.stop()
   dfOut.sort_values(by=['Tonage','Application','Steps','Evap','HR','Access','Circuit'],inplace=True)
   dfOut = dfOut.fillna("")
   for sheet in worksheets:
      dfs.append(dfOut.query('Tonage == @sheet'))
   for num in range(len(worksheets)):
      del dfs[num]['Tonage']
      dfp = dfs[num].groupby(by=['Application','Steps','Evap','HR','Access','Circuit']).sum()
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
   input("Program completed. Press Enter to close...")
