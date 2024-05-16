# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 15/12/2020
# Update Date: 14/6/2022
# DrawingLogMaintenance.py: Will categorize all drawings in excel workbook
#       by type

#Libraries
import pandas as pd
import openpyxl as pyxl
import os
import glob

#Variables
dfs = []
# inputWorkbook = 'Drawing Log.xlsx'
# outputWorkbook = 'Drawings by Unit.xlsx'
inputWorkbook = r'S:\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawing Log.xlsx'
outputWorkbook = r'S:\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawings by Unit.xlsx'
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
      dType = row[7]
      newRow = {'Unit':row[1],'Tonage':row[2],'SubUnit':row[3],'Circuit':row[4],
               dType:row[0]}
      newRow[dType] = hyperlink(newRow[dType])
      if row[5] == 1:
         dfOut = pd.concat([dfOut, pd.DataFrame([newRow],columns = {'Unit','Tonage','SubUnit','Circuit',dType})], ignore_index = True)
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
   writer.save()

" Grabs hyperlink of the first file matching the drawing name "
def hyperlink(drawing):
   fileLocation = 'S:\_Approved for use\DRAWINGS (PDF)\\'+drawing+'*'
   if len(glob.glob(fileLocation)) > 0:
      fileLocation = glob.glob(fileLocation)[0]
      drawing = '=HYPERLINK("{}", "{}")'.format(fileLocation, drawing)
   return drawing

" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input("Program completed. Press Enter to close...")
