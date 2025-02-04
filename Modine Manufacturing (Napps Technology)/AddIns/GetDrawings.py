# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 2/3/2021
# Update Date: 4/2/2025
# GetDrawings.py

#Libraries
import os
import sys
import glob
import pandas as pd
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Ancel\Python_Modules\AdIns')
from Loader import Loader
#pylint: enable=wrong-import-position

class GetDrawings:

   drawingWorkbook = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawing Log.xlsx'

   def __init__(self, wo='1111A-01', type='FWCD', tons=45, sub='DCST'):
      """Pulls in all drawings current for a Work order and adds them to the job file.

      Leave one blank line.  The rest of this docstring should contain an
      overall description of the module or program.  Optionally, it may also
      contain a brief description of exported classes and functions and/or usage
      examples.

      Functions:
         main: Driver of the program
         getWorkOrder: Finds or makes the file for the work order
         dataQueries: Queries through the data to pull in the appropriate drawings
         copyFiles: Copies drawings from Approved for use and to Job Folder
         typeMenu: Menu for unit Types
         tonsMenu: Menu for Tonnage options for selected unit
         subMenu: Menu for Sub Unit options for selected unit and tonnage
         hyperlink: Builds path of drawings that exist in approved for use
      """
      self.wo = wo
      self.type = type
      self.tons = tons
      self.sub = sub

      self.drawingList = []
      self.drawingsDF = dataQueries()

   def findWO(self):
      self.wo = input("Please enter the Work Order Number (ex: 1111A-01)\n")

   def findUnit(self):
      dfOut = self.drawingsDF
      self.type = typeMenu(dfOut['Unit'].unique())
      dfOut = dfOut.query('Unit == @self.type')
      self.tons = tonsMenu(dfOut['Tonage'].unique())
      dfOut = dfOut.query('Tonage == @self.tons').copy()
      dfOut.fillna({'SubUnit':"None"}, inplace=True)
      self.sub = subMenu(dfOut['SubUnit'].unique())

   def collect(self):
      drawings = pd.DataFrame(columns = ['Unit','Tonage','SubUnit','Circuit','Drawing'])
      dfOut = self.drawingsDF.query('Unit == @self.type')
      print(dfOut)
      dfOut = dfOut.query('Tonage == @self.tons').copy()
      print(dfOut)
      dfOut.fillna({'SubUnit':"None"}, inplace=True)
      print(dfOut)
      dfOut = dfOut.query('SubUnit == @self.sub')
      print(dfOut)
      print("\nPulling Drawing Names...\n")
      for index, row in dfOut.iterrows():
         row['Drawing'] = hyperlink(row['Drawing'])
         drawings = pd.concat([drawings, pd.DataFrame([row],columns = drawings.columns)], ignore_index = True)
      for index, row in drawings.query('Drawing != "None"').iterrows():
         self.drawingList.append(row['Drawing'])
      self.copy()

   def copy(self):
      wo = self.wo
      path = fr"\\{Shared_Drive}\_A NTC GENERAL FILES\_JOB FILES\Job " + wo + "*"
      # path = "U:\_Programs\Python\File Drop Test\Job " + wo + "*"
      files = glob.glob(path)
      if len(files) > 0:
         filelocation = files[0] + r"\Drawings - Mechanical"
      else:
         filelocation = fr"\\{Shared_Drive}\_A NTC GENERAL FILES\_JOB FILES\Job " + wo + r"\Drawings - Mechanical"
      if len(self.drawingList) == 0:
         print("There are no drawing on file for this unit\n")
         return
      os.system('mkdir "{}"'.format(filelocation))
      for drawing in self.drawingList:
         os.system('copy "{}" "{}"'.format(drawing, filelocation))

" Filters data to user selection "
def dataQueries():
   drawingWorkbook = GetDrawings.drawingWorkbook
   dfIn = pd.read_excel(drawingWorkbook)
   dfOut = pd.DataFrame(columns = ['Unit','Tonage','SubUnit','Circuit','Drawing'])
   loader = Loader("Collecting active Drawing Numbers...", "Drawings Collected\n", .1).start()
   for index, row in dfIn.iterrows():
      newRow = {'Unit':row[1],'Tonage':row[2],'SubUnit':row[3],'Circuit':row[4],
               'Drawing':row[0]}
      if row[5] == True:
         dfOut = pd.concat([dfOut, pd.DataFrame([newRow],columns = dfOut.columns)], ignore_index = True)
   loader.stop()
   dfOut.sort_values(by=['Unit','Tonage','SubUnit','Circuit'],inplace=True)
   return dfOut

" Handles Type selection "
def typeMenu(types):
   count = 0
   print("\n| Unit Type")
   print("|----------")
   for  type in types:
      print("| {}: {}".format(count+1, type))
      count += 1
   print("|----------")
   selection = int(input("What is the Unit Type?\n"))
   return types[selection - 1]

" Handles Tons selection "
def tonsMenu(tons):
   count = 0
   print("\n| Unit Tonage")
   print("|------------")
   for  ton in tons:
      print("| {}: {}".format(count+1, ton))
      count += 1
   print("|------------")
   selection = int(input("What Tonnage is the Unit?\n"))
   return tons[selection - 1]

" Handles Subtype selection "
def subMenu(subs):
   count = 0
   print("\n| Subunit Type")
   print("|-------------")
   for  sub in subs:
      print("| {}: {}".format(count+1, sub))
      count += 1
   print("|-------------")
   selection = int(input("What is the Subunit Type?\n"))
   return subs[selection - 1]

" Grabs hyperlink of the first file matching the drawing name "
def hyperlink(drawing):
   fileLocation = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\\'+drawing+'*'
   if len(glob.glob(fileLocation)) > 0:
      fileLocation = sorted(glob.iglob(fileLocation), key=os.path.getctime, reverse=True)[0]
   else:
      fileLocation = "None"
   return fileLocation


if __name__ == "__main__":
   obj = GetDrawings('1111A-01', 'CGWR', 40, 'None')
   # obj.findWO()
   # obj.findUnit()
   obj.collect()
