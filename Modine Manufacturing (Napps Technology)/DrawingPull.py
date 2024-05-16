# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 2/3/2021
# Update Date: 29/12/2022
# DrawingPull.py

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
#Libraries
import os
import sys
import glob
import pandas as pd

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
from Loader import Loader

#Variables
drawingWorkbook = r'S:\_Approved for use\DRAWINGS (PDF)\Drawing Logs\Drawing Log.xlsx'#U:\_Programs\Python\Drawing Log.xlsx'

#Functions
" Main Finction "
def main():
   jobFolder = getWorkOrder()
   drawingList = dataQueries()
   copyFiles(jobFolder, drawingList)

" Creates file path for job folder "
def getWorkOrder():
   wo = input("Please enter the Work Order Number (ex: 4501A-01)\n")
   path = "S:\_A NTC GENERAL FILES\_JOB FILES\Job " + wo + "*"
   # path = "U:\_Programs\Python\File Drop Test\Job " + wo + "*"
   files = glob.glob(path)
   if len(files) > 0:
      filelocation = files[0] + "\Drawings - Mechanical"
   else:
      filelocation = "S:\_A NTC GENERAL FILES\_JOB FILES\Job " + wo + "\Drawings - Mechanical"
      # filelocation = "U:\_Programs\Python\File Drop Test\Job " + wo + "\Drawings - Mechanical"
   return filelocation

" Filters data to user selection "
def dataQueries():
   drawingList = []
   dfIn = pd.read_excel(drawingWorkbook)
   dfOut = pd.DataFrame(columns = ['Unit','Tonage','SubUnit','Circuit','Drawing'])
   drawings = pd.DataFrame(columns = ['Unit','Tonage','SubUnit','Circuit','Drawing'])
   loader = Loader("Collecting active Drawing Numbers...", "Drawings Collected\n", .1).start()
   for index, row in dfIn.iterrows():
      newRow = {'Unit':row[1],'Tonage':row[2],'SubUnit':row[3],'Circuit':row[4],
               'Drawing':row[0]}
      if row[5] == True:
         dfOut = pd.concat([dfOut, pd.DataFrame([newRow],columns = dfOut.columns)], ignore_index = True)
   loader.stop()
   dfOut.sort_values(by=['Unit','Tonage','SubUnit','Circuit'],inplace=True)
   print(dfOut)
   type = typeMenu(dfOut['Unit'].unique())
   dfOut = dfOut.query('Unit == @type')
   tons = tonsMenu(dfOut['Tonage'].unique())
   dfOut = dfOut.query('Tonage == @tons')
   dfOut['SubUnit'] = dfOut['SubUnit'].fillna("None")
   sub = subMenu(dfOut['SubUnit'].unique())
   dfOut = dfOut.query('SubUnit == @sub')
   print("\nPulling Drawing Names...\n")
   for index, row in dfOut.iterrows():
      row['Drawing'] = hyperlink(row['Drawing'])
      drawings = pd.concat([drawings, pd.DataFrame([row], columns=drawings.columns)], ignore_index = True)
   for index, row in drawings.query('Drawing != "None"').iterrows():
      drawingList.append(row['Drawing'])
   return drawingList

" Copies drawings to job folder "
def copyFiles(jobFolder, drawingList):
   if len(drawingList) == 0:
      print("There are no drawing on file for this unit\n")
      return
   os.system('mkdir "{}"'.format(jobFolder))
   for drawing in drawingList:
      os.system('copy "{}" "{}"'.format(drawing, jobFolder))

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
   fileLocation = 'S:\_Approved for use\DRAWINGS (PDF)\\'+drawing+'*'
   if len(glob.glob(fileLocation)) > 0:
      fileLocation = glob.glob(fileLocation)[0]
   else:
      fileLocation = "None"
   return fileLocation

" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input("Program completed. Press Enter to close...")
