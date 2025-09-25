# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 18/9/2025
# Update Date: 25/9/2025
# Part_Issuing_Check.py

"""This program takes a parts issued report for a range of units and says what parts are needed.

A .CSV file is loaded in before being sorted. The rows are filtered to get
rows with part numbers and work order numbers. The rows are then assigned
their work order numer and sorted by part number before being saved to an
excel sheet. 

Functions:
   main: Driver of the program
   fileCheck: Checks if a file is open
"""
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
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
from MenuMaker import makeMenu
#pylint: enable=wrong-import-position

#Variables
folderPath = fr"\\{Shared_Drive}\Purchase Orders\PO's to approve from Owatonna"\
             r"\reorder reports\Work In Process Reports"
fileStartPath = folderPath + r'\*.CSV'
finalFolder = folderPath + r'\Processed Files'

#Functions
def main():
   """Pulls in a file, sorts the data, and saves to a new file"""
   loader = Loader("Loading Work Order Files...", "Work Order Files Loaded", 0.1).start()
   files = sorted(glob.iglob(fileStartPath), key=os.path.getmtime, reverse=True)
   loader.stop()

   filenamelist = []
   for file in files:
      filenamelist.append(file.split("\\")[-1:][0])

   makeMenu("Recent Readable Files", filenamelist)
   select = int(input("Select the File you would like to read\n"))
   filePath = files[select -1]
   fileName = filePath.split('\\')[-1:][0].split('.')[:1][0]

   loader = Loader("Loading Data...", "Data Loaded", 0.1).start()

   file = fileCheck(filePath)
   if file == 0:
      return

   file.close()

   with open(filePath,encoding="utf-8") as f:
      keywords = ("-","WORK ORDER NO:","MAKE FOR: S/O")
      lines = [line.split() for line in f if any(keyword in line for keyword in keywords)]

   count = 0
   parts = []
   workOrder = "1111A01"
   workOrderDate = "01/01/25"
   for line in lines:
      if line[0] == "MAKE":
         check = lines[count-1][3]
         if len(check) == 7:
            workOrder = check
            workOrderDate = lines[count-1][6]
      if len(line) == 10 and line[0] != "MAKE":
         line.insert(1,workOrder)
         line.insert(2,workOrderDate)
         parts.append(line)
      count += 1

   loader.stop()

   loader = Loader("Sorting Data...", "Data Sorted", 0.1).start()
   rawdf = pd.DataFrame(parts, columns = ["COMPONENT NUMBER","W/O","DATE","STEP",
                                          "U/M","QTY/PARENT","QTY REQD",
                                          "QTY ISSUED","MATL COST",
                                          "FIX OVERHEAD","VAR OVERHEAD","TOTAL"])

   cleandf = rawdf.drop_duplicates().reset_index(drop=True)
   cleandf = cleandf.sort_values(by = ["COMPONENT NUMBER","W/O"])
   partdf = cleandf.drop(columns=["STEP","U/M","MATL COST","FIX OVERHEAD",
                                  "VAR OVERHEAD","TOTAL"])
   partdf.index += 1

   loader.stop()

   pd.set_option('display.max_rows', None)

   loader = Loader("Writing to Excel...", "Data Written", 0.1).start()
   partdf = partdf.astype({'DATE': 'datetime64[ns]',
                                       'QTY/PARENT': 'float',
                                       'QTY REQD': 'float',
                                       'QTY ISSUED': 'float',})
   saveFile = finalFolder + "\\" + fileName + " Processed.xlsx"
   writer = pd.ExcelWriter(saveFile, engine='xlsxwriter')
   partdf.to_excel(writer, sheet_name="Parts")
   writer.close()
   loader.stop()

   print(f"File Saved as:\n{saveFile}")


def fileCheck(logFile):
   """Checks if a specified file is open in another program or by another user.
   
   Parameters:
      LogFile (str): Specified file location

   Returns:
      file (file): Specified file with read permissions
   """
   try:
      file = open(logFile,"r",encoding="utf-8")
   except PermissionError as err:
      print(f"Permission Error: {err}")
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      input('Program Terminating. Press ENTER to Close...')
      return 0
   return file

if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")
