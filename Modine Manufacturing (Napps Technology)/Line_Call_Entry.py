#pylint: disable = invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 30/11/2023
# Update Date: 17/4/2024
# Line_Call_Entry.py

"""A one line summary of the module or program, terminated by a period.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
"""
#Libraries
import os
import sys
import pandas as pd
import openpyxl as pyxl
import Generate_ECR as ECR
from datetime import datetime
from dotenv import load_dotenv
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
from MenuMaker import makeMenu
#pylint: enable=wrong-import-position

#Variables.
year = datetime.today().year
logFile = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Line Calls\Line Call Log {year}.xlsx'
templateFile = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Line Calls\Line Call Log Template.xlsx'

#Functions
" Main Finction "
def main():

   " Checking if this year's file exists "
   try: 
      file = open(logFile)
   except FileNotFoundError as err:
      loader = Loader("Creating {} Line Call Log...".format(year),"{} Line Call Log Created\n".format(year))
      os.system('copy "{}" "{}"'.format(templateFile, logFile))
      workBook = pyxl.load_workbook(logFile)
      workBook.active = 15
      workSheet = workBook.active
      workSheet["A2"].value = year
      workBook.active = 0
      workSheet = workBook.active
      workBook.save(logFile)
      workBook.close()
      loader.stop()


   " Checking to See if Log File is in Use "
   file = fileCheck(logFile)
   if file == 0:
      return
   file.close()

   day = datetime.now()

   options = getTables(logFile)

   makeMenu("Product Family", options["Product Family"].dropna())
   unit = options["Product Family"][int(input("What Type of unit is being worked on?\n"))-1]

   workOrder = input("What is the associated Work Order(s)?\nPress ENTER if there is no Work Order.\n")

   parts = input("What parts or drawings were inorrect?\nPress ENTER if there are no parts or drawings.\n")

   makeMenu("Department", options["Department"].dropna())
   department = options["Department"][int(input("What department does this issue belong to?\n"))-1]

   makeMenu(department, options[department].dropna())
   issueCode = options[department][int(input("What is the Associated Issue Code?\n"))-1]

   issueDescription = input("Please describe the issue.\n")
   
   makeMenu("Production Contact", options["Production Contact"].dropna())
   contact = options["Production Contact"][int(input("Who is the production contact?\n"))-1]

   ecr = checkECR(contact, unit, parts, workOrder)

   loader = Loader("Adding Line Call to Table...", "Line Call Added.")
   lineCall = [day,workOrder,unit,contact,parts,department,issueCode,issueDescription,ecr]
   columns = ["B","C","D","E","F","G","H","I","J",]

   workBook = pyxl.load_workbook(logFile)
   workBook.active = 0
   workSheet = workBook.active
   maxRowWithData = 0
   for index, cell in enumerate(workSheet['B']):
      if not cell.value != None:
         maxRowWithData = index
         row = maxRowWithData + 1
         break

   " Creating Cell Style | Only needed to be run once "
   # cellStyle = NamedStyle(name = "cellStyle")
   # cellStyle.font = Font(name='Arial', size=12)
   # bd = Side(style='thin', color="000000")
   # cellStyle.border = Border(left=bd, top=bd, right=bd, bottom=bd)
   # cellStyle.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
   # workBook.add_named_style(cellStyle)

   for count in range(len(columns)):
      cell = columns[count]+str(row)
      workSheet[cell].value = lineCall[count]
      workSheet[cell].style = "cellStyle"
      if columns[count] == "B":
         workSheet[cell].number_format = "m/d/yyyy"

   workBook.save(logFile)
   workBook.close()
   loader.stop()
   
   input('Program Completed. Press ENTER to Close...')

def fileCheck(logFile):
   try:
      file = open(logFile,"r+")
   except PermissionError as err:
      print("Permission Error: {0}".format(err))
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      input('Program Terminating. Press ENTER to Close...')
      return 0 
   return file

def getTables(logFile):
   options = pd.read_excel(logFile, sheet_name = "Lookup Values")
   options = options.loc[:, ~options.columns.str.contains('^Unnamed')]
   return options

def checkECR(contact, unit, parts, workOrder):
   selection = input("Is an ECR required? Y/N\n")
   if selection == "y" or selection == "Y":
      return getECR(contact, unit, parts, workOrder)
   elif selection == "n" or selection == "N":
      return ""
   else:
      print("Invalid selection. Please Try again.\n")
      return checkECR(contact, unit, parts, workOrder)
   
def getECR(contact, unit, parts, workOrder):
   selection = input("Is there already an associated ECR? Y/N\n")
   if selection == "y" or selection == "Y":
      return input("Please enter the Number for the accociated ECR. (Ex. #250)\n")
   elif selection == "n" or selection == "N":
      return ECR.lineCall(contact, unit, parts, workOrder)
   else:
      print("Invalid selection. Please Try again.\n")
      return getECR(contact, unit, parts, workOrder)

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()
