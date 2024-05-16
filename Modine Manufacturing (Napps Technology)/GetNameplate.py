# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 7/5/2021
# Update Date: 30/6/2021
# GetDrawings.py

#Libraries
import os
import sys
import copy
import glob
import pandas as pd
import openpyxl as pyxl
from openpyxl.worksheet.datavalidation import DataValidation as DV

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
from Loader import Loader
from MenuMaker import menu
from GetDrawings import GetDrawings

class GetNameplate:

   list_of_files = glob.glob(r'I:\Engineering\Nameplates\Name Plate Generator Rev *.xlsx') # * means all if need specific format then *.csv
   nameplateExcel = max(list_of_files, key=os.path.getctime)

   def __init__(self):
      """A one line summary of the module or program, terminated by a period.

      Leave one blank line.  The rest of this docstring should contain an
      overall description of the module or program.  Optionally, it may also
      contain a brief description of exported classes and functions and/or usage
      examples.

      Functions:
         main: Driver of the program
      """

      self.runMenu = False

   def main(self):
      self.getUnit()
      if self.runMenu:
         if self.unit == "CCAR" or self.unit == "CGWR" or self.unit == "WCCU" or self.unit == "NWCR":
            options = getOptions(self.nameplateExcel, 1)
            self.modelString = self.unit + getString(options)
         elif self.unit == "CICD" or self.unit == "FWCD":
            options = getOptions(self.nameplateExcel, 2)
            self.modelString = self.unit + getString(options)
         elif self.unit == "ACCS" or self.unit == "ACCM" or self.unit == "ACCU" or self.unit == "ACCR":
            options = getOptions(self.nameplateExcel, 3)
            self.modelString = self.unit + getString(options)
         else:
            print("An incorrect unit was entered")
      print(self.modelString)

   def getUnit(self):
      self.unit = input("Please enter the unit code (ex: FWCD)\n").upper()
      if len(self.unit) > 4:
         self.modelString = self.unit
         correctLength = checkString(self.unit)
         if correctLength == False:
            print("The given model string is incorrect")
            self.getUnit()
      elif len(self.unit) != 4:
         print("The given unit code is incorrect")
         self.getUnit()
      else:
         self.runMenu = True


def checkString(modelString):
   unit = modelString[:4]
   check = False
   if unit == "CCAR" or unit == "CGWR" or unit == "WCCU" or unit == "NWCR":
      if len(modelString) == 25:
         check = True
   if unit == "CICD" or unit == "FWCD":
      if len(modelString) == 21:
         check = True
   if unit == "ACCS" or unit == "ACCM" or unit == "ACCU" or unit == "ACCR":
      if len(modelString) == 20:
         check = True
   return check

def getOptions(file, unit):
   if unit == 1:
      options = pd.read_excel(file, sheet_name = "Model Strings", nrows = 7)
      options = options.loc[:, ~options.columns.str.contains('^Unnamed')]
      return options
   elif unit == 2:
      options = pd.read_excel(file, sheet_name = "Model Strings", header = 15, nrows = 7)
      options = options.loc[:, ~options.columns.str.contains('^Unnamed')]
      return options
   elif unit == 3:
      options = pd.read_excel(file, sheet_name = "Model Strings", header = 27, nrows = 10)
      options = options.loc[:, ~options.columns.str.contains('^Unnamed')]
      return options

def getString(options):
   size = input("What size unit is this?\n")
   if len(size) == 2:
      size = "0" + size
   modelString = size
   options = options.drop(options.columns[[0,1]],axis=1)
   for feature, values in options.iteritems():
      optionsList = list(values.dropna().values)
      if len(optionsList) == 1:
         modelString += optionsList[0].split("-")[0].replace(" ","")
         continue
      menu([feature] + optionsList)
      modelString += input("Please make a selection.\n").upper()
   return modelString

if __name__ == "__main__":
   obj = GetNameplate()
   obj.main()