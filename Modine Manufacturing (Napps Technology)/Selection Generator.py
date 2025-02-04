# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 4/2/2025
# Selection Generator.py

#Libraries
import os
import sys
import copy
import pandas as pd
import openpyxl as pyxl
from datetime import datetime
from openpyxl import workbook
from dotenv import load_dotenv
from openpyxl.worksheet.datavalidation import DataValidation as DV

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')
Drawing_Drive = os.getenv('Drawing_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
from MenuMaker import makeMenu
#pylint: enable=wrong-import-position

#Variables
userFile = fr'\\{Drawing_Drive}\Engineering\Performance Software\Selection Users.txt'
startFile = fr'\\{Drawing_Drive}\Engineering\Performance Software\_JESS Output Format Test.xlsx'
repTable = fr'\\{Drawing_Drive}\Engineering\Performance Software\_Sales Reps.xlsx'

#Functions
" Main Finction "
def main():
   dfReps = pd.read_excel(repTable, index_col = 0)
   user = userMenu()
   running = True
   while running:
      print("\nWhat is the selection for?")
      selectionType = input("Enter tonnage and unit type. ex: '50T FWCD'\n")
      salesRep = input("\nWho is the sales rep?\n")
      salesOffice = getSalesOffice(salesRep, dfReps)
      jobName = input("\nPlease enter the job name.\n")
      filePath = createFile(user, selectionType, salesRep, jobName)
      writeFile(filePath, startFile, salesRep, salesOffice[0], jobName)
      dfReps = salesOffice[1]
      running = finishMenu()
   try:
      repWriter = pd.ExcelWriter(repTable)
      dfReps.to_excel(repWriter, sheet_name='Sheet1')
      repWriter.close()
   except Exception as e:
      print("Program crashed do to: {}".format(e))
      print("The Reps Table was not Updated")
      input("Press Entetr to Close the Program...")


def userMenu():
   file = open(userFile,"r")
   names = file.readlines()
   file.close()
   for num in range(len(names)-1):
      names[num] = names[num][:-1]
   makeMenu("User", names)
   selection = int(input("Who is running the selection?\n"))
   if selection > len(names):
      print("An incorrect value was given.")
      print("Please select an available user.")
      user = userMenu()
   else:
      user = "_" + names[selection-1] + "'s Runs"
   return user

def finishMenu():
   print("\nThe file has been created.")
   selection = input("Do you need to create a second file? (Y/N)\n")
   if selection == 'Y' or selection == 'y':
      return True
   elif selection == 'N' or selection == 'n':
      return False
   else:
      print("The input given is not a valid answer. Try again.")
      return finishMenu()

def getSalesOffice(rep, dfReps):
   dfNewReps = dfReps
   if dfReps['SalesRep'].str.contains(rep).any():
      dfQuery = dfReps.query('SalesRep == "{}"'.format(rep))
      office = dfQuery.iloc[0]["SalesOffice"]
      count = dfQuery.iloc[0]["Count"] + 1
      dfReps.loc[dfReps["SalesRep"] == rep, "Count"] = count
      dfNewReps = dfReps
      print("Works for " + office)
      print("Selection run: {}".format(count))
   else:
      print("This sales rep is not listed.")
      selection = input("Is this a new rep? (Y/N)\n")
      if selection == 'Y' or selection == 'y':
         office = input("Please enter the name of the sales office.\n")
         dfNewReps = addRep(rep, office, dfReps)
      else:
         print("The name you have entered is not listed and may be incorrect.")
         rep = input("Please enter the name of the sales rep.\n")
         answer = getSalesOffice(rep, dfReps)
         office = answer[0]
         print()
   return [office,dfNewReps]

def addRep(rep, office, dfReps):
   dfNewReps = pd.concat([dfReps, pd.DataFrame([{"SalesRep" : rep, "SalesOffice" : office, "Count" : 1}])], ignore_index = True)
   return dfNewReps

def createFile(user, type, rep, name):
   filePath = []
   userFolder = fr'\\{Drawing_Drive}\Engineering\Performance Software' + '\\' + user + '\\'
   fileName = ", ".join([type, rep, name])
   day = datetime.now().strftime('%y%m%d')
   folderName = day + "-" + fileName
   filePath.append(userFolder + folderName)
   filePath.append(filePath[0] + '\\' + fileName + '.xlsx')
   return filePath

def writeFile(filePath, startFile, salesRep, salesOffice, jobName):
   os.system('mkdir "{}"'.format(filePath[0]))
   os.system('copy "{}" "{}"'.format(startFile, filePath[1]))
   workBook = pyxl.load_workbook(filename = filePath[1])
   workSheet = workBook.active
   for sheet in [1, 4, 7, 11, 17]:
      workBook.active = sheet
      workSheet = workBook.active
      workSheet['P7'].value = salesRep
      workSheet['P8'].value = salesOffice
      workSheet['P9'].value = jobName
   loader = Loader("Creating File...", "File Created", 0.1).start()
   addValidation(workBook)
   del workBook['Sacrifice']     #Sacrifice Sheet holdsall the bad Validation and is then deleted
   workBook.save(filePath[1])
   workBook.close()
   loader.stop()

def addValidation(workBook):
   addVal = []
   tenCount = DV(type="whole", operator="between", formula1=0, formula2=10, allow_blank=True)
   addVal.append([tenCount])
   addVal[0].append([["P7","P10"],["Q27:Q29"],["Q27:Q29"],["P7","P10"],["Q30:Q31"],["Q30:Q31"],["P7","P10"],["Q27:Q29"],["Q27:Q29"],["P8","P12"],["P7","P10"],["Q29:Q31"],[0],["Q29:Q31"],[0],[0],["P7","P10"],["Q27"],["Q27"]])
   frameType = DV(type="list", formula1="'Look Up Tables'!$V$5:$V$7", allow_blank=True)
   addVal.append([frameType])
   addVal[1].append([[0],["P4"],[0],[0],[0],[0],[0],["P4"],[0],[0],[0],["P4"],[0],["P4"],[0],[0],[0],["P4"]])
   fwcdType = DV(type="list", formula1="'Look Up Tables'!$AH$7:$AH$14", allow_blank=True)
   addVal.append([fwcdType])
   addVal[2].append([[0],[0],[0],[0],["P4"],["P4"],[0],["Q3"],["Q3"]])
   circuitCount = DV(type="list", formula1="'Look Up Tables'!$AA$5:$AA$6", allow_blank=True)
   addVal.append([circuitCount])
   addVal[3].append([[0],["P5"],["P5"],[0],[0],[0],[0],["P5"],["P5"],[0],[0],["P5"],[0],["P5"]])
   fluidType = DV(type="list", formula1="'Look Up Tables'!$J$5:$J$7", allow_blank=True)
   addVal.append([fluidType])
   addVal[4].append([[0],["P14"],[0],[0],["P14","P21"],[0],[0],["P14"],[0],[0],[0],["P14"],["P14"],[0],[0],[0],[0],["P14"]])
   glycol = DV(type="list", formula1="'Look Up Tables'!$I$5:$I$18", allow_blank=True)
   addVal.append([glycol])
   addVal[5].append([[0],["P15"],[0],[0],["P15","P22"],[0],[0],["P15"],[0],[0],[0],["P15"],["P15"],[0],[0],[0],[0],["P15"]])
   compressor = DV(type="list", formula1="'Look Up Tables'!$B$5:$B$54", allow_blank=True)
   addVal.append([compressor])
   addVal[6].append([["P30:P32"],["P27:P28"],["P27:P28"],["P31:P33"],["P30:P31"],["P30:P31"],["P30:P32"],["P27:P28"],["P27:P28"],["P33:P35"],["P30:P32"],["P29:P30"],[0],["P29:P30"]])
   fans = DV(type="list", formula1="'Look Up Tables'!$Z$5:$Z$7", allow_blank=True)
   addVal.append([fans])
   addVal[7].append([[0],["P29"],["P29"],[0],[0],[0],[0],["P29"],["P29"],[0],[0],[0],[0],[0],[0],[0],[0],["P27"]])
   evap = DV(type="list", formula1="'Look Up Tables'!$T$5:$T$9", allow_blank=True)
   addVal.append([evap])
   addVal[8].append([[0],["P48"],[0],[0],["P44"],[0],[0],["P45"],[0],[0],[0],["P47"],[0],["P47"]])
   cond = DV(type="list", formula1="'Look Up Tables'!$U$5:$U$6", allow_blank=True)
   addVal.append([cond])
   addVal[9].append([[0],[0],[0],[0],["P45"]])
   volt = DV(type="list", formula1="'Look Up Tables'!$L$5:$L$7", allow_blank=True)
   addVal.append([volt])
   addVal[10].append([[0],["P26"],[0],[0],["P29"],[0],[0],["P26"],[0],[0],[0],["P28"],[0],["P28"],[0],[0],[0],["P26"]])
   ambient = DV(type="list", formula1="'Look Up Tables'!$K$5:$K$6", allow_blank=True)
   addVal.append([ambient])
   addVal[11].append([[0],["P23"],[0],[0],[0],[0],[0],["P23"],[0],[0],[0],["P25"],["P25"],[0],[0],[0],[0],["P23"]])
   addPump = DV(type="list", formula1="'Look Up Tables'!$AY$34:$AY$40", allow_blank=True)
   addVal.append([addPump])
   addVal[12].append([[0],["Q30"],["Q30"]])
   twoCount = DV(type="whole", operator="between", formula1=0, formula2=2, allow_blank=True)
   addVal.append([twoCount])
   addVal[13].append([[0],["Q30"],["Q30"]])
   twentyCount = DV(type="whole", operator="between", formula1=0, formula2=20, allow_blank=True)
   addVal.append([twentyCount])
   addVal[14].append([["Q30:Q32"],[0],[0],["Q31:Q33"],[0],[0],["Q30:Q32"],[0],[0],["Q33:Q35"],["Q30:Q32"],[0],[0],[0],[0],["A1"]]) # A! set because anything else destroys the Secondary Heating
   # calcMOP_MCA = DV(type="list", formula1="'Look Up Tables'!$AU$5:$AU$6", allow_blank=True)
   # addVal.append([calcMOP_MCA])
   # addVal[13].append([["S24"],[0],[0],["S27"]])
   ## _JESS Output Format 210215 containes cell setup
   writeValidation(workBook, addVal)

def writeValidation(workBook, addVal):
   for type in addVal:
      validate = type[0]
      sheetValidate = []
      for sheet in range(len(type[1])):
         sheetValidate.append(copy.deepcopy(validate))
         workBook.active = sheet
         workSheet = workBook.active
         workSheet.add_data_validation(sheetValidate[sheet])
         for cell in type[1][sheet]:
            if cell == 0:
               pass
            else:
               if len(cell) > 4:
                  cell = cell.split(":")
                  for piece in workSheet["{}".format(cell[0]):"{}".format(cell[1])]:
                     sheetValidate[sheet].add(piece[0])
               else:
                  sheetValidate[sheet].add(workSheet[cell])

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()
