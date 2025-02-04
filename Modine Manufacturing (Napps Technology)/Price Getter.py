# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 4/2/2025
# Price Getter.py
# Rev 1

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
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
#pylint: enable=wrong-import-position

#variables
folder = "2024 Q1.2"

#Input Sheets
costedfile = fr'\\{Shared_Drive}\Ancel\Pricing\{folder}\BM_CostedMaterials.xls'
optionfile = fr'\\{Shared_Drive}\Ancel\Pricing\{folder}\sage bom options.xls'
partfile = fr'\\{Shared_Drive}\Ancel\Pricing\{folder}\IM_PriceList.xls'

#Output Sheets
optionOut = fr'\\{Shared_Drive}\Ancel\Pricing\{folder}\option price out.xlsx'
topOut = fr'\\{Shared_Drive}\Ancel\Pricing\{folder}\top price out.xlsx'

#Functions
" Main Finction "
def main():
   loader = Loader("Reading Pricing Files...", "Prices Collected\n", .1).start()

   " Costed Bill of Materials Load "
   dfCost = pd.read_excel(costedfile)
   print(dfCost)
   dfCost = pd.concat([dfCost, pd.read_excel(costedfile, sheet_name = "Sheet2")], ignore_index = True)
   print(dfCost)
   dfCost = dfCost.fillna("")
   costs = pd.DataFrame(getOptionCosts(dfCost), columns = ['BOM','Feature','Option','Price'])
   print(costs)

   " Top Level Bom Current Prices Load "
   dfParts = pd.read_excel(partfile)
   dfParts = dfParts.fillna("")
   del dfParts["Unnamed: 1"]
   partCosts = pd.DataFrame(getPartCosts(dfParts), columns = ['BOM','Price','New Price'])
   # print(partCosts)

   " Options and Current Prices Load "
   dfOpt = pd.read_excel(optionfile)
   # print(dfOpt)

   loader.stop()

   dfOpt["newCost"] = 0
   loader = Loader("Updating Option Costs...", "Pricing Updated\n", .1).start()
   dfOut = updateOptionPrice(costs, dfOpt)
   loader.stop()
   # print(dfOut)

   loader = Loader("Updating BOM Costs...", "BOM Pricing Updated\n", .1).start()
   dfOutBom = updateBOMPrice(costs, partCosts)
   print(dfOutBom)
   loader.stop()

   optionWriter = pd.ExcelWriter(optionOut)
   dfOut.to_excel(optionWriter, sheet_name='Sheet1')
   optionWriter.close()

   BOMWriter = pd.ExcelWriter(topOut)
   dfOutBom.to_excel(BOMWriter, sheet_name='Sheet1')
   BOMWriter.close()

   input('Program Completed. Press ENTER to Close...')


def getOptionCosts(dfCost):
   costs = []
   count = 0
   for index,row in dfCost.iterrows():
      # print(index)
      if index == 0:
         costs.append([row["\nBill Number"],row["\nOption"],0,0])    #Set the fitrst row of the array
         count += 1
      elif row["\nBill Number"] != "":                                             #Finds Header rows for calculation
         if row["\nOption"] != "Base":                               #If the Feature is not Base, separate Feature and Option
            featureCode = row["\nOption"].split("-")[0]
            optionCode = row["\nOption"].split("-")[1]
            try:
               option = int(optionCode)
            except:
               option = optionCode
            try:
               feature = int(featureCode)
            except:
               feature = featureCode
            costs.append([row["\nBill Number"],feature,option,0])    #Set distict feature and option
         else:
            costs.append([row["\nBill Number"],row["\nOption"],0,0]) #Set Feature to Base
         costs[count-1][3] = dfCost['Unnamed: 15'][index-1]          #Add the Pricing from the line above to the previous entry
         count += 1
   costs[count-1][3] = dfCost['Unnamed: 15'][len(dfCost)-1]                #Write in Final cost to complete list
   return costs

def getPartCosts(dfCost):
   costs = []
   count = 0
   for index,row in dfCost.iterrows():
      if row["BOM"] != "":
         if row["BOM"] == "Run Date:":
            break
         costs.append([row["BOM"],dfCost['Sales'][index+1],0])
         count += 1
   return costs

def updateOptionPrice(costs, dfOpt):
   for index,row in dfOpt.iterrows():
      inBOM = row["BOM"]
      inFeat = row["Feature"]
      inOpt = row["Option"]
      cost = costs.query('BOM == @inBOM and Feature == @inFeat and Option == @inOpt').reset_index(drop = True)
      if len(cost) != 0:
         # print(cost["Price"][0])
         # print(type(cost["Price"]))
         dfOpt.loc[index, "newCost"] = round(float(cost["Price"][0])/.43,2)
         # dfOpt["newCost"][int(index)] = round(cost["Price"][0]/.43,2)
   return dfOpt

def updateBOMPrice(costs, partCosts):
   partCosts["New Price 2"] = 0
   for index,row in partCosts.iterrows():
      inBOM = row["BOM"]
      cost = costs.query('BOM == @inBOM and Feature == "Base"').reset_index(drop = True)
      if len(cost) != 0:
         partCosts.loc[index, "New Price"] = round(float(cost["Price"][0])/.45,2)
         partCosts.loc[index, "New Price 2"] = round((float(cost["Price"][0])/.45)*1.03,2)
         # partCosts["New Price"][int(index)] = round(cost["Price"][0]/.45,2)
   return partCosts

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()