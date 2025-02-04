# pylint: disable=all

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 27/8/2024
# Update Date: 4/2/2025
# Summarize_ECR.py

"""Summarizes open ECRs into a single table for quick status Updates.

This program loads all of the ECR Forms from the ECR folders that are 
in the top level of the change requests folder. It then reads specific 
cells to get the general data for the open ECRs and exports that to a
file. 

Functions:
   main: Driver of the program
"""
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
summary_sheet = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Requests\ECR Summary.xlsx'
loader = Loader("Collecting Open ECRs...", "ECRs Collected", 0.1).start()
list_of_files = glob.glob(fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Requests\Engineering Change Request *\Engineering Change Request *.xlsx') # * means all if need specific format then *.xlsx
loader.stop()

#Functions
" Main Finction "
def main():
   dataCollection = []
   fileData = [None,None,None,None,None,None,None]
   loader = Loader("Reading Open ECRs...", "ECR Data Collected", 0.1).start()
   for file in list_of_files:
      workbook = pyxl.load_workbook(file, data_only = True)
      worksheet = workbook["Engineering Change Request"]
      fileData[0] = str(worksheet["B2"].value)[27:]
      fileData[1] = str(worksheet["G7"].value)
      fileData[2] = str(worksheet["D38"].value)
      fileData[3] = str(worksheet["G9"].value)
      fileData[4] = str(worksheet["J14"].value)
      fileData[5] = str(worksheet["J16"].value)
      fileData[6] = str(worksheet["G24"].value)
      dataCollection.append(fileData[:])
   dfOut = pd.DataFrame(dataCollection, columns = ["Number", "Date", "Requester", "Request", "Product", "Parts", "Explanation"])
   loader.stop()

   loader = Loader("Writing to Excel...", "Data Written to ECR Summary", 0.05).start()
   writer = pd.ExcelWriter(summary_sheet, engine='xlsxwriter')
   dfOut.to_excel(writer, sheet_name="ECR Summary")
   writer.close()
   loader.stop()


" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")