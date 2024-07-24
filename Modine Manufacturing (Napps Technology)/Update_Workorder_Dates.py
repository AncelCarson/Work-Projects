# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 22/7/24
# Update Date: 24/7/24
# Update_Workorder_Dates.py

"""Creates a csv file with most recent prodiction start dates for all open Work Orders.

This program pulls in the most recent GEMBA board and an export from SAGE with
all the open work orders. It then searches the GEMBA for those order numbers 
and saves the updated Production Start date. That is saved as a csv file that 
can be imported back into SAGE. 

Functions:
   main: Driver of the program
   fileCheck: Checks if a file is open
"""
#Libraries
import os
import sys
import glob
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
from Loader import Loader

#Variables
GEMBA_path = r'S:\GEMBA Board\*.xlsx' # * means all if need specific format then *.xlsx
WorkOrder_path = r'U:\_Programs\Open Workorders.csv'
WorkOrderImport_path = r'U:\_Programs\Open Workorders Update.csv'

#Functions
" Main Finction "
def main(): #Line 2 (H2O) - Schedule
   loader = Loader("Loading GEMBA Files...", "GEMBA Files Loaded", 0.1).start()
   GEMBAs = sorted(glob.iglob(GEMBA_path), key=os.path.getctime, reverse=True)
   latest_GEMBA = GEMBAs[0]
   loader.stop()

   file = fileCheck(latest_GEMBA)
   if file == 0:
      return
   file.close()

   loader = Loader("Loading GEMBA Tables...", "GEMBA Tables Loaded", 0.1).start()
   dfAir = pd.read_excel(latest_GEMBA, sheet_name = 'Line 1 (AIR) - Schedule', header = 4)
   dfWater = pd.read_excel(latest_GEMBA, sheet_name = 'Line 2 (H2O) - Schedule', header = 4)
   dfIn = pd.concat([dfAir, dfWater])
   loader.stop()

   loader = Loader("Loading Work Order Tables...", "Work Order Tables Loaded", 0.1).start()
   dfWO = pd.read_csv(WorkOrder_path, names = ["WORK_ORDER", "WO_STATUS", "EFFECTIVE_DATE", "WO_DUE_DATE", "SCHED_REL_DATE"])
   loader.stop()

   loader = Loader("Formatting Data...", "Data Formatted", 0.1).start()
   dfIn = dfIn[dfIn.Priority != "Shipped"]
   dfIn = dfIn[dfIn.Priority.notna()]
   dfIn = dfIn[dfIn["Production \nStart \nDate"].notna()]
   workOrderDates = dfIn[["Work Order#", "Production \nStart \nDate"]]
   workOrderDates.sort_values(by = ["Production \nStart \nDate"], inplace = True)
   workOrderDates["Work Order#"] = workOrderDates.apply(lambda x: x["Work Order#"][:7], axis = 1)
   workOrderDates["Production \nStart \nDate"] = workOrderDates.apply(lambda x: x["Production \nStart \nDate"].strftime('%m%d%Y'), axis = 1)
   workOrderDates.reset_index(drop = True, inplace = True)
   loader.stop()

   for index, order in dfWO.iterrows():
      workOrder = order.WORK_ORDER
      if not workOrderDates[workOrderDates["Work Order#"] == workOrder]["Production \nStart \nDate"].empty:
         newDate = str(workOrderDates[workOrderDates["Work Order#"] == workOrder]["Production \nStart \nDate"].values[0])
      else:
         print("Work Order {} is missing in GEMBA".format(workOrder))
         continue
      order.EFFECTIVE_DATE = newDate
      order.WO_DUE_DATE = newDate
      order.SCHED_REL_DATE = newDate
   
   dfWO.to_csv(WorkOrderImport_path)

def fileCheck(logFile):
   """Checks if a specified file is open in another program or by another user.
   
   Parameters:
      LogFile (str): Specified file location

   Returns:
      file (file): Specified file with read permissions
   """
   try:
      file = open(logFile,"r")
   except PermissionError as err:
      print("Permission Error: {0}".format(err))
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      input('Program Terminating. Press ENTER to Close...')
      return 0 
   return file

" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input('Program Completed. Press ENTER to Close...')