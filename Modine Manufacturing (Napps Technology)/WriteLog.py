# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 27/9/23
# Update Date: 14/6/24
# WriteLog.py

"""Program logs an activity and the time that it was started.

The program opens the day's log file and appends a line to 
the very bottom. This line will have a time stamp followed 
by the activity that was started. That activity is prompred 
in the program for user input. It then saves the file and 
terminates.

Functions:
   main: Driver of the program
   makeFile: Creates a log file if one is not created 
"""
#Starting Query // Moved to the very beginning to reduce loading time
text = input("Log Task Start:\n")

#Libraries
import os
import sys
import glob
from datetime import datetime

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')

#Variables
folder = r"U:\Daily Log\*" # * means all if need specific format then *.txt
today = datetime.now().date()
fileDate = lambda x: datetime.fromtimestamp(os.path.getctime(x)).date()

#Functions
" Main Finction "
def main():
   """Adds a task to the current day's log with with a time stamp"""
   list_of_files = glob.glob(folder) 
   if len(list_of_files) == 0:
      latest_file = makeFile()
   else:
      latest_file = max(list_of_files, key=os.path.getctime)
   print(fileDate(latest_file))
   print(today)
   if fileDate(latest_file) != today:
      latest_file = makeFile()
   logTime = datetime.now().strftime('%H:%M: ')
   logLine = "\n" + logTime + text
   with open(latest_file,"a") as f:
      f.write(logLine)
   if text == "End Week":
      import WeekendReport
      WeekendReport.main()

def makeFile():
   """Creates a log file for the day if one does not exist already
   
   Returns:
      file (str): File path to the newly created log file
   """
   fileDay = datetime.now().strftime('%d/%m/%Y')
   fileTime = datetime.now().strftime('%H:%M')
   day = datetime.now().strftime('%y%m%d')
   file = folder[:-2] + "\\Log " + day + ".txt"
   with open(file,"x") as f:
      f.write("Daily Log {0}\n{1}: Day Start".format(fileDay, fileTime))
   return file

" Checks if this program is being called "
if __name__ == "__main__":
   main()