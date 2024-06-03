# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 27/9/23
# Update Date: 3/6/23
# WriteLog.py

"""A one line summary of the module or program, terminated by a period.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
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
folder = r"U:\Daily Log\*"
today = datetime.now().date()
fileDate = lambda x: datetime.fromtimestamp(os.path.getctime(x)).date()

#Functions
" Main Finction "
def main():
   list_of_files = glob.glob(folder) # * means all if need specific format then *.csv
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