# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 12/7/2020
# Update Date: 27/9/2023
# StartDay.py: Opens programs and sets up items for the day

#Libraries
import os
import glob
from datetime import datetime

#Variables
folder = r"U:\Daily Log"

#Functions
" Main Finction "
def main():
   selection = menu()
   for char in selection:
      file = int(char)
      if file == 1:
         os.startfile(r'C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe')
      elif file == 2:
         os.startfile(r'C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE')
      elif file == 3:
         list_of_files = glob.glob(r'S:\_Component Definitions\_ACCX Component Definition\*') # * means all if need specific format then *.csv
         latest_file = max(list_of_files, key=os.path.getctime)
         os.startfile(latest_file)
      elif file == 4:
         list_of_files = glob.glob(r'S:\GEMBA Board\*.xlsx') # * means all if need specific format then *.csv
         latest_file = max(list_of_files, key=os.path.getctime)
         os.startfile(latest_file)
      elif file == 5:
         os.startfile(r'C:\NTC Vault')
      else:
         print("Invalid selection, please try again\n")
   fileDay = datetime.now().strftime('%d/%m/%Y')
   fileTime = datetime.now().strftime('%H:%M')
   day = datetime.now().strftime('%y%m%d')
   file = folder + "\\Log " + day + ".txt"
   with open(file,"x") as f:
      f.write("Daily Log {0}\n{1}: Day Start".format(fileDay, fileTime))

def menu():
   print("|  Program selection  |")
   print("|---------------------|")
   print("| 1) Solidworks       |")
   print("| 2) Outlook          |")
   print("| 3) ACCX Product Def.|")
   print("| 4) Gemba            |")
   print("| 5) PDM              |")
   print("|---------------------|")
   return input("Which programs are being opened?\n")

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()