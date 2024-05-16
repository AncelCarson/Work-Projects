# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 21/6/2022
# Update Date: 7/9/2022
# TechChartGenerator.py
# Rev: 1

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
import numpy as np
import pandas as pd
from playsound import playsound
from datetime import timedelta
import matplotlib.pyplot as plt

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
from JetsonChime import chime
from MenuMaker import menu
from MenuMaker import makeMenu
from JetsonControl import logo


#Variables
chartFolder = r'S:\_NTC Technical Support Database\Chart Viewer\\'
wavChime = r'S:\\Ancel\\Python_Modules\\AdIns\\Resources\\jetson.wav'

#Functions
" Main Finction "
def main():
   logo()
   chime()
   filenamelist = []
   menuList = []
   showColumns = []

   for filename in os.listdir(chartFolder):
      if filename.endswith(".xls"):
         filenamelist.append(filename)

   makeMenu("Readable Files", filenamelist)
   select = int(input("Select the File you would like to read\n"))

   file = chartFolder + filenamelist[select-1]
   chartData = pd.read_excel(file, header = 6, nrows = 120)
   menuItems = np.char.strip(list(chartData.columns.values), chars = " ")
   for count in range(len(menuItems)):
      menuList.append(str(str(count) + ": " + menuItems[count]))

   runType = runSelect()

   first = True
   Showmenu = True
   while True:
      showColumns = []
      if Showmenu:
         menuList[0] = 'Chart Options'
         menu(menuList)
         Showmenu = False
      if first:
         selection = input("List data you want displayed separated by commas\n")
         selections = selection.split(",")
         dataLines = [int(value) for value in selections]
         first = False
      else:
         menu(["Additional","A: End Program","B: Add Data","C: Remove Data","D: Make New Selection","E: Change Run Type"])
         selection = input("Please select the next step\n")
         if selection == 'a' or selection == 'A':
            break
         elif selection == 'b' or selection == 'B':
            getSelections(dataLines, menuList)
            selection = input("List additioanl data you want displayed separated by commas\n")
            selections = selection.split(",")
            for value in selections:
               dataLines.append(int(value))
         elif selection == 'c' or selection == 'C':
            getSelections(dataLines, menuList)
            selection = input("List data you want removed separated by commas\n")
            selections = selection.split(",")
            for value in selections:
               del dataLines[dataLines.index(int(value))]
         elif selection == 'd' or selection == 'D':
            first = True
            continue
         elif selection == 'e' or selection == 'E':
            runType = runSelect()

      for clm in dataLines:
         showColumns.append(chartData.columns[clm])


      if runType == 1:
         chartData.plot(x = ' TIME     ', y = showColumns)
         plt.show()
      elif runType == 2:
         getRuntime(showColumns, chartData)
         input("Press enter to Continue...")

def chime():
   playsound(wavChime)

def runSelect():
   menu(["Program Options","A: Make a chart","B: Find Runtime"])
   runSelect = input("Please select the next step\n")
   if runSelect == 'a' or runSelect == 'A':
      return 1
   elif runSelect == 'b' or runSelect == 'B':
      return 2

def getSelections(dataLines, menuList):
   usedLines = ["Plotted Data"]
   for value in dataLines:
      usedLines.append(menuList[value])
   menu(usedLines)

def getRuntime(showColumns, chartData):
   for clm in showColumns:
      idx = (chartData[clm] == 0).idxmax()
      end = chartData[' TIME     '].loc[chartData.index[0]]
      start = chartData[' TIME     '].loc[chartData.index[idx-1]]
      t1 = timedelta(hours=start.hour, minutes=start.minute, seconds=start.second)
      t2 = timedelta(hours=end.hour, minutes=end.minute, seconds=end.second)
      print("Runtime for {0} was {1}".format(clm,t2 - t1))

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()