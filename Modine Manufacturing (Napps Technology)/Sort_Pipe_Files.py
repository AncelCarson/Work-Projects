# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 13/11/24
# Update Date: 4/2/2025
# Sort_Pipe_Files.py

"""Sorts Approved step files and simulated prt files into Approved for Use.

This program reads the file contents of the Shared Drive and Drawings drive to find
bent pipe step files and prt files respectively. It compares those lists and moves files
with matching names into their respective sub folders 

Functions:
   main: Driver of the program
"""
#Libraries
import os
import sys
import glob
import shutil
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')
Drawing_Drive = os.getenv('Drawing_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
#pylint: enable=wrong-import-position

#Variables
stepLocation = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Bent Pipe\STEP Files'
prtSource = fr'\\{Drawing_Drive}\_Drawings Pending Approval (PDF)\Bent Pipe'
prtLocation = fr'\\{Shared_Drive}\_Approved for use\DRAWINGS (PDF)\Bent Pipe\BENDER Files'

loader1 = Loader("Collecting File Lists...", "Files Collected", 0.1).start()
stepFiles = glob.glob(stepLocation + r"\*.step")
prtFiles = glob.glob(prtSource + r"\*.prt")
fileList = [os.path.basename(x).split(".")[0] for x in stepFiles]
loader1.stop()

folders = {"2014": r"\FORMED PIPE .625 (CPPR-TBGA-2014)",
           "2016": r"\FORMED PIPE .875 (CPPR-TBGA-2016)",
           "2017": r"\FORMED PIPE 1.125 (CPPR-TBGA-2017)",
           "2018": r"\FORMED PIPE 1.375 (CPPR-TBGA-2018)",
           "2019": r"\FORMED PIPE 1.625 (CPPR-TBGA-2019)",
           "2020": r"\FORMED PIPE 2.125 (CPPR-TBGA-2020)",}

#Functions
def main():
   """ Main Finction """
   count = 0
   missingFiles = []
   existingFiles = []

   loader2 = Loader("Sorting Files...", "Files Sorted", 0.1).start()
   for fileName in fileList:
      try:
         # returns the file path if it exists in the lsit of part files
         prtFile = [s for s in prtFiles if fileName in s][0]
         size = fileName.split("-")[2][:4]
         folder = folders[size]
         os.rename(stepFiles[count],(stepLocation + folder + "\\" + fileName + ".step"))
         shutil.move(prtFile,(prtLocation + folder + "\\" + fileName + ".prt"))
      except IndexError:
         missingFiles.append(fileName)
      except FileExistsError:
         existingFiles.append(fileName)
      count += 1
   loader2.stop()

   print("The following files were missing:")
   print(*missingFiles, sep="\n")

   print("The following files already existed in approved folder:")
   print(*existingFiles, sep="\n")


if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")
