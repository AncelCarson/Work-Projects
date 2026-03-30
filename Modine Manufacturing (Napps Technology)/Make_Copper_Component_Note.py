# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 30/3/26
# Update Date: 30/3/26
# Make_Copper_Component_Note.py

"""This program generates a list of fopper parts for a unit.

When given a job number, this program will find the folder, 
read copper fittings list, and generate a list of part numbers.

Functions:
   main: Driver of the program
   findFolder: Finds the Bent Pipe Folder(s) in the Drawings Folder
   checkApproval: Checks if a Part is on a different folder
   makeNote: Creates the .txt that holds all the information
"""
#Libraries
import os
import sys
from collections import defaultdict

import glob
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')
Drawing_Drive = os.getenv('Drawing_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
from Get_Job_Copper_Files import Get_Job_Copper_Files
#pylint: enable=wrong-import-position

#Variables
path = fr"\\{Shared_Drive}\_A NTC GENERAL FILES\_JOB FILES\Job "

#Functions
def main():
   """Prints "Hello World" to the terminal"""
   wo = input("Please enter the Work Order Number (ex: 1111A-01)\n")
   folder = path + wo + "*"
   files = glob.glob(folder)
   if len(files) == 1:
      drawingFolder = files[0] + r"\Drawings - Mechanical"
   elif len(files) == 0:
      print("The specified folder does not exist. Please try again")
      return
   else:
      print("There are multiple files that fit match.")
      for folder in files:
         print(folder)
      print("Please correct the folders or narrow your input")
      return

   bendFolders = findFolder(drawingFolder)
   if len(bendFolders) == 0:
      print("There is not folder with the name Bent Pipe")
      print("Please rename the folder or add it")
      return

   jobInst = Get_Job_Copper_Files()
   pipeFolders = defaultdict(str)
   for bendFolder in bendFolders:
      drawingFolder = os.path.dirname(bendFolder)
      pdfs = glob.glob(drawingFolder + r"\*.PDF")
      jobInst.setFiles(pdfs)
      parts = jobInst.getParts()
      pipeFolders[bendFolder] = parts

   loader = Loader("Generating Part Note(s)...", "Note(s) generated\n", .1).start()
   allFlags = []
   for bendFolder in bendFolders:
      makeNote(bendFolder, pipeFolders)
   loader.stop()

   if len(allFlags) != 0:
      errors = {"none": "No note found",
                "count": "Missing or incorrect note fields",
                "length": "Note Length and calculated length did not match",}
      print("The following files had note errors:")
      for flag in allFlags:
         print(f"{flag[0]}: {errors[flag[1]]}")
      print("Check the pipe in Bend Pro and run this simulation again")

def findFolder(folder) -> list:
   """Filters through the Mechanical - Drawings folder to find the Bent Pipe folder"""
   folders = []
   for root, directs, _ in os.walk(folder):
      for direct in directs:
         if "Bent Pipe" in direct:
            folders.append(os.path.join(root, direct))
   return folders

def makeNote(folder, pipeFolders) -> list:
   """Creates the formatted not in the job folder"""
   flags = []
   fileName = folder.split("\\")[-1:][0] + " Parts List.txt"
   filePath = folder + "\\" + fileName
   parts = pipeFolders[folder]
   try:
      with open(filePath,"x",encoding="utf-8") as f:
         f.write("Qty\tPart Number\n")
         for part in dict(sorted(parts.items())):
            qty = parts[part]
            f.write(f"{qty}\t{part}\n")
   except FileExistsError:
      print(f"\n\n{fileName} already exists.\nPlease delete it and run the script again.\n")
   return flags

def to8th(num) -> float:
   """Rounds a value to the nearest 1/8"""
   return round(num * 8)/8

if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")
