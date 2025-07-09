# pylint: disable=invalid-name,bad-indentation
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 5/6/2025
# Update Date: 9/7/2025
# Make_Bent_Pipe_Note.py

"""This program generates a list of cut lengths for a unit.

When given a job number, this program will find the folder, 
read the bent pipe files, and generate a list of part numbers
and cut lengths.

Functions:
   main: Driver of the program
   findFolder: Finds the Bent Pipe Folder(s) in the Drawings Folder
   getLengths: Gets the list of tube lengths
   makeNote: Creates the .txt that holds all the information
   to8th: Rounds a float to the nearest 1/8"
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
from Get_Job_Pipe_Files import Get_Job_Pipe_Files
#pylint: enable=wrong-import-position

#Variables
path = fr"\\{Shared_Drive}\_A NTC GENERAL FILES\_JOB FILES\Job "
prtSource = fr"\\{Drawing_Drive}\_Drawings Pending Approval (PDF)\Bent Pipe"

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

   jobInst = Get_Job_Pipe_Files()
   pipeFolders = defaultdict(str)
   for bendFolder in bendFolders:
      drawingFolder = os.path.dirname(bendFolder)
      pdfs = glob.glob(drawingFolder + r"\*.PDF")
      jobInst.setFiles(pdfs)
      pipes = jobInst.getPipes()
      pipeFolders[bendFolder] = pipes
      print("Copying Pipe files to Folder...")
      flags = copyPrts(bendFolder, pipes.keys())
      print("Files Copied")
      if len(flags) > 0:
         print("There was an issue copying some of the pipe files\n"\
               "See the list below:")
         for flag in flags:
            print(flag)

   loader = Loader("Generating Pipe Note(s)...", "Note(s) generated\n", .1).start()
   allFlags = []
   for bendFolder in bendFolders:
      lengths = getlengths(bendFolder)
      flags = makeNote(lengths, bendFolder, pipeFolders)
      for flag in flags:
         allFlags.append(flag)
   loader.stop()

   if len(allFlags) != 0:
      print("The following files had Lengths that did not match:")
      for flag in allFlags:
         print(flag)
      print("Check the pipe in Bend Pro and run this simulation again")

def findFolder(folder) -> list:
   """Filters through the Mechanical - Drawings folder to find the Bent Pipe folder"""
   folders = []
   for root, directs, _ in os.walk(folder):
      for direct in directs:
         if "Bent Pipe" in direct:
            folders.append(os.path.join(root, direct))
   return folders

def copyPrts(bendFolder, pipes):
   """Copies the listed part files into the job folder"""
   flags = []
   for pipe in pipes:
      file = glob.glob(fr"{prtSource}\{pipe}.prt")
      if len(file) == 1:
         os.system(f'copy "{file[0]}" "{bendFolder}"')
      else:
         flags.append(pipe)
   return flags

def getlengths(folder) -> list:
   """Reads the date from each .prt file and returns a list"""
   pipes = []
   files = glob.iglob(folder + r'\*.prt')
   for file in files:
      partNumber = os.path.basename(file).split(".")[0]
      # data = open(file, encoding="utf-8")
      with open(file, encoding="utf-8") as data:
         content = data.readlines()
      length = to8th(float([s for s in content if "TubeLength=" in s][0].strip().split("=")[1]))
      note = [s for s in content if "Notes=" in s][0].strip().split("=")[1:]
      cutLength = to8th(float(note[1].strip().split('"')[0]))
      frontCut = to8th(float(note[2].strip().split('"')[0]))
      endCut = to8th(float(note[3].strip().split('"')[0]))
      pipes.append([partNumber, cutLength, frontCut, endCut, length == cutLength])
   return pipes

def makeNote(lengths, folder, pipeFolders) -> list:
   """Creates the formatted not in the job folder"""
   flags = []
   fileName = folder.split("\\")[-1:][0] + " List.txt"
   filePath = folder + "\\" + fileName
   pipes = pipeFolders[folder]
   with open(filePath,"x",encoding="utf-8") as f:
      f.write("Qty\tPart Number\t\tTotal Length\tFront Cut\tEnd Cut\n")
      for length in lengths:
         qty = pipes[length[0]]
         f.write(f"{qty}\t{length[0]}\t{length[1]}\t\t{length[2]}\t\t{length[3]}\n")
         if length[4] is False:
            flags.append(length[0])
   return flags

def to8th(num) -> float:
   """Rounds a value to the nearest 1/8"""
   return round(num * 8)/8

if __name__ == "__main__":
   main()
   input("Program completed. Press ENTER to close...")
