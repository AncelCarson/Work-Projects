# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 22/5/23
# Update Date: 31/7/24
# JobFolderGenerator.py

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
import glob

#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
from GetDrawings import GetDrawings

#Variables
documentPath = r'S:\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Documents\Order Process Sheets\\'
standardSheet = r"Standard Order Process Checklist.xlsx"
specialSheet = r"Design Special Order Process Checklist.xlsx"
releaseName = r"_Job Release Checklist"
jobsFolder = r'S:\_A NTC GENERAL FILES\_JOB FILES'
quotesFolder = r'S:\_A NTC GENERAL FILES\_Quotes & Misc Info\QUOTES\Quotes *\**\Quote *.pdf'
subFolders = [r"\Config",r"\Drawings - Electrical",r"\Drawings - Mechanical",r"\Emails",r"\Pics",r"\Shortage Reports",r"\Submittal"]

#Functions
" Main Finction "
def main():
   quoteFilePath = getquote()
   [jobName,jobNumber,quoteFile] = getJob(quoteFilePath)
   folderPath = jobsFolder + jobName
   quotePath = folderPath + "\\" + quoteFile
   os.system('mkdir "{}"'.format(folderPath))
   os.system('copy "{}" "{}"'.format(quoteFilePath, quotePath))

   for folder in subFolders:
      os.system('mkdir "{}"'.format(folderPath + folder))
      print("{} folder made".format(folder))

   releaseSheet = documentPath + releaseName + ".xlsx"
   releasePath = folderPath + "\\" + releaseName + " " + jobNumber + ".xlsx"
   os.system('copy "{}" "{}"'.format(releaseSheet, releasePath))

   choice = input("Is this a design Special Job? Y/N\n").upper()
   if choice == "Y":
      processSheet = documentPath + specialSheet
      sheetPath = folderPath + "\\" + jobNumber + " " + specialSheet
   else:
      processSheet = documentPath + standardSheet
      sheetPath = folderPath + "\\" + jobNumber + " " + standardSheet
      choice = input("Do you want to pull the drawings for this job? Y/N\n").upper()
      if choice == "Y":
         drawings = GetDrawings()
         drawings.wo = jobNumber
         drawings.findUnit()
         drawings.collect()

   os.system('copy "{}" "{}"'.format(processSheet, sheetPath))
   os.startfile(folderPath)


def getquote():
   quotes = glob.glob(quotesFolder, recursive = True)
   quote = input("What quote is being used to make the Sales Order?\n").upper()
   quoteFile = None
   for i, s in enumerate(quotes):
      if quote in s:
         quoteFile = quotes[i]
         break 
   if quoteFile == None:
      print("\n!-----------------------------------")
      print("No quote was found with that number.")
      print("Make sure the referenced quote is saved in the quotes folders and try again.")
      print("!-----------------------------------\n")
      return getquote()
   print("\nIs the following the correct quote?")
   print(quoteFile)
   choice = input("Y/N\n").upper()
   if choice == "Y":
      return quoteFile
   else:
      return getquote()

def getJob(quoteFile):
   folders = quoteFile.split("\\")
   quoteFile = folders[len(folders)-1]
   fileParts = quoteFile[6:-4].split(" ")
   del fileParts[-3:]
   if len(fileParts[0]) == 4:
      fileParts[0] += "A-01"
   else:
      fileParts[0] += "-01"
   jobFolder = "\\Job"
   for part in fileParts:
      jobFolder += " {}".format(part)
   print("When was this order recieced?")
   jobFolder += " " + input("Format answer as follows: M D YY. ex: 5 27 23\n")
   print()
   print(jobFolder)
   return [jobFolder, fileParts[0], quoteFile]

" Checks if this program is being called "
if __name__ == "__main__":
   main()