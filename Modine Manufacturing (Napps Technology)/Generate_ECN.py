#pylint: disable = all,invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 23/8/2022
# Update Date: 16/4/2025
# Generate_ECN.py

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
import ssl
import glob
import smtplib
import getpass
import pandas as pd
import openpyxl as pyxl
from datetime import datetime
from dotenv import load_dotenv
from win32com.client import Dispatch

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from MenuMaker import makeMenu
#pylint: enable=wrong-import-position

#Variables
logFile = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Notices\ECN Log 231108.xlsx'
list_of_files = glob.glob(fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Engineering Change Notification Form*.xlsx') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)

#Functions
" Main Function "
def main():
   dfECN = pd.read_excel(logFile, index_col = 0)

   " Checking to See if Log File is in Use "
   try:
      file = open(logFile,"r+")
   except PermissionError as err:
      print("Permission Error: {0}".format(err))
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      input('Program Terminating. Press ENTER to Close...')
      return
   file.close()

   day = datetime.now().strftime('%m/%d/%Y')
   dfIn = pd.read_excel(latest_file, sheet_name = 'Lookup Tables')
   
   " Collecting User Data "
   userID = getUser(dfIn)
   priority = getPriority(dfIn)
   startDate = getStartDate(day)
   ECRdata = getECR()
   departmentData = getDepartments(dfIn)
   overview = input("Briefly describe the change or give it a name.\n")
   items = input('List the affected items separated by commas.\n')
   email = dfIn["Emails"][userID-1]
   userPass = getUserPass(email)

   " Setting up user defined variables "
   user = dfIn["User"][userID-1]
   ECRnum = ECRdata[0]
   ECRfolder = ECRdata[1]
   departments = departmentData[0]
   departmentIDs = departmentData[1]
   emailList = getEmails(departmentIDs, dfIn, email)
   for count in range(8-len(departments)):
      departments.append(" ")
   itemList = items.split(",")

   " Adding new row to ECN Log and saving file "
   dfNewECR = pd.concat([dfECN, pd.DataFrame([{"Overview" : overview, "Submitted By" : user, 
                         "Submission Date" : day, "Effective Date" : startDate, 
                         "Associated ECR" : ECRnum, "Affected Items" : items,
                         "Status" : "Incomplete"}])], ignore_index = True)
   
   " Opening Log File to Save new row "
   try:
      logWriter = pd.ExcelWriter(logFile)
   except PermissionError as err:
      print("Permission Error: {0}".format(err))
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      print("Answers given will be closed and not saved")
      input('Program Terminating. Press ENTER to Close...')
      return
   dfNewECR.to_excel(logWriter, sheet_name='Sheet1')
   logWriter.close()

   " Generating ECN File "
   requestID = dfNewECR.shape[0] - 1
   ECNfolder = createForm(requestID, user, priority, day, startDate, ECRnum, overview, departments, itemList)

   " Creating ECR Shortcut "
   if ECRfolder != 0:
      path = os.path.join(ECNfolder, 'ECR #{}.lnk'.format(ECRnum))
      shell = Dispatch("WScript.Shell")
      shortcut = shell.CreateShortCut(path)
      shortcut.Targetpath = ECRfolder
      shortcut.IconLocation = ECRfolder
      shortcut.save()

   " Sending Emails "
   if userPass != None:
      sendEmail(emailList, email, userPass, ECNfolder, requestID, overview, user)

   input("Program has completed. Press ENTER to close...")

def getUser(dfIn):
   makeMenu("Users", dfIn["User"].dropna())
   return int(input("Who is making the ECN?\n"))

def getPriority(dfIn):
   makeMenu("Priority", dfIn["ChangeType"].dropna())
   return int(input("What is the ECN Priority?\n"))

def getStartDate(day):
   selection = input('Will this change be implemented immediately? Y/N\n')
   if selection == "y" or selection == "Y":
      return day
   elif selection == "n" or selection == "N":
      return input("When will the changwe go into effect? mm/dd/yyyy\n")
   else:
      print("Invalid selection. Please Try again.\n")
      return getStartDate(day)

def getECR():
   selection = input('Is there an associated ECR with this Notification? Y/N\n')
   if selection == "y" or selection == "Y":
      ECRnum = input("What is the associated ECR #?\n")
      return [ECRnum, fr"\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Requests\Completed\Engineering Change Request {ECRnum}"]
   elif selection == "n" or selection == "N":
      return ["N/A",0]
   else:
      print("Invalid selection. Please Try again.\n")
      return getECR()

def getDepartments(dfIn):
   departments = []
   makeMenu("Departments", dfIn["Department"].dropna())
   selection = input("List Departments with Action Items separated by commas\n")
   selections = selection.split(",")
   departmentIDs = [int(value) for value in selections]
   for item in departmentIDs:
      departments.append(dfIn["Department"][item-1])
   print(departments)
   return [departments, departmentIDs]

def getUserPass(email):
   selection = input('Does an email notification need to be sent? Y/N\n')
   if selection == "y" or selection == "Y":
      return getpass.getpass("Please enter the email password for {}?\n".format(email))
   elif selection == "n" or selection == "N":
      return None
   else:
      print("Invalid selection. Please Try again.\n")
      return getUserPass(email)

def getEmails(departments, dfIn, email):
   longList = []
   shortList = []
   for item in departments:
      longList += dfIn["EmailList"][item-1].split(",")
   [shortList.append(x) for x in longList if x not in shortList]
   if email in shortList:
      shortList.remove(email)
   return shortList

def createForm(requestID, user, priority, day, startDate, ECRnum, overview, departments, itemList):
   workBook = pyxl.load_workbook(latest_file)
   departmentCells = ['C31','C33','C35','C37','C39','C41','C43','C45']
   workSheet = workBook.active

   " Filling out Header" 
   workSheet['B2'].value = "ENGINEERING CHANGE REQUEST " + str(requestID)
   workSheet['F6'].value = user
   workSheet['P6'].value = priority
   workSheet['F8'].value = day
   workSheet['P8'].value = startDate
   workSheet['Y8'].value = ECRnum
   workSheet['F10'].value = overview

   " Setting Notified Departments "
   for count in range(len(departmentCells)):
      workSheet[departmentCells[count]].value = departments[count]

   " Listing Affective Parts "
   for count in range(len(itemList)):
      workSheet["C" + str(52 + count)].value = itemList[count]

   filePath = createFile(requestID)
   workBook.save(filePath[0])
   os.startfile(filePath[0])
   return filePath[1]

def createFile(requestID):
   day = datetime.now().strftime('%y%m%d')
   ECNFolder = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Notices\Engineering Change Notification {requestID}'
   ECNFile = fr'\\{Shared_Drive}\Engineering Change Requests (ECR)\Change Notices\Engineering Change Notification {requestID}\Engineering Change Notification {requestID} {day}.xlsx'
   try:
      os.system('mkdir "{}"'.format(ECNFolder))
      os.system('copy "{}" "{}"'.format(latest_file, ECNFile))
   except OSError as eer:
      print("OS Error: {0}".format(eer))
      print("A File was not created.")
      return
   return [ECNFile, ECNFolder]

def sendEmail(emailList, email, userPass, ECNfolder, requestID, overview, user):
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   try:
      server.login(email, userPass)
   except smtplib.SMTPAuthenticationError:
      print("\n!----------------------------!")
      print("Password entered is incorrect.")
      userPass = getpass.getpass("Please enter the email password for {}?\n".format(email))
      sendEmail(emailList, email, userPass, ECNfolder, requestID, overview, user)
      return

   TO = ", ".join(emailList)
   FROM = email
   SUBJECT = "ECN #{0}: {1}".format(requestID, overview)

   text = """
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Engineering Change Notification #{0} has been submitted. Please complete the necessary Action Items to complete the Notice.</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>The ECN can be found <a href="{1}">here</a>.</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Please reply all to this email thread once your portion has been completed.</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>-{2}</p>
   """.format(requestID, ECNfolder, user)

   msg = MIMEMultipart('alternative')
   msg['Subject'] = SUBJECT
   msg['From'] = FROM
   msg['To'] = TO

   msg.attach(MIMEText(text, 'html'))

   server.sendmail(FROM, emailList, msg.as_string())
   print(SUBJECT + " has sent to the following addresses: {}".format(emailList))

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()