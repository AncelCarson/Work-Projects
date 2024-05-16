# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 15/5/2024
# Generate_ECR.py

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
import ssl
import glob
import smtplib
import pandas as pd
import openpyxl as pyxl
from datetime import datetime
from dotenv import load_dotenv

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#Secret Variables
load_dotenv()
AEmail = os.getenv('AEmail')
APassW = os.getenv('APassW')

#Variables
logFile = r'S:\Engineering Change Requests (ECR)\Change Requests\ECR Log 231121.xlsx'
list_of_files = glob.glob(r'S:\Engineering Change Requests (ECR)\Engineering Change Request Form*.xlsx') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)

#Functions
" Main Finction "
def main():
   dfECR = pd.read_excel(logFile, index_col = 0)

   " Checking to See if Log File is in Use "
   file = fileCheck(logFile)
   if file == 0:
      return
   file.close()

   day = datetime.now().strftime('%m/%d/%Y')
   user = input("Who is requesting the ECR? (Ex. Ancel)\n")
   userCode = input("Please Enter your Initials? (Ex. AC)\n")
   product = input("What Unit type will be affected? (Ex. CCAR-BASE-050)\n")
   changeType = changeMenu()
   affectedItems = input('What are the affected Parts or Assemblies?\n')
   workOrder = input('What is the affected work order? (Ex. 3476A)\n')
   dfNewECR = pd.concat([dfECR, pd.DataFrame([{"Requested By" : user, "Product Line" : product, "Type of Change" : changeType,
                         "Part/Subassemblies Affected" : affectedItems, "Work Order Number" : workOrder,
                         "Date of Request" : day}])], ignore_index = True)
   requestID = dfNewECR.shape[0] - 1
   createForm(requestID, day, workOrder, product, affectedItems, user, userCode)

   " Checking to See if Log File is still not in Use "
   file = fileCheck(logFile)
   if file == 0:
      return
   file.close()

   " Opening Log File to Save new row "
   logWriter = pd.ExcelWriter(logFile)
   dfNewECR.to_excel(logWriter, sheet_name='Sheet1')
   logWriter.close()
   input('Program Completed. Press ENTER to Close...')

def lineCall(contact, unit, parts, workOrder):
   dfECR = pd.read_excel(logFile, index_col = 0)

   " Checking to See if Log File is in Use "
   file = fileCheck(logFile)
   if file == 0:
      return "Need ECR"
   file.close()

   day = datetime.now().strftime('%m/%d/%Y')
   user = contact
   affectedItems = parts

   userCode = input("Please Enter your Initials? (Ex. AC)\n")
   product = input("What is the capacity of the unit with the issue? (Ex. 50 Ton)\n") + " " + unit
   changeType = changeMenu()
   dfNewECR = pd.concat([dfECR, pd.DataFrame([{"Requested By" : user, "Product Line" : product, "Type of Change" : changeType,
                         "Part/Subassemblies Affected" : affectedItems, "Work Order Number" : workOrder,
                         "Date of Request" : day}])], ignore_index = True)
   requestID = dfNewECR.shape[0] - 1
   createForm(requestID, day, workOrder, product, affectedItems, user, userCode)

   " Checking to See if Log File is still not in Use "
   file = fileCheck(logFile)
   if file == 0:
      return str(requestID)
   file.close()

   " Opening Log File to Save new row "
   logWriter = pd.ExcelWriter(logFile)
   dfNewECR.to_excel(logWriter, sheet_name='Sheet1')
   logWriter.close()
   print("\n!-- ECR Log Updated with ECR #{} --!\n".format(requestID))
   return str(requestID)

def fileCheck(logFile):
   try:
      file = open(logFile,"r+")
   except PermissionError as err:
      print("Permission Error: {0}".format(err))
      print("The file is open by another user")
      print("Opening Log File...")
      os.startfile(logFile)
      print("Ask user to close the Log file then run the program again")
      input('Program Terminating. Press ENTER to Close...')
      return 0 
   return file

def changeMenu():
   print('|   Change Type    |')
   print('|------------------|')
   print('|1: BOM Change     |')
   print('|2: Drawing Change |')
   print('|3: Both           |')
   print('|4: Other          |')
   print('|------------------|')
   selection = int(input('What is the Change Type?\n'))
   if selection == 1:
      return "BOM Change"
   elif selection == 2:
      return "Drawing Change"
   elif selection == 3:
      return "BOM and Drawing Change"
   elif selection == 4:
      return input("Input change type.\n")
   else:
      print('Incorrect selection\n\n')
      return changeMenu()

def createForm(requestID, day, workOrder, product, affectedItems, user, userCode):
   workBook = pyxl.load_workbook(latest_file)
   workSheet = workBook.active
   workSheet['B2'].value = "ENGINEERING CHANGE REQUEST " + str(requestID)
   workSheet['G7'].value = day
   workSheet['N38'].value = day
   workSheet['V7'].value = workOrder
   workSheet['J14'].value = product
   workSheet['J16'].value = affectedItems
   workSheet['D38'].value = user
   filePath = createFile(requestID, userCode)
   workBook.save(filePath)
   os.startfile(filePath)


def createFile(requestID, userCode):
   day = datetime.now().strftime('%y%m%d')
   ECRFolder = 'S:\Engineering Change Requests (ECR)\Change Requests\Engineering Change Request {0}'.format(requestID)
   ECRFile = 'S:\Engineering Change Requests (ECR)\Change Requests\Engineering Change Request {0}\Engineering Change Request {0} {1} {2}.xlsx'.format(requestID, day, userCode)
   try:
      os.system('mkdir "{}"'.format(ECRFolder))
      os.system('copy "{}" "{}"'.format(latest_file, ECRFile))
   except OSError as eer:
      print("OS Error: {0}".format(eer))
      print("A File was not created.")
      return
   sendEmail(ECRFolder, requestID)
   return ECRFile

def sendEmail(ECRFolder, requestID):
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   server.login(AEmail, APassW)

   TO = "acarson@nappstech.com, matthew.m.bevan@modine.com"
   FROM = "techsupport@nappstech.com"
   SUBJECT = "ECR #{0}".format(requestID)

   text = """
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Engineering Change Request #{0} has been submitted. Please review the request and notify the appropriate individual.</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>The ECR can be found <a href="{1}">here</a>.</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
   <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>-Engineering</p>
   """.format(requestID, ECRFolder)

   msg = MIMEMultipart('alternative')
   msg['Subject'] = SUBJECT
   msg['From'] = FROM
   msg['To'] = TO

   msg.attach(MIMEText(text, 'html'))

   server.sendmail(FROM, TO, msg.as_string())
   print(SUBJECT + " has been sent to the Engineering Manager")

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()
