#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 15/5/2023
# Update Date: 12/6/2024
# RepSendEmail.py

"""This Program sends out formatted emails to a list of sels reps with open jobs.

This program is to be run after running RepFollowupSheets.py. It reads the Email
List sheet and sends 1 email per row. For each email, the body of text is changed
to reflect a specific job before it is sent. 

Functions:
   main: Driver of the program
   sendEmail: Sends an email for each job from the data file
   getMessage: Builds the body of the email with information specific to the recipient
"""

#Libraries
import smtplib
import ssl

import os
import glob
import time
import pandas as pd
from dotenv import load_dotenv


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Secret Variables
load_dotenv()
AEmail = os.getenv('AEmail')  #Saved User Email 
APassW = os.getenv('APassW')  #Saved User Password 

#Variables
file_path = r'S:\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Follow Up Data\*.xlsx' # * means all if need specific format then *.xlsx
files = sorted(glob.iglob(file_path), key=os.path.getctime, reverse=True)
inputWorkbook = files[0]

def main():
   """Generates a dateframe based off of the given data"""
   dfIn = pd.read_excel(inputWorkbook, sheet_name = 'Email List', header = 0)
   print(dfIn)
   emailList = dfIn[['EmailAddress','First Name','CUSTOMER_REFERENCE','QuoteNum','Dollars','DaysLate','Sender','Company','Units']].copy()
   emailList['Jetson'] = False
   emailList.loc[emailList['Company'] == 'Jetson', 'Jetson'] = True

   print(emailList)

   print("Pulling email list from {0}.".format(inputWorkbook))

   sendEmail(emailList)

def sendEmail(emails):
   """Takes the dataframe and sends the email
   
   Parameters:
      emails (dateframe): The list if jobs with emails, names, and other specific information
   """
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   server.login(AEmail, APassW)

   #The test is intended to make sure the emails will send without sending emails to the reps. 
   test = input("Is this a test run? (Y/N)\n")
   if test == "n" or test == "N":
      pass
   else:
      emails = [["acarson@nappstech.com",'Napps','Napps Anas Test Chiller','1111','$32,566',95,'Anas','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson Anas Test Chiller','1111','$41,041',56,'Anas','Jetson',' for (1) 40Ton ACCS',True],
               ["acarson@nappstech.com",'Napps','Napps Preston Test Chiller','1111','$32,566',95,'Preston','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson Preston Test Chiller','1111','$41,041',56,'Preston','Jetson',' for (1) 40Ton ACCS',True],
               ["acarson@nappstech.com",'Napps','Napps General Test Chiller','1111','$32,566',95,'None','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson General Test Chiller','1111','$41,041',56,'None','Jetson',' for (1) 40Ton ACCS',True],]
      emails = pd.DataFrame(emails, columns = ['EmailAddress','First Name','CUSTOMER_REFERENCE','QuoteNum','Dollars','DaysLate','Sender','Company','Units','Jetson'])

   #This test is intended to check that the program is reading the file correctly before sending emails.
   sendIt = False
   send = input("Do you want to send the emails? (Y/N)\n")
   if send == "y" or send == "Y":
      sendIt = True
   else:
      sendIt = False #Yes, this is redundant. Erroneously sending emails scares me so I am leaving it in.

   for index, row in emails.iterrows():

      TO = row['EmailAddress']
      note = row['Company']
      if row['Jetson']:
         FROM = "sales@jetsonhvac.com"
      else:
         FROM = "sales@nappstech.com"

      if row['DaysLate'] >= 150:
         days = 150
      elif row['DaysLate'] >= 120:
         days = 120
      elif row['DaysLate'] >= 90:
         days = 90
      elif row['DaysLate'] >= 60:
         days = 60
      else:
         days = 30

      text = getMessage(row['First Name'], row['CUSTOMER_REFERENCE'], row['QuoteNum'], row['Dollars'], row['Sender'], row['Units'], row['Jetson'], days) # Name, Job Name, Quote Number, Cost, Sender Name, Quotes Units, Jetson (Yes/No), Days Late

      SUBJECT = '{0}: {1} Quote Follow Up'.format(row['QuoteNum'], row['CUSTOMER_REFERENCE'])

      msg = MIMEMultipart('alternative')
      msg['Subject'] = SUBJECT
      msg['From'] = FROM
      msg['To'] = TO

      msg.attach(MIMEText(text, 'html'))

      if sendIt:
         # Send Email and login after timeout
         try:
            server.sendmail(FROM, TO, msg.as_string())
         except SMTPRecipientsRefused:
            print("Connection Timed Out. Reconnecting...")
            server.login(AEmail, APassW)
            print("Connection Restored")
            server.sendmail(FROM, TO, msg.as_string())
         except Exception as e:
            print("An error has occured. The email for {0} was not sent.".format(row['QuoteNum']))
            print("Update the Follow Up Data Workbook before running this program again.")
            print("Delete all jobs in the rows above {0} and save the file.".format(row['QuoteNum']))
            input("Program Terminating. Press ENTER to Close...")
            return

      print("{0} day {1} email sent to {2}: {3} Quote #{4}".format(days, note, row['First Name'], row['EmailAddress'], row['QuoteNum']))
      time.sleep(2)

   server.quit()

def getMessage(name, job, quote, price, sender, units, Jetson, days):
   """Returns a string of HTLM code with user specific data to be the body of the email
   
   Parameters:
      name (str): The first name of the email recipient
      job (str): The Job Name
      quote (str): The Job Number
      price (str): The Price of the Job
      sender (str): The inside Sales person who is responsible for the email reciepient
      units (str): The Description of units on the job
      Jetson (bool): Wether or not this email is from the Jetson Inbox
      days (str): How many days it has been since the last follow up. 

   Returns:
      body (str) + signature (str): The body of the emailwith specific data
   """
   body = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Hey {0},</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>You were provided a equipment quote on a project referenced as {1}{4}. &nbsp;The quote number was {2} and it was for {3}.&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>By <u><span style="background:yellow;">HIGHLIGHTING</span></u> the appropriate answer, please let me know if this project:</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Has Been Won&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Production Release in &nbsp; &nbsp; &nbsp; &nbsp;30 days &nbsp; &nbsp; &nbsp; &nbsp;60 day &nbsp; &nbsp; &nbsp; &nbsp;90 days &nbsp; &nbsp; &nbsp; &nbsp;Unknown &nbsp;&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Was Abandoned (project did not move forward or decided to repair instead of replace)&nbsp; &nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Has Been Lost to a Competitor</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>ArticChill &nbsp; &nbsp; &nbsp; &nbsp;Carrier &nbsp; &nbsp; &nbsp; &nbsp;Climacool &nbsp; &nbsp; &nbsp; &nbsp;Daikin &nbsp; &nbsp; &nbsp; &nbsp;York/JCI &nbsp; &nbsp; &nbsp; &nbsp;MultiStack &nbsp; &nbsp; &nbsp; &nbsp;Other&nbsp; &nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>If lost to a competitor it would help a lot to know the reason why. &nbsp;Please <span style="background:yellow;">HIGHLIGHT</span> the reason.</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Price &nbsp; &nbsp; &nbsp; &nbsp;Lead Time &nbsp; &nbsp; &nbsp; &nbsp;Product or Option missing &nbsp; &nbsp; &nbsp; &nbsp;Footprint/Size &nbsp; &nbsp; &nbsp; &nbsp;Efficiency &nbsp; &nbsp; &nbsp; &nbsp;Other&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      """.format(name, job, quote, price, units)
   PrestonJetson = """
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Preston Ware</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Regional Sales Manager<strong>&nbsp;&ndash;&nbsp;</strong>Commercial Industrial</span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Modine Manufacturing Company</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>P:</span></strong><span style="font-size:13px;color:#44546A;">&nbsp;</span><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="tel:903-758-2900"><span style="color:#0563C1;">903-758-2900 Ext. 145</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/"><span style="color:#0563C1;">www.JetsonHVAC.com</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Email:</span><span style='font-size:13px;font-family:"Arial",sans-serif;color:red;'>&nbsp;</span><u><span style='font-size:13px;font-family:  "Arial",sans-serif;color:#0563C1;'><a href="mailto:Sales@JetsonHVAC.com"><span style="color:#0563C1;">Sales@JetsonHVAC.com</span></a></span></u></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   AnasJetson = """
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Anas Shahid</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Inside Sales Engineer<strong>&nbsp;&ndash;&nbsp;</strong>Commercial Industrial</span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Modine Manufacturing Company</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>P:</span></strong><span style="font-size:13px;color:#44546A;">&nbsp;</span><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="tel:903-758-2900"><span style="color:#0563C1;">903-758-2900 Ext. 142</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/"><span style="color:#0563C1;">www.JetsonHVAC.com</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Email:</span><span style='font-size:13px;font-family:"Arial",sans-serif;color:red;'>&nbsp;</span><u><span style='font-size:13px;font-family:  "Arial",sans-serif;color:#0563C1;'><a href="mailto:Sales@JetsonHVAC.com"><span style="color:#0563C1;">Sales@JetsonHVAC.com</span></a></span></u></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   JetsonSignature = """
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Jetson Sales Team</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'></strong>Commercial Industrial</span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Modine Manufacturing Company</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>P:</span></strong><span style="font-size:13px;color:#44546A;">&nbsp;</span><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="tel:903-758-2900"><span style="color:#0563C1;">903-758-2900</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/"><span style="color:#0563C1;">www.JetsonHVAC.com</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Email:</span><span style='font-size:13px;font-family:"Arial",sans-serif;color:red;'>&nbsp;</span><u><span style='font-size:13px;font-family:  "Arial",sans-serif;color:#0563C1;'><a href="mailto:Sales@JetsonHVAC.com"><span style="color:#0563C1;">Sales@JetsonHVAC.com</span></a></span></u></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   PrestonNapps = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Preston Ware</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Inside Sales Engineer</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Phone: 903-758-2900 X145</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Fax:&nbsp;903-758-2903</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Napps Technology Corporation</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>905 W. Cotton Street</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Longview, TX 75604</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><a href="www.nappstech.com">www.nappstech.com</a></span></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   AnasNapps = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Anas Shahid</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Inside Sales Engineer</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Phone: 903-758-2900 X142</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Fax:&nbsp;903-758-2903</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Napps Technology Corporation</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>905 W. Cotton Street</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Longview, TX 75604</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><a href="www.nappstech.com">www.nappstech.com</a></span></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   NappsSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Napps Sales Team</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Phone: 903-758-2900 X142</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Fax:&nbsp;903-758-2903</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Napps Technology Corporation</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>905 W. Cotton Street</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Longview, TX 75604</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><a href="www.nappstech.com">www.nappstech.com</a></span></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)

   if Jetson:
      # if sender == "Preston":
      #    signature = PrestonJetson
      # elif sender == "Anas":
      #    signature = AnasJetson
      # else:
      #    signature = JetsonSignature
      signature = AnasJetson
   else:
      # if sender == "Preston":
      #    signature = PrestonNapps
      # elif sender == "Anas":
      #    signature = AnasNapps
      # else:
      #    signature = NappsSignature
      signature = PrestonNapps

   return body + signature

if __name__ == "__main__":
   main()
   input("Program Completed. Press ENTER to Close...")