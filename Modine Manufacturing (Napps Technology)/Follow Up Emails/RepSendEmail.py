#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 15/5/2023
# Update Date: 15/5/2024
# RepSendEmail.py

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
AEmail = os.getenv('AEmail')
APassW = os.getenv('APassW')

#Variables
inputWorkbook = r'S:\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Follow Up Data\Job Followup 240415.xlsx'

def main():
   dfIn = pd.read_excel(inputWorkbook, sheet_name = 'Email List', header = 0)
   print(dfIn)
   emailList = dfIn[['EmailAddress','First Name','CUSTOMER_REFERENCE','QuoteNum','Dollars','DaysLate','Sender','Company','Units']].copy()
   emailList['Jetson'] = False
   emailList.loc[emailList['Company'] == 'Jetson', 'Jetson'] = True

   print(emailList)

   sendEmail(emailList)

def sendEmail(emails):
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   server.login(AEmail, APassW)

   test = input("Is this a test run? (Y/N)\n")
   if test == "n" or test == "N":
      pass
   else:
      emails = [["acarson@nappstech.com",'Napps','Napps Anas Test Chiller','1111','$32,566',95,'Anas','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson Anas Test Chiller','1111','$41,041',56,'Anas','Jetson',' for (1) 40Ton ACCS',True],
               ["acarson@nappstech.com",'Napps','Napps Tom Test Chiller','1111','$32,566',95,'Tom','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson Tom Test Chiller','1111','$41,041',56,'Tom','Jetson',' for (1) 40Ton ACCS',True],
               ["acarson@nappstech.com",'Napps','Napps General Test Chiller','1111','$32,566',95,'None','Napps',' for (1) 40Ton CGWR',False],
               ["acarson@nappstech.com",'Jetson','Jetson General Test Chiller','1111','$41,041',56,'None','Jetson',' for (1) 40Ton ACCS',True],]
      emails = pd.DataFrame(emails, columns = ['EmailAddress','First Name','CUSTOMER_REFERENCE','QuoteNum','Dollars','DaysLate','Sender','Company','Units','Jetson'])

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

      # # Send Email and login after timeout
      # try:
      #    server.sendmail(FROM, TO, msg.as_string())
      # except SMTPRecipientsRefused:
      #    print("Connection Timed Out. Reconnecting...")
      #    server.login(AEmail, APassW)
      #    print("Connection Restored")
      #    server.sendmail(FROM, TO, msg.as_string())

      print("{0} day {1} email sent to {2}: {3} Quote #{4}".format(days, note, row['First Name'], row['EmailAddress'], row['QuoteNum']))
      time.sleep(2)

   server.quit()

def getMessage(name, job, quote, price, sender, units, Jetson, days):
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
   TomJetson = """
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Tom Armstrong</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;font-family:"Arial",sans-serif;'>Regional Sales Manager<strong>&nbsp;&ndash;&nbsp;</strong>Commercial Industrial</span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>Modine Manufacturing Company</span></strong></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>D:</span></strong><span style="font-size:13px;color:#44546A;">&nbsp;</span><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="tel:951-389-4741"><span style="color:#0563C1;">951-389-4741</span></a></span></p>
      <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:13px;font-family:"Arial",sans-serif;color:#003764;'>C:</span></strong><span style="font-size:13px;color:#44546A;">&nbsp;</span><span style='font-size:13px;font-family:"Arial",sans-serif;'><a href="tel:951-544-6283"><span style="color:#0563C1;">951-544-6283</span></a></span></p>
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
   TomNapps = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Tom Armstrong</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Regional Sales Manager</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>D: 951-389-4741</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>C: 951-544-6283</p>
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
      if sender == "Tom":
         signature = TomJetson
      elif sender == "Anas":
         signature = AnasJetson
      else:
         signature = JetsonSignature
   else:
      if sender == "Tom":
         signature = TomNapps
      elif sender == "Anas":
         signature = AnasNapps
      else:
         signature = NappsSignature

   return body + signature

if __name__ == "__main__":
   main()
   input("Program Completed. Press ENTER to Close...")