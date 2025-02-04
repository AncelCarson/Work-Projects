#pylint: disable = all,invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

import smtplib
import ssl

import os
import time
import getpass
import pandas as pd
from dotenv import load_dotenv

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

inputWorkbook = fr'\\{Shared_Drive}\Ancel\Sales\Job Followups\Job Followup 230426.xlsx'
# JetsonFile = fr'\\{Shared_Drive}\Ancel\Sales\Job Followups\Jetson Low Lead Times 220923.pdf'
# JetsonAttachment = open(JetsonFile, 'rb')
# TraneFile = fr'\\{Shared_Drive}\Ancel\Sales\Job Followups\CGWR, CCAR and CICD Low Lead Times 220923.pdf'
# TraneAttachment = open(TraneFile, 'rb')

def main():
   dfIn = pd.read_excel(inputWorkbook, sheet_name = 'Email List', header = 0)
   print(dfIn)
   emailList = dfIn[['EmailAddress','First Name','CUSTOMER_REFERENCE','QuoteNum','Dollars','Office','DaysLate']].copy()
   emailList['Jetson'] = False
   emailList.loc[emailList['Office'] == 'House (Do not use)', 'Jetson'] = True
   emailList.loc[emailList['Office'] == 'Independent- Non Trane', 'Jetson'] = True

   print(emailList)

   sendEmail(emailList)

def sendEmail(emails):
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   email, password = userLogin(server)
   server.login(email, password)

   # emails = [["acarson@nappstech.com",'Napps','Napps Test Chiller','1111','$32,566','Trane',95,False],
   #           ["acarson@nappstech.com",'Jetson','Jetson Test Chiller','1111','$41,041','Independant',56,True],]

   count = 0

   for index, row in emails.iterrows():
   # for row in emails:
      
      # FROM = "sales@nappstech.com"
      # TO = "acarson@nappstech.com"
      # name = 'Trevor'
      # job = 'HYOSUNG PROCESS CHILLER'
      # quote = '0003395'
      # price = '$32,566'
      # email = True

      TO = row[0]
      if row[7]:
         FROM = "sales@jetsonhvac.com"
         note = 'Jetson'
      else:
         FROM = "sales@nappstech.com"
         note = 'Napps'

      if row[6] >= 150:
         days = 150
      elif row[6] >= 120:
         days = 120
      elif row[6] >= 90:
         days = 90
      elif row[6] >= 60:
         days = 60
      else:
         days = 30

      text = getMessage(row[1], row[2], row[3], row[4], row[7], days) # Name, Job Name, Quote Number, Cost, Jetson (Yes/No), Days Late

      SUBJECT = '{0}: {1} Quote Follow Up'.format(row[3], row[2])

      msg = MIMEMultipart('alternative')
      msg['Subject'] = SUBJECT
      msg['From'] = FROM
      msg['To'] = TO
      
      # # Jetson File Notification
      # if note == 'Jetson':
      #    with open(JetsonFile, "rb") as JetsonAttachment:
      #       part = MIMEBase("application", "octet-stream")
      #       part.set_payload(JetsonAttachment.read())

      #    encoders.encode_base64(part)

      #    part.add_header("Content-Disposition", "JetsonAttachment", filename= "Jetson Low Lead Times 220923.pdf")
      #    msg.attach(part)

      # # Trane File Notification
      # if note == 'Napps':
      #    with open(TraneFile, "rb") as TraneAttachment:
      #       part = MIMEBase("application", "octet-stream")
      #       part.set_payload(TraneAttachment.read())

      #    encoders.encode_base64(part)

      #    part.add_header("Content-Disposition", "TraneAttachment", filename= "CGWR, CCAR and CICD Low Lead Times 220923.pdf")
      #    msg.attach(part)

      msg.attach(MIMEText(text, 'html'))

      # # Send Email and login after timeout
      # try:
      #    server.sendmail(FROM, TO, msg.as_string())
      # except SMTPRecipientsRefused:
      #    print("Connection Timed Out. Reconnecting...")
      #    server.login(email, password)
      #    print("Connection Restored")
      #    server.sendmail(FROM, TO, msg.as_string())

      print("{0} day {1} email sent to {2}: {3} Quote #{4}".format(days, note, row[1], row[0], row[3]))
      time.sleep(2)

   server.quit()

def getMessage(name, job, quote, price, email, days):
   body = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Hey {0},</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>You were provided a equipment quote on a project referenced as {1}. &nbsp;The quote number was {2} and it was for {3}.&nbsp;</p>
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
      """.format(name, job, quote, price)
   JetsonSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;line-height:105%;font-family:"Arial",sans-serif;color:#0055A0;'>Ancel Carson</span><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><br> <strong><span style="color:#FFC000;">Applications & Projects Engineer</span></strong></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>Napps Technology&nbsp;</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#FFC000;'>|</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;JETSON INNOVATIONS</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;</span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.nappstech.com/">www.nappstech.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/">www.jetsonhvac.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>Direct:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:903-758-2900"><span style="color:#00549F;">903-758-2900 Ext. 146</span></a></span></span></p> 
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)
   NappsSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Ancel Carson</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Applications & Projects Engineer</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Phone: 903-758-2900 X146</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Fax:&nbsp;903-758-2903</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Napps Technology Corporation</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>905 W. Cotton Street</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Longview, TX 75604</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><a href="www.nappstech.com">www.nappstech.com</a></span></p>
      <p><span style="font-size: 6px;">{}</span></p>""".format(days)

   if email:
      signature = JetsonSignature
   else:
      signature = NappsSignature

   return body + signature

def userLogin(server):
   email = input("What is your Napps Email?\n")
   password = getpass.getpass("What is your password? (this will not be saved)\n")
   try:
      server.login(email, password)
   except smtplib.SMTPAuthenticationError:
      print("\n!----------------------------!")
      print("Password entered is incorrect.")
      print("Please enter the email password for {}?".format(email))
      password = input("(Password will show for reference but will still not be saved)\n")
      try:
         server.login(email, password)
      except smtplib.SMTPAuthenticationError:
         print("\n!----------------------------!")
         print("Password entered is incorrect again")
         print("Please re-enter your email and try again.")
         userLogin(server)
   return [email,password]

if __name__ == "__main__":
   main()
   input("Program Completed. Press ENTER to Close...")
