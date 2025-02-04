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

inputWorkbook = fr'\\{Shared_Drive}\Ancel\Pricing\JESS Announcements\Job Followup 221117.xlsx'
filename = "Water-Cooled Chillers Released in JESS 230117a.pdf"
JetsonFile = fr'\\{Shared_Drive}\Ancel\Pricing\JESS Announcements\Water-Cooled Chillers Released in JESS 230117a.pdf'
JetsonAttachment = open(JetsonFile, 'rb')
# TraneFile = fr'\\{Shared_Drive}\Ancel\Sales\Job Followups\CGWR, CCAR and CICD Low Lead Times 220923.pdf'
# TraneAttachment = open(TraneFile, 'rb')

def main():
   dfIn = pd.read_excel(inputWorkbook, sheet_name = 'Email List', header = 0)
   print(dfIn)
   emailList = dfIn[['EmailAddress','First Name','Office',]].copy()
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

   # emails = [["acarson@nappstech.com",'Napps','Trane',False],
   #           ["acarson@nappstech.com",'Jetson','Independant',True],]

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
      if row[3]:
         FROM = "sales@jetsonhvac.com"
         note = 'Jetson'
         SUBJECT = 'Jetson Announcement'
      else:
         FROM = "sales@nappstech.com"
         note = 'Napps'
         SUBJECT = 'Napps Technology Announcement'
         print("Trane Email Skipped")
         continue

      text = getMessage(row[1], row[3])

      

      msg = MIMEMultipart('alternative')
      msg['Subject'] = SUBJECT
      msg['From'] = FROM
      msg['To'] = TO
      
      # Jetson File Notification
      if note == 'Jetson':
         with open(JetsonFile, "rb") as JetsonAttachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(JetsonAttachment.read())

         encoders.encode_base64(part)

         part.add_header("Content-Disposition", "JetsonAttachment", filename= filename)
         msg.attach(part)

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

      print("{0} email sent to {1}".format(note, row[0]))
      time.sleep(2)

   server.quit()

def getMessage(name, email):
   body = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Hey {0},</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Attached is a letter from the president regarding updated to our Jetson Engineering Selection Software (JESS).&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Let us know if you have any questions about the announcement, running selections, or getting into the software.&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Have a great rest of your day,&nbsp;</p>
      """.format(name)
   JetsonSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;line-height:105%;font-family:"Arial",sans-serif;color:#0055A0;'>Ancel Carson</span><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><br> <strong><span style="color:#FFC000;">Applications & Projects Engineer</span></strong></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>Napps Technology&nbsp;</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#FFC000;'>|</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;JETSON INNOVATIONS</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;</span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.nappstech.com/">www.nappstech.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/">www.jetsonhvac.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>Direct:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:903-758-2900"><span style="color:#00549F;">903-758-2900 Ext. 146</span></a></span></span></p>
      """
   NappsSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><strong>Ancel Carson</strong></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Applications & Projects Engineer</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Phone: 903-758-2900 X146</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Fax:&nbsp;903-758-2903</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Napps Technology Corporation</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>905 W. Cotton Street</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><u>Longview, TX 75604</u></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><a href="www.nappstech.com">www.nappstech.com</a></span></p>
      """

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
