#-*- coding: utf-8 -*-

import smtplib
import ssl

import os
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

inputWorkbook = r'S:\Ancel\Pricing\JESS Announcements\TA Email List_R5.xlsx'

def main():
   dfIn = pd.read_excel(inputWorkbook, sheet_name = 'Email List', header = 0)
   print(dfIn)
   emailList = dfIn[['Email','First Name','Company','Send','East/West']].copy()
   emailList['Jetson'] = True

   emailList = emailList[emailList.Send != 'N']
   emailList = emailList.iloc[:300]

   print(emailList)

   sendEmail(emailList)

def sendEmail(emails):
   SERVER = "smtp.office365.com"

   server = smtplib.SMTP(host = SERVER, port = 587)
   context = ssl.create_default_context()    
   server.starttls(context=context)
   server.login(AEmail, APassW)

   # emails = [["acarson@nappstech.com",'Napps','Trane',False],
   #           ["acarson@nappstech.com",'Jetson','Independant',True],
   #           ["tarmstrong@nappstech.com",'Jetson','Independant',True],]

   count = 0

   for index, row in emails.iterrows():
   # for row in emails:

      TO = row[0]
      if row[3]:
         FROM = "sales@jetsonhvac.com"
         note = 'Jetson'
         SUBJECT = 'Lead Time Update- Jetson Chillers'
      else:
         FROM = "sales@nappstech.com"
         note = 'Napps'
         SUBJECT = 'Napps Technology Announcement'
         print("Trane Email Skipped")
         continue

      text = getMessage(row[1], 'W')

      

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

      #    part.add_header("Content-Disposition", "JetsonAttachment", filename= filename)
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
      #    server.login("Acarson@nappstech.com", "!cebergTARDISna1")
      #    print("Connection Restored")
      #    server.sendmail(FROM, TO, msg.as_string())

      print("{0} email sent to {1}".format(note, row[0]))
      time.sleep(2)

   server.quit()

def getMessage(name, region):

   body = """
      <table border="0" cellspacing="0" cellpadding="0"><tbody><tr><td valign="top" width="150">
      <p><strong>Current Lead Time:</strong></p></td><td valign="top" width="138">
      <p>Air Cooled:</p></td><td valign="top" width="96">
      <p><strong>18 Weeks</strong></p></td><td valign="top" width="343">
      <p>build time from final PO; agreed bill of materials.</p></td></tr><tr><td valign="top" width="150">
      <p>&nbsp;</p></td><td valign="top" width="138">
      <p>Water Cooled<strong>:</strong></p></td><td valign="top" width="96">
      <p><strong>18 Weeks</strong></p></td><td valign="top" width="343">
      <p><strong>&nbsp;</strong></p></td></tr><tr><td valign="top" width="150">
      <p>&nbsp;</p></td><td valign="top" width="138">
      <p>Heat Recovery:</p></td><td valign="top" width="96">
      <p><strong>18 Weeks</strong></p></td><td valign="top" width="343">
      <p><strong>&nbsp;</strong></p></td></tr><tr><td valign="top" width="150">
      <p>&nbsp;</p></td><td valign="top" width="138">
      <p>Heat Pump:</p></td><td valign="top" width="96">
      <p><strong>30 Weeks</strong></p></td><td valign="top" width="343">
      <p><strong>&nbsp;</strong></p></td></tr><tr><td valign="top" width="150">
      <p>&nbsp;</p></td><td valign="top" width="138">
      <p>With Shell &amp; Tube:</p></td><td valign="top" width="96">
      <p><strong>22-24 Weeks</strong></p></td><td valign="top" width="343">
      <p><strong>&nbsp;</strong></p></td></tr><tr><td valign="top" width="150">
      <p>&nbsp;</p></td><td valign="top" width="138">
      <p>With Pump Skids:</p></td><td valign="top" width="96">
      <p><strong>22-26 Weeks</strong></p></td><td valign="top" width="343">
      <p><strong>&nbsp;</strong></p></td></tr></tbody></table>
      <p><em>Standard units can be as little as 12-13 weeks, depending on stock</em>.</p>
      <p><strong>NPD Announcements</strong>: &nbsp; R-454B units to be available for shipment end of Q4 2023, released in JESS Selection Software by Q3 2023.</p>
      <p>Give me a shout if you have any questions.</p>
      <p>Best,</p>
      """

   TomSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;line-height:105%;font-family:"Arial",sans-serif;color:#0055A0;'>Tom Armstrong</span><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><br> <strong><span style="color:#F8B124;">Business Development Director</span></strong></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>Napps Technology&nbsp;</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#F8B124;'>|</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;JETSON INNOVATIONS</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;</span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.nappstech.com/">www.nappstech.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/">www.jetsonhvac.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>D:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:951-389-4741"><span style="color:#00549F;">951.389.4741</span></a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>C:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:951-544-6283"><span style="color:#00549F;">951.544.6283</span></a></span></span></p>
      """
   PauloSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style='font-size:13px;line-height:105%;font-family:"Arial",sans-serif;color:#0055A0;'>Paulo Herrera</span><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><br> <strong><span style="color:#F8B124;">Inside Sales Engineer</span></strong></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>Napps Technology&nbsp;</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#F8B124;'>|</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;JETSON INNOVATIONS</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;</span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.nappstech.com/">www.nappstech.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/">www.jetsonhvac.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>Direct:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:903-758-2900"><span style="color:#00549F;">903-758-2900 Ext. 146</span></a></span></span></p>
      """
   JetsonSignature = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>Napps Technology&nbsp;</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#FFC000;'>|</span></strong><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;JETSON INNOVATIONS</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#1F3864;'>&nbsp;</span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.nappstech.com/">www.nappstech.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;'><a href="http://www.jetsonhvac.com/">www.jetsonhvac.com</a></span></span></p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;line-height:105%;font-size:15px;font-family:"Calibri",sans-serif;'><strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>Direct:</span></strong><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#5F6060;'>&nbsp;&nbsp;</span><span style="color:blue;text-decoration:underline;"><span style='font-size:12px;line-height:105%;font-family:"Arial",sans-serif;color:#00549F;'><a href="tel:903-758-2900"><span style="color:#00549F;">903-758-2900 Ext. 146</span></a></span></span></p>
      """

   if region == "W":
      signature = TomSignature
   elif region == "E":
      signature = PauloSignature
   else:
      signature = JetsonSignature

   return body + signature

if __name__ == "__main__":
   main()
   input("Program Completed. Press ENTER to Close...")