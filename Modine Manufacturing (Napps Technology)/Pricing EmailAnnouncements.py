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

inputWorkbook = r'S:\Ancel\Pricing\Pricing Announcements\Job Followup 221117.xlsx'
filename = "Jetson Lead Time and Price Adjust 230308.pdf"
JetsonFile = r'S:\Ancel\Pricing\Pricing Announcements\Jetson Lead Time and Price Adjust 230308.pdf'
JetsonAttachment = open(JetsonFile, 'rb')
# TraneFile = r'S:\Ancel\Sales\Job Followups\CGWR, CCAR and CICD Low Lead Times 220923.pdf'
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
   server.login(AEmail, APassW)

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
         SUBJECT = 'Jetson Price Increase Announcement'
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
      #    server.login(AEmail, APassW)
      #    print("Connection Restored")
      #    server.sendmail(FROM, TO, msg.as_string())

      print("{0} email sent to {1}".format(note, row[0]))
      time.sleep(2)

   server.quit()

def getMessage(name, email):
   body = """
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Hey {0},</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Attached is a letter from the president regarding upcoming Price Increases.&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;Jetson chiller and condensing unit lead times remain at an industry leading 18-20 weeks for standard chillers and condensing units. Heat pump lead time is currently 30 weeks. Please take advantage of this industry low lead time to close orders and don&rsquo;t hesitate to reach us at <a href="mailto:sales@JetsonHVAC.com"><span style="color:windowtext;text-decoration:none;">sales@JetsonHVAC.com</span></a>, if we can be of assistance.</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;The following average product increases will be in effect for quotes after March 31, 2023.</p>
      <table style="border-collapse:collapse;border:none;">
         <tbody>
            <tr>
                  <td style="width: 157.1pt;border: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Air-Cooled Chillers (model ACC)</p>
                  </td>
                  <td style="width: 121.5pt;border-top: 1pt solid windowtext;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-image: initial;border-left: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>Average Price Increase</p>
                  </td>
            </tr>
            <tr>
                  <td style="width: 157.1pt;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-left: 1pt solid windowtext;border-image: initial;border-top: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:right;'>10 &ndash; 20 tons</p>
                  </td>
                  <td style="width: 121.5pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>13%</p>
                  </td>
            </tr>
            <tr>
                  <td style="width: 157.1pt;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-left: 1pt solid windowtext;border-image: initial;border-top: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:right;'>25 &ndash; 50 tons</p>
                  </td>
                  <td style="width: 121.5pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>10%</p>
                  </td>
            </tr>
            <tr>
                  <td style="width: 157.1pt;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-left: 1pt solid windowtext;border-image: initial;border-top: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:right;'>60 &ndash; 80 tons</p>
                  </td>
                  <td style="width: 121.5pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>2%</p>
                  </td>
            </tr>
            <tr>
                  <td style="width: 157.1pt;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-left: 1pt solid windowtext;border-image: initial;border-top: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>Water-Cooled Chillers (model FWC)</p>
                  </td>
                  <td style="width: 121.5pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>&nbsp;</p>
                  </td>
            </tr>
            <tr>
                  <td style="width: 157.1pt;border-right: 1pt solid windowtext;border-bottom: 1pt solid windowtext;border-left: 1pt solid windowtext;border-image: initial;border-top: none;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:right;'>20 &ndash; 80 tons</p>
                  </td>
                  <td style="width: 121.5pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 5.4pt;vertical-align: top;">
                     <p style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri",sans-serif;text-align:center;'>2%</p>
                  </td>
            </tr>
         </tbody>
      </table>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;Special components used in heat pumps, variable speed compressors and variable speed fans continue to face conditions that impact our ability to source, manufacture, and deliver products to you in a cost effective and timely manner. We have recently seen moderation in inflation and many vendor lead times have flattened &ndash; not at historic norms, but many vendors are holding lead time or showing some reduction in time.</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp;</p>
      <p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Thank you for your partnership and cooperation as we find ways to best serve you.</p>
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

if __name__ == "__main__":
   main()
   input("Program Completed. Press ENTER to Close...")