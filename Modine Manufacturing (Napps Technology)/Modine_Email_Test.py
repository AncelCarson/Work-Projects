#pylint: disable = all,invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 4/29/2025
# Start.py

"""A one line summary of the module or program, terminated by a period.

Leave one blank line.  The rest of this docstring should contain an
overall description of the module or program.  Optionally, it may also
contain a brief description of exported classes and functions and/or usage
examples.

Functions:
   main: Driver of the program
"""
#Libraries
import smtplib
import ssl

import time
import getpass
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Variables
def main():
   # email = input("Please enter the User Email:\n")
   # password = getpass.getpass("Please enter the User Password:\n")
   sendEmail("Ancel_Test@modine.com")
   input("Program complete. Press ENTER to Close...")

#Functions
" Main Finction "
def sendEmail(email):
   if email == "brian@pocams.com":
      SERVER = "smtpout.secureserver.net"
      PORT = 465
      context = ssl.create_default_context()
      try:
         server = smtplib.SMTP_SSL(host = SERVER, port = PORT, context=context)
      except Exception as err:
         print(err)
         return
      server.ehlo()
      # server.starttls()
   else:
      SERVER = "mail-na.modine.com"
      PORT = 25
      server = smtplib.SMTP("mail-na.modine.com", 25)
      context = ssl.create_default_context()    
      server.starttls(context=context)

   # server.login(email,password)

   # try:
   #    server.login(email)
   # except Exception as err:
   #    print(err)
   #    return

   TO = "ancel.h.carson@modine.com"
   FROM = email
   SUBJECT = "Test Email"

   msg = MIMEMultipart('alternative')
   msg['Subject'] = SUBJECT
   msg['From'] = FROM
   msg['To'] = TO

   text = """<p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>This is a test message. Let me know if you go it.</p>"""

   msg.attach(MIMEText(text, 'html'))

   try:
      server.sendmail(FROM, TO.split(","), msg.as_string())
   except Exception as err:
      print(err)
      return

   server.quit()

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()