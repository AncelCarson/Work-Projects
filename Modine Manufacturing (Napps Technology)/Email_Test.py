#pylint: disable = all,invalid-name,bad-indentation,non-ascii-name
#-*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 2/3/2021
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
   email = input("Please enter the User Email:\n")
   password = getpass.getpass("Please enter the User Password:\n")
   sendEmail(email,password)
   input("Program complete. Press ENTER to Close...")

#Functions
" Main Finction "
def sendEmail(email,password):
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
      SERVER = "smtp.office365.com"
      PORT = 587
      server = smtplib.SMTP("smtp.office365.com", 587)
      context = ssl.create_default_context()    
      server.starttls(context=context)

   # server.login(email,password)

   try:
      server.login(email,password)
   except Exception as err:
      print(err)
      return

   TO = "ancel.h.carson@modine.com"
   FROM = email
   SUBJECT = "Test Email"

   msg = MIMEMultipart('alternative')
   msg['Subject'] = SUBJECT
   msg['From'] = FROM
   msg['To'] = TO

   text = """<p style='margin-top:0in;margin-right:0in;margin-bottom:.0001pt;margin-left:0in;font-size:16px;font-family:"Calibri",sans-serif;'>This is a test message</p>"""

   msg.attach(MIMEText(text, 'html'))

   try:
      server.sendmail(FROM, TO, msg.as_string())
   except Exception as err:
      print(err)
      return

   server.quit()

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()