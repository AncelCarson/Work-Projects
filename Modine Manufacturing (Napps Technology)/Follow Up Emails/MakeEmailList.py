# pylint: disable=invalid-name,bad-indentation,all
# -*- coding: utf-8 -*-

# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 9/5/2023
# Update Date: 4/2/2025
# MakeEmailList.py

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
import glob
import time
import locale
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

#Secret Variables
load_dotenv()
Shared_Drive = os.getenv('Shared_Drive')

#custom Modules
#pylint: disable=wrong-import-position
sys.path.insert(0,fr'\\{Shared_Drive}\Programs\Add_ins')
from Loader import Loader
#pylint: enable=wrong-import-position

#Variables
file_path = fr'\\{Shared_Drive}\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Follow Up Data\*.xlsx' # * means all if need specific format then *.xlsx
files = sorted(glob.iglob(file_path), key=os.path.getctime, reverse=True)
latest_file = files[0]
checkFile = files[4]
state_File = fr'\\{Shared_Drive}\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Regional Data\Company State.xlsx'
rep_File = fr'\\{Shared_Drive}\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Regional Data\Rep State.xlsx'
locale.setlocale( locale.LC_ALL, '' )

#Functions
" Main Finction "
def main():
   followUpDays = int((time.time() - os.path.getctime(checkFile)) / (60 * 60 * 24))
   loader = Loader("Loading Tables...", "Tables Loaded", 0.1).start()
   dfIn = pd.read_excel(latest_file, sheet_name = 'Sheet1', header = 0)
   repDf = pd.read_excel(rep_File, sheet_name = 'Sheet1', header = 0)
   stateDf = pd.read_excel(state_File, sheet_name = 'Sheet1', header = 0)
   data = dfIn.copy()
   loader.stop()
   # print(dfIn)

   loader = Loader("Formatting Data...", "Data Formatted", 0.1).start()
   dfIn["QuoteNum"] = dfIn.apply(lambda x: x["SalesOrderNo"][3:] if x["SalesOrderNo"][0] == "0" else x["SalesOrderNo"][:4], axis = 1)
   dfIn["DaysLate"] = dfIn.apply(lambda x: int((datetime.now() - x["OrderDate"]).days), axis = 1)
   dfIn["MemoDays"] = dfIn.apply(lambda x: 1000 if pd.isna(x["MemoDate"]) else int((datetime.now() - x["MemoDate"]).days), axis = 1)
   dfIn["ReminderDays"] = dfIn.apply(lambda x: 1000 if pd.isna(x["ReminderStartDate"]) else int((datetime.now() - x["ReminderStartDate"]).days), axis = 1)
   dfIn["First Name"] = dfIn.apply(lambda x: str(x["Salesperson"]).split(" ")[0], axis = 1)
   dfIn["Dollars"] = dfIn.apply(lambda x: locale.currency(x["Total Quote"], grouping=True ), axis = 1)
   dfIn.fillna({'CUSTOMER_REFERENCE':"Chiller Quote"}, inplace=True)
   dfIn.fillna({'EmailAddress':"No Email"}, inplace=True)
   dfIn.fillna({'First Name':""}, inplace=True)

   quotesList = dfIn[["SalesOrderNo","QuoteNum","OrderDate","DaysLate","Status","Cancel\nReasonCode","First Name","EmailAddress","SalespersonName","SalesOffice","State","CUSTOMER_REFERENCE","COMMENTS","Dollars","Total Quote","MemoDate","ReminderStartDate","MemoDays","ReminderDays"]].copy()
   quotesList["CurrentStatus"] = ""
   quotesList["QuotetoClose"] = ""
   quotesList["RevivedQuote"] = ""
   quotesList["Sender"] = "None"
   quotesList["Company"] = quotesList.apply(getCompany, axis = 1)
   # quotesList["MemoDays"] = quotesList.apply(getMemoDays, axis = 1)

   quotesList.sort_values(by=["QuoteNum","OrderDate","SalesOrderNo"], ascending=[True, False, False], inplace=True)
   quotesList.reset_index(drop=True, inplace=True)
   loader.stop()
   # print(quotesList)

   loader = Loader("Setting Quote Status...", "Status Set", 0.1).start()
   quoteStuff = [0000,"Complete","Current"]
   for index, row in quotesList.iterrows():
      # Set Current Status
      quote = row["QuoteNum"]
      if quote != quoteStuff[0]:
         status = "Current"
      else:
         status = "Old"
      quotesList.at[index,"CurrentStatus"] = status
      
      if status == "Old":
          # Find Quotes that Should be Closed
         if row["Status"] == "Quote" and quoteStuff[1] != "Quote":
            quotesList.at[index,"QuotetoClose"] = "Flag"

         # Find Quotes that have been Revived
         if quoteStuff[1] == "Quote" and row["Status"] != "Quote" and quoteStuff[2] == "Current":
            quotesList.at[index,"RevivedQuote"] = "Flag"

      # Set Values for Next Loop
      quoteStuff = [quote, row["Status"], status]

      if pd.isna(row["State"]):
         office = stateDf.loc[stateDf["Rep"] == row["SalesOffice"]]
         if office.empty:
            continue
         office.reset_index(drop=True, inplace=True)
         quotesList.at[index,"State"] = office["State"].iat[0]
         row["State"] = office["State"].iat[0]

      for col in repDf:
         if repDf[col].isin([row["State"]]).any():
            quotesList.at[index,"Sender"] = col
            break


   # cleaning any bad Sender Emails
   quotesList["Company"] = quotesList.apply(lambda x: "None" if x["Company"] == "" else x["Company"], axis = 1)

   loader.stop()
   # print(quotesList)

   if input("Do you want to review Quotes to be Closed?: y/n\n") == "y":
      closeQuotes(quotesList)

   # if input("Do you want to review Quotes that have been Revived?: y/n\n") == "y":
   #    reviveQuotes(quotesList)

   # Filters out Old Quotes
   emailList = quotesList.loc[quotesList["CurrentStatus"] == "Current"]
   emailList = emailList.loc[emailList["Status"] == "Quote"]
   
   # Filters out Any Quote with no email and saves them to a list
   noEmailList = emailList.loc[emailList["EmailAddress"] == "No Email"]
   emailList = emailList.loc[emailList["EmailAddress"] != "No Email"]

   # Filters out any quote that has a reminder date set in the future
   emailList = emailList.loc[((emailList["ReminderStartDate"] < datetime.now()) | (pd.isna(emailList["ReminderStartDate"]))) & (emailList["DaysLate"] > 60)]

   # Filters out any quote that is either 150 Days old or hasnt had a response in over 90 Days adn saves them to a list
   closeList = emailList.loc[(emailList["DaysLate"] > (followUpDays + 60)) & (emailList["MemoDays"] > followUpDays) & (emailList["ReminderDays"] > followUpDays)]
   emailList = emailList.loc[(emailList["DaysLate"] < (followUpDays + 60)) | (emailList["MemoDays"] < followUpDays) | (emailList["ReminderDays"] < followUpDays)]
   # print(closeList)

   emailList["Units"] = emailList.apply(getUnits, axis = 1)

   #Writes the data to the original excel file
   emailList.reset_index(drop=True, inplace=True)
   closeList.reset_index(drop=True, inplace=True)
   loader = Loader("Generating Email File...", "File Made", 0.1).start()
   writer = pd.ExcelWriter(latest_file)
   data.to_excel(writer, sheet_name="Sheet1")
   quotesList.to_excel(writer, sheet_name="120 Day Followup")
   closeList.to_excel(writer, sheet_name="Quotes to Close")
   emailList.to_excel(writer, sheet_name="Email List")
   writer.close()
   loader.stop()

   if noEmailList.empty:
      print("\n\nAll Quotes have an associated email\n")
   else:
      print("\n\nThe Following Quotes have no Email address so no email will be sent.")
      print(noEmailList[["SalesOrderNo","OrderDate","First Name","CUSTOMER_REFERENCE","COMMENTS"]])
      print("\nAdd an Address to the Quote to add them to the email list\n")

   return emailList

def closeQuotes(quotesList):
   closeList = quotesList.loc[quotesList["QuotetoClose"] == "Flag"]
   if closeList.empty:
      print("\n\nThere are no Quotes to close\n")
      return
   for index, row in closeList.iterrows():
      quote = row["QuoteNum"]
      quotes = quotesList.loc[quotesList["QuoteNum"] == quote]
      print(quotes[["SalesOrderNo","OrderDate","DaysLate","Status","Cancel\nReasonCode","First Name","CUSTOMER_REFERENCE","Dollars","CurrentStatus"]])
      if input("Should Quote #{} be altered? y/n\n".format(quote)) == "y":
         if input("Should Quote #{} be skipped? y/n\n".format(quote)) == "y":
            referenceQuote = ""
         else:
            referenceQuote = input("Which Quote should be referenced?\n")
         for i, r in quotesList.loc[quotesList["QuoteNum"] == quote].iterrows():
            if r["SalesOrderNo"] == referenceQuote:
               status = "Current"
            else:
               status = "Old"
            quotesList.at[i,"CurrentStatus"] = status
      # print(quotesList.loc[quotesList["QuoteNum"] == quote])

def reviveQuotes(quotesList):
   reviveList = quotesList.loc[quotesList["RevivedQuote"] == "Flag"]
   print(reviveList)

def getCompany(row):
   if row["SalespersonName"] == "House (Do not use)" or row["SalespersonName"] == "Independent- Non Trane":
      return "Jetson"
   else:
      return "Napps"

def getUnits(row):
   if pd.isna(row["COMMENTS"]):
      return " for some of our units"
   units = row["COMMENTS"].split(";")
   if len(units) == 0:
      return " for some of our units"
   count = 1
   unitsString = ""
   for unit in units:
      unit = unit.lstrip()
      unit = unit.split(" ")
      if count == 1:
         unitsString += " for "
      elif count == len(units):
         unitsString += " and "
      if len(unit) != 3:
         return " for some of our units"
      unit.append(unit[0][5:])
      if len(unit[3]) <= 1:
         return " for some of our units"
      unit[0] = unit[0][:4]
      if unit[0] == "SPLT":
         unit[0] = "Split System"
      unitString = "({}) {} Ton {} {}".format(unit[2], unit[3], unit[0], unit[1])
      unitsString += unitString
      if count != len(units) and len(units) > 2:
         unitsString += ","
      count += 1
   print(unitsString)
   return unitsString

" Checks if this program is being called "
if __name__ == "__main__":
   main()
   input("Program is completed, Press ENTER to close...")