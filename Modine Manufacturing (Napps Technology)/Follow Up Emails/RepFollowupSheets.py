# Author: Ancel Carson
# Orginization: Napps Technology Comporation
# Creation Date: 14/10/2020
# Update Date: 12/18/23
# RepFollowupSheets.py

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
import pandas as pd
from datetime import datetime


#custom Modules
sys.path.insert(0,r'S:\Programs\Add_ins')
import MakeEmailList as MEL
from Loader import Loader

#Variables
start_folder = r'S:\NTC Books of Knowledge\Sales (Part and Unit Quotes, Customer Interactions, Pricing, Reports, Binders)\Job Followups\Followup '
day = datetime.now().strftime('%y%m%d')
folder = start_folder + day
os.system('mkdir "{}"'.format(folder))


#Functions
" Main Finction "
def main():
   day = datetime.now().strftime('%y%m%d')
   emailList = MEL.main()
   illegal_characters = "[]:*?/\\"
   emailList["SalesOffice"] = emailList.apply(lambda x: "".join(char for char in x["SalesOffice"] if char not in illegal_characters), axis = 1)
   companies = emailList["Company"].unique()
   senders = emailList["Sender"].unique()

   loader = Loader("Generating Summary Sheets...", "Sheets Generated", 0.1).start()
   for company in companies:
      for sender in senders:
         # groups.append(emailList.loc[(emailList["Company"] == company) & (emailList["Sender"] == sender)])
         group = emailList.loc[(emailList["Company"] == company) & (emailList["Sender"] == sender)]
         if group.empty:
            continue
         offices = group["SalesOffice"].unique()
         if sender == "None":
            sender = "General"
         file = folder + "\\" + sender + " " + company + " Follow Up List.xlsx"
         writer = pd.ExcelWriter(file, engine='xlsxwriter')
         performance = pd.DataFrame(offices, columns = ["Representatives"])
         performance["Quoted"] = "=SUMIFS('Master'!D:D,'Master'!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-1)))"
         performance["Won"] = "=SUMIFS('Master'!D:D,'Master'!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-2)),'Master'!F:F,D1)"
         performance["% Conversion"] = "=(INDIRECT(ADDRESS(ROW(),COLUMN()-1))/INDIRECT(ADDRESS(ROW(),COLUMN()-2)))*100"
         performance["Lost"] = "=SUMIFS(Master!D:D,Master!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-4)),Master!F:F,F1)"
         performance["Abandoned"] = "=SUMIFS(Master!D:D,Master!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-5)),Master!F:F,G1)"
         performance["Hold"] = "=SUMIFS(Master!D:D,Master!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-6)),Master!F:F,H1)"
         performance["Design"] = "=SUMIFS(Master!D:D,Master!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-7)),Master!F:F,I1)"
         performance["Hot"] = "=SUMIFS(Master!D:D,Master!B:B,INDIRECT(ADDRESS(ROW(),COLUMN()-8)),Master!F:F,J1)"
         performance.to_excel(writer, sheet_name="Performance")
         master = group[["SalesOffice","SalesOrderNo","Total Quote","CUSTOMER_REFERENCE"]].copy()
         master["Status"] = ""
         master["Reply"] = ""
         master["Follow Up"] = ""
         master["Comments"] = ""
         master.to_excel(writer, sheet_name="Master")
         for office in offices:
            page = group.loc[group["SalesOffice"] == office, ["SalesOffice","SalesOrderNo","Dollars","CUSTOMER_REFERENCE"]].copy()
            page["Status"] = "=VLOOKUP(INDIRECT(ADDRESS(ROW(),COLUMN()-3)),'Master'!C:I,4,FALSE)"
            page["Reply"] = "=VLOOKUP(INDIRECT(ADDRESS(ROW(),COLUMN()-4)),'Master'!C:I,5,FALSE)"
            page["Follow Up"] = "=VLOOKUP(INDIRECT(ADDRESS(ROW(),COLUMN()-5)),'Master'!C:I,6,FALSE)"
            page["Comments"] = "=VLOOKUP(INDIRECT(ADDRESS(ROW(),COLUMN()-6)),'Master'!C:I,7,FALSE)"
            page.to_excel(writer, sheet_name=office)
         worksheet = writer.sheets['Master']
         worksheet.data_validation('F2:F500', {'validate': 'list',
                                  'source': ['Won','Lost',' Abandoned','Hold','Design','Hot','Open']})
         writer.close()
   loader.stop()

" Checks if this program is beiong called "
if __name__ == "__main__":
   main()
   input("Program is completed, Press ENTER to close...")


