### Ancel Carson
### 8/9/2019
### Windows 10
### Python command line, Notepad, IDLE
###
### This program will update all fields in the
### DI Databases that reference data from
### Drilling Info

# Required Libraries: os, pandas, datetime
# Imported Modules: LoginCredentials

import os
import pandas as pd
from datetime import datetime
from datetime import timedelta

READMELocation = r'Documents\ToFlatFile READ ME.txt'
fileLocation = r'MonthlyFlatFile.csv'

## Main method
def main():
    print('This program will make a flat file of the Production Data')
    selection = menu1()

    if selection == '1':
        startDate = datetime.today()
        endDate = datetime.today()
    elif selection == '2':
        print('Enter all dates in mm/dd/yyyy format')
        selectedStart = input('Enter selected day\n')
        startDate = datetime.strptime(selectedStart, '%m/%d/%Y')
        endDate = startDate
    elif selection == '3':
        monthSelection = menu2()
        if monthSelection == '1':
            numMonths = input('How many months back should be included?\n')
            endDate = datetime.today() 
            startDate = endDate - timedelta(days = (int(numMonths)*30))
        elif monthSelection == '2':
            print('Enter all dates in mm/dd/yyyy format')
            selectedStart = input('Enter start date\n')
            selectedEnd = input('Enter end date\n')
            startDate = datetime.strptime(selectedStart, '%m/%d/%Y')
            endDate = datetime.strptime(selectedEnd, '%m/%d/%Y')
        else:
            print('The program will now end')
            return 1
    else:
            print('The program will now end')
            return 1
        

    exportData(startDate, endDate)

    if __name__ == '__main__':
        LC.CloseConnection()
    input('\nThe program has finished.\nPress enter to continue...')
            

## Selection of single month or multiple
def menu1():

    print("Select the appropriate options from the menus.")
    print("|---------------------------|")
    print("| 1) Output current month   |")
    print("| 2) Output selected month  |")
    print("| 3) Output range of months |")
    print("| 4) Open Program READ ME   |")
    print("| 5) Exit Program           |")
    print("|---------------------------|")
    selected = input("\nType selected number and press Enter...\n")
    if selected == '4':
        os.startfile(READMELocation)
        return menu1()
    else:
        return selected

##Selecting range of months
def menu2():
    
    print("Select the appropriate options from the menus.")
    print("|---------------------------------------------|")
    print("| 1) Select today and a number of months back |")
    print("| 2) Select range of dates                    |")
    print("| 3) Exit Program                             |")
    print("|---------------------------------------------|")
    return input("\nType selected number and press Enter...\n")


## Runs selecting data and adding it to .csv file
def exportData(startDate, endDate):
    cnxn = LC.Accesscnxn
    cursor = LC.Accesscursor
    monthlyProduction = []

    if startDate == endDate:
        endDate = startDate + timedelta(days = 30)
        startDate = '{:%Y-%m}'.format(startDate)
        endDate = '{:%Y-%m}'.format(endDate)
        for row in cursor.execute("SELECT [ARIES Master].DBSKEY, [ARIES Master].PROPNUM, \
            Statuses.UnitNumber, [Production Monthly].PROD_DATE, [Production Monthly].PROD_MONTH_NO, \
            [Production Monthly].GAS, [Production Monthly].OIL, [Production Monthly].WATER,\
            [Production Monthly].UPDATE_DATE FROM (([DI Production Download] LEFT JOIN Statuses\
            ON [DI Production Download].[Unit Number] = Statuses.UnitNumber) LEFT JOIN \
            [ARIES Master] ON Statuses.UnitNumber = [ARIES Master].VEF_UNIT_NO) LEFT JOIN \
            [Production Monthly] ON Statuses.API14 = [Production Monthly].API_NO\
            WHERE PROD_DATE BETWEEN ? AND ?", startDate, endDate):
            monthlyProduction.append(row)
    else:
        if startDate > endDate:
            tempStart = startDate
            tempEnd = endDate
            endDate = tempStart
            startDate = tempEnd
        for row in cursor.execute("SELECT [ARIES Master].DBSKEY, [ARIES Master].PROPNUM, \
            Statuses.UnitNumber, [Production Monthly].PROD_DATE, [Production Monthly].PROD_MONTH_NO, \
            [Production Monthly].GAS, [Production Monthly].OIL, [Production Monthly].WATER,\
            [Production Monthly].UPDATE_DATE FROM (([DI Production Download] LEFT JOIN Statuses\
            ON [DI Production Download].[Unit Number] = Statuses.UnitNumber) LEFT JOIN \
            [ARIES Master] ON Statuses.UnitNumber = [ARIES Master].VEF_UNIT_NO) LEFT JOIN \
            [Production Monthly] ON Statuses.API14 = [Production Monthly].API_NO \
            WHERE PROD_DATE BETWEEN ? AND ?", startDate, endDate):
            monthlyProduction.append(row)
        

    if monthlyProduction == []:
        print('There was no data in the selected timeframe')
    else:
        columnNames = ['DBSKEY','PROPNUM','UNIT_NO','PROD_DATE','PROD_MONTH','GAS','OIL','WATER','UPDAT_DATE']
        dfP = pd.DataFrame.from_records(monthlyProduction, columns = columnNames)
        try:
            export_dfP = dfP.to_csv(fileLocation, index = False)
        except:
            print('Make sure MonthlyFlatFile.csv is closed')
            input('Press enter when the file is closed...\n')
            try:
                export_dfP = dfP.to_csv(fileLocation, index = False)
            except:
                print('File was not closed. Exiting program')
                return 1
        print('Data has been exported')
	
## Main control
if __name__ == '__main__':
    import LoginCredentials as LC
    main()
else:
    from Modules import LoginCredentials as LC
    READMELocation = r'Modules\Documents\ToFlatFile READ ME.txt'
    fileLocation = r'Modules\MonthlyFlatFile.csv'
