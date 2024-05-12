### Ancel Carson
### 8/7/2019
### Windows 10
### Python command line, Notepad, IDLE
###
### This program will walk the user through
### adding new wells to both databases.

# Required Libraries: os, pandas
# Imported Modules: LoginCredentials
print("Loading Libraries.")
import os
import pandas as pd


excelLocation = r'Monthly Status Update.xlsx'
READMELocation = r'Documents\StatusUpdate READ ME.txt'
infoLocation = r'Documents\Access Database Documentation.txt'


## Main method of program
def main():
    print('\n\nThis program is for updating the Statuses table.')
    print('It will walk through the steps and run functions.')
    print('Both the Access and SQL Databases will be updated.\n')

    menuSelect = menu()

    if (menuSelect == '1' or menuSelect == '2' or menuSelect == '100'):
        updateProcess(menuSelect)

    # Ensures connection is not closed early
    if __name__ == '__main__':
        LC.CloseConnection()
    input('\nThe program has finished.\nPress enter to continue...')


## Menu for selection desired function
def menu():
    print("Select the appropriate options from the menus.")
    print("|-----------------------------|")
    print("| 1) Run with instructions    |")
    print("| 2) Run without instructions |")
    print("| 3) Run update only          |")
    print("| 4) Open Database Document   |")
    print("| 5) Open program READ ME     |")
    print("| 6) Exit Program             |")
    print("|-----------------------------|")
    selected = input("\nType selected number and press Enter...\n")
    if selected == '3':
        selected = caseThree()
    if selected == '4':
        os.startfile(infoLocation)
        selected = menu()
    if selected == '5':
        os.startfile(READMELocation)
        selected = menu()

    print()
    return selected


## Used as a check to ensure databases are ready
def caseThree():
    print('This option requires that both databases have been prepared.')
    print('Is this the case? \n\t1) Yes\n\t2) No\n\t3) What is "Prepared"?')
    selected = input("\nType selected number and press Enter...\n")
    if selected == '1':
        return '100'
    elif selected == '2':
        return menu()
    elif selected == '3':
        definePrepared()
        return caseThree()
    else:
        print('\nPlease select 1 or 2')
        return caseThree()


## Definition to clarify to user
def definePrepared():
    print('\nTo be "prepared", both databases need to be properly set up.')
    print('In the Access database, both "Statuses" and "Status Months" must be up to date.')
    print('\tStatuses must have all well statuses for each month.')
    print('\tStatus Months must have one entry for each month column in Statuses.')
    print('In the SQL database, there must be the same number of columns as Access.')
    print('All date columns must share the same name.\n')


## Runs three variations of the program based off menu1
def updateProcess(choice):
    if choice == '1':
        print('This program is divided into steps.')
        print('To advance the sequence, press Enter when prompted.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('To start, the Status Update excel file will be opened.')
        os.startfile(excelLocation)
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('The monthly "G/L" file from accounting will also be needed.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('Once both are open copy and paste the following values:')
        print('\tUnit Number\n\tVEFI Name\n\tStatuses')
        print('Copy values from monthly accounting "G/L" file.')
        print('Paste values to the "Statuses from GL file" sheet in excel.')
        print('\nAll wells without statuses are assumed "PDP".')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('Once added, open the "Statuses to Update" sheet.')
        print('Click inside column A and select "Design" from the toolbar.')
        print('In design, click the refresh icon.')
        print('For readability sake, hide and new columns that appeared.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('Scroll to the bottom of the sheet.')
        print('Make sure columns 3 and 4 are the same length as column 1.')
        print('If they are not, select the last two cells and drag the formula down.')
        print('The Access Database will now be opened.')
        os.startfile(databaseLocation)
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('In the Database, open the "Statuses" table.')
        print('Select the last column and add a new "Short Text" field.')
        print('Title the column the month followed be the year in two digits.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        newMonth = input('Type the exact name of the column added here and press Enter.\n')
        print('\n-------------------------------------------------------------------------\n')
        newMonth = nameCheck(newMonth)
        print('From the excel sheet, copy all the values in the 4th column.')
        print('Be sure to start at the second row so there is no blank cell.')
        print('Once copied, add those values to the new column of the Access database.')
        print('Select the column header and paste.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('In the excel file and select the "Statuses not in Access" sheet.')
        print('Scroll to the last column and select the last cell in row 1.')
        print('Drag the formula a few cells over.')
        print('There should be the same number of colums in Access and Excel.')
        print('Filter out all blanks in the "Unit Number" column.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('Fill in the STATE, API10, API14, and OP_CODE.')
        print('Copy the values and return to the table in Access.')
        print('Select the "*" at the bottom of the table and paste the values.')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        print('Save and close both Excel and Access')
        input('Press Enter to continue...\n')
        print('-------------------------------------------------------------------------\n')
        addNewData(newMonth)
    if choice == '2':
        print('Opening Excel and Access')
        os.startfile(excelLocation)
        os.startfile(databaseLocation)
        input('Press Enter to when both have been updated...\n')
        print('-------------------------------------------------------------------------\n')
        newMonth = input('Type the exact name of the column added here and press Enter.\n')
        print('\n-------------------------------------------------------------------------\n')
        newMonth = nameCheck(newMonth)
        addNewData(newMonth)
    if choice == '100':
        syncDatabases()


def nameCheck(newMonth):
    Accesscursor = LC.Accesscursor
    
    try:
        Accesscursor.execute("SELECT [{col_name}] FROM [Statuses]".format(col_name = newMonth))
    except:
        print('Name entered is not the same as in Access')
        newMonth = input('Type the exact name of the column added here and press Enter.\n')
        print('\n-------------------------------------------------------------------------\n')
        nameCheck(newMonth)
    else:
        return newMonth
    

## Adds new column to SQL and new entry to Status Months in Access
def addNewData(newMonth):
    Accesscnxn = LC.Accesscnxn
    Accesscursor = LC.Accesscursor
    SQLcnxn = LC.SQLcnxn
    SQLcursor = LC.SQLcursor
    count = 0

    # Collecting months for which there is data
    for month in Accesscursor.execute("SELECT Months FROM [Status Months]"):
        count += 1

    count += 1

    Accesscursor.execute("INSERT INTO [Status Months](ID, Months) VALUES (?, ?)", str(count),\
                         str(newMonth))
    SQLcursor.execute("ALTER TABLE Statuses ADD [{col_name}] nvarchar(MAX)".format(col_name = newMonth))
    Accesscnxn.commit()
    SQLcnxn.commit()

    syncDatabases()

    
## Inputs new wells and most recent status fot all wells to SQL
def syncDatabases():
    firstMonth = True
    inSQL = False
    Accesscnxn = LC.Accesscnxn
    Accesscursor = LC.Accesscursor
    SQLcnxn = LC.SQLcnxn
    SQLcursor = LC.SQLcursor
    AccessStatuses = []
    AccessMonths = []
    SQLWells = []
    AccessWells = []
    monthString = []

    print('Collectiong most recent status data from Access.')

    # Collecting months for which there is data
    for month in Accesscursor.execute("SELECT Months FROM [Status Months]"):
        AccessMonths.append(month[0])

    # Filtering to most recent month
    currentMonth = AccessMonths[len(AccessMonths)-1]

    # Retrirving most recent status for all wells
    for status in Accesscursor.execute("SELECT UnitNumber,[{col_name}] FROM [Statuses]"\
        .format(col_name = currentMonth)):
        AccessStatuses.append(status)

    # Formatting data
    dfNS = pd.DataFrame.from_records(AccessStatuses, columns = ['UnitNumber', 'Status'])
    dfNS = dfNS.sort_values(by=['UnitNumber'])

    print('Adding Most recent status data to SQL Server.')

    # Adding new statuses to SQL
    for index,row in dfNS.iterrows():
        SQLcursor.execute("UPDATE Statuses SET [{col_name}] = ? WHERE UNIT_NO = ?"\
            .format(col_name = currentMonth), str(row['Status']), str(row['UnitNumber']))

    print('Adding any new wells to the SQL Server.')

    # Retrieving well number of all wells in SQL
    for well in SQLcursor.execute("SELECT UNIT_NO FROM Statuses"):
        SQLWells.append(well[0])

    SQLWells.sort()

    # Retrieving data of all wells in Access
    for well in Accesscursor.execute("SELECT * FROM Statuses"):
        AccessWells.append(well)


    # Formatting data
    columns = ['UnitNumber', 'BTU', 'VEFIName', 'State', 'API10', 'API14', 'LOE', 'OP_CODE',\
               'PRIOR_OIL', 'PRIOR_GAS', 'LOCATION']
    allColumns = columns + AccessMonths
    
    dfNW = pd.DataFrame.from_records(AccessWells)
    dfNW.columns = allColumns
    dfNW = dfNW.sort_values(by=['UnitNumber'])

    # Adding new wells from Access to SQL
    for index,row in dfNW.iterrows():
        inSQL = False
        count = 0
        for well in SQLWells:
            if row['UnitNumber'] == well:
                inSQL = True
                del SQLWells[count]
                break
            count += 1
        if inSQL == False:
            # Inserting data for non variable fields
            SQLcursor.execute("INSERT INTO Statuses(UNIT_NO, VEFI_NAME, API10, API14, OP_CODE)\
                VALUES(?, ?, ?, ?, ?)", str(row['UnitNumber']), str(row['VEFIName']),\
                str(row['API10']), str(row['API14']), str(row['OP_CODE']))
            # Inserting data for variable months
            for month in AccessMonths:
                SQLcursor.execute("UPDATE Statuses SET [{col_name}] = ? WHERE UNIT_NO = ?"\
                    .format(col_name = month), str(row[month]), str(row['UnitNumber']))

    # Commit changes to databases
    LC.SQLcnxn.commit()
    LC.Accesscnxn.commit()


## Main control
if __name__ == '__main__':
    import LoginCredentials as LC
    databaseLocation = LC.databaseLocation
    main()
else:
    from Modules import LoginCredentials as LC
    databaseLocation = LC.databaseLocation
    excelLocation = r'Modules\Monthly Status Update.xlsx'
    READMELocation = r'Modules\Documents\StatusUpdate READ ME.txt'
    infoLocation = r'Modules\Documents\Access Database Documentation.txt'
