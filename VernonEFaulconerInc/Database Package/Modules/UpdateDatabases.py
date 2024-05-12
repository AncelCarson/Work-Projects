### Ancel Carson
### 7/30/2019
### Windows 10
### Python command line, Notepad, IDLE
###
### This program will update all fields in the
### DI Databases that reference data from
### Drilling Info

# Required Libraries: os, warnings, pandas, datetime
# Imported Modules: LoginCredentials
print("Loading Libraries.")
import os
import warnings
import pandas as pd
from datetime import datetime

warnings.filterwarnings("ignore")

admin = False
READMELocation = r'Documents\UpdateDatabases READ ME.txt'


## Main method of program
def main():
        
    print("\nThis program can update all DI databases.\n")
    DatabaseSelect = Menu1()

    # Database selection filter
    if DatabaseSelect == '1':
        cursor = LC.SQLcursor
        cnxn = LC.SQLcnxn
        both = False
    elif DatabaseSelect == '2':
        cursor = LC.Accesscursor
        cnxn = LC.Accesscnxn
        both = False
    elif DatabaseSelect == '3':
        cursor = LC.SQLcursor
        cnxn = LC.SQLcnxn
        both = True
    else:
        if __name__ == '__main__':
            LC.CloseConnection()

        # Final message to let user know program is complete
        input('\nThe program has finished.\nPress enter to continue...')
        return 1


    print()
    UpdateSelect = Menu2()

    # Operation selection filter
    if UpdateSelect == '1':
        NewAPICollection(both, cursor, cnxn)
    elif UpdateSelect == '2':
        StatusUpdate(both, cursor, cnxn)
    elif UpdateSelect == '3':
        ProductionUpdate(DatabaseSelect, False, cursor, cnxn)
    elif UpdateSelect == '4':
        NewAPICollection(both, cursor, cnxn)
        StatusUpdate(both, cursor, cnxn)
        ProductionUpdate(DatabaseSelect, False, cursor, cnxn)
    elif UpdateSelect == '5':
        ProductionUpdate(DatabaseSelect, True, cursor, cnxn)

    if __name__ == '__main__':
        LC.CloseConnection()

    # Final message to let user know program is complete
    input('\nThe program has finished.\nPress enter to continue...')


## Menu for selection correct database
def Menu1():

    print("Select the appropriate options from the menus.")
    print("|---------------------------|")
    print("| 1) Update SQL Database    |")
    print("| 2) Update Access Database |")
    print("| 3) Update Both Databases  |")
    print("| 4) Open Program READ ME   |")
    print("| 5) Exit Program           |")
    print("|---------------------------|")
    selected = input("\nType selected number and press Enter...\n")
    if selected == 'admin':
        global admin
        admin = True
        return Menu1()
    elif selected == '4':
        os.startfile(READMELocation)
        return Menu1()
    else:
        return selected
    

## Menu for selecting correct operation
def Menu2():
    print("|--------------------------------|")
    print("| 1) Collect data for new APIs   |")
    print("| 2) Update Statuses             |")
    print("| 3) Collect new Production data |")
    print("| 4) Run 1 - 3                   |")
    print("| 5) Update Production data      |")
    print("| 6) Exit Program                |")
    print("|--------------------------------|")
    selected = input("\nType selected number and press Enter...\n")
    if selected == 'admin':
        global admin
        admin = True
        return Menu2()
    else:
        return selected


## Collecting Public Data for wells recently added
def NewAPICollection(both, cursor, cnxn):
    NewAPINumbers = []
    PublicData = []
    APICount = 0
    
    # Collectiong API Numbers that are in APIs table but not Public Data table
    print("Looking for New API Numbers.")
    # Internal command is SQL
    for APINum in cursor.execute("SELECT API14 FROM Statuses WHERE NOT EXISTS\
        (SELECT API FROM [Public Data] WHERE API = Statuses.API14)"):
        NewAPINumbers.append(APINum)

    # Query run for all APIs found above
    for row in NewAPINumbers:
        # Output of API numbers
        print(NewAPINumbers[APICount][0])
        # Query control for DirectAccess Drilling Info
        for row in LC.d2.query('well-rollups', api14 = NewAPINumbers[APICount][0], \
            fields='WellName,WellNumber,ProductionType,Field,CountyParish,State,\
            DIBasin,ProducingReservoir,DrillType,API14,OperatorAlias,SpudDate,\
            CompletionDate,FirstProdDate,LastProdDate,WellStatus,SurfaceHoleLatitudeWGS84,\
            SurfaceHoleLongitudeWGS84,OilGatherer,GasGatherer,MeasuredDepth,\
            TrueVerticalDepth,GrossPerforatedInterval,UpperPerforation,LowerPerforation'):
            # Query added to Array
            PublicData.append(row)
        APICount += 1

    # Conversion of JSON Format through Pandas
    dfP = pd.DataFrame(PublicData)

    # Convert dates to Date format
    if APICount != 0:
        dfP['SpudDate'] = pd.to_datetime(dfP['SpudDate'])
        dfP['SpudDate'] = dfP['SpudDate'].dt.date
        dfP['CompletionDate'] = pd.to_datetime(dfP['CompletionDate'])
        dfP['CompletionDate'] = dfP['CompletionDate'].dt.date
        dfP['FirstProdDate'] = pd.to_datetime(dfP['FirstProdDate'])
        dfP['FirstProdDate'] = dfP['FirstProdDate'].dt.date
        dfP['LastProdDate'] = pd.to_datetime(dfP['LastProdDate'])
        dfP['LastProdDate'] = dfP['LastProdDate'].dt.date

    dfPF = dfP.fillna("")
    dfPF = dfPF.drop_duplicates()

    # Adding Drilling Info Query to Access database
    print("Adding any New API Numbers.")
    for index,row in dfPF.iterrows():
        WellName = str(row['WellName'])
        if ";" in WellName:
            WellName = WellName.split(";",1)[1]
        AssetName = WellName + " " + row['WellNumber']
        if both == True:
            # Internal Commands are SQL
            LC.SQLcursor.execute("INSERT INTO [Public Data]\
                (ASSETNAME,LEASE,WELLNUM,MAJOR,FIELD,COUNTY,STATE,\
                DI_BASIN,RESERVOIR,DRILL_TYPE,API,OPERATOR,SPUD_DATE,\
                COMP_DATE,FIRST_PROD,LAST_PROD,DI_STATUS,LATITUDE,\
                LONGITUDE,OIL_GATHERER,GAS_GATHERER,MD,TVD,\
                PERF_INTERVAL,UPPER_PERF,LOWER_PERF) \
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",\
                AssetName,str(row['WellName']),str(row['WellNumber']),\
                str(row['ProductionType']),str(row['Field']),str(row['CountyParish']),\
                str(row['State']),str(row['DIBasin']),str(row['ProducingReservoir']),\
                str(row['DrillType']),str(row['API14']),str(row['OperatorAlias']),\
                str(row['SpudDate']),str(row['CompletionDate']),str(row['FirstProdDate']),\
                str(row['LastProdDate']),str(row['WellStatus']),str(row['SurfaceHoleLatitudeWGS84']),\
                str(row['SurfaceHoleLongitudeWGS84']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['MeasuredDepth']),str(row['TrueVerticalDepth']),\
                str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),str(row['LowerPerforation']))
            LC.Accesscursor.execute("INSERT INTO [Public Data]\
                (ASSETNAME,LEASE,WELLNUM,MAJOR,FIELD,COUNTY,STATE,\
                DI_BASIN,RESERVOIR,DRILL_TYPE,API,OPERATOR,SPUD_DATE,\
                COMP_DATE,FIRST_PROD,LAST_PROD,DI_STATUS,LATITUDE,\
                LONGITUDE,OIL_GATHERER,GAS_GATHERER,MD,TVD,\
                PERF_INTERVAL,UPPER_PERF,LOWER_PERF) \
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",\
                AssetName,str(row['WellName']),str(row['WellNumber']),\
                str(row['ProductionType']),str(row['Field']),str(row['CountyParish']),\
                str(row['State']),str(row['DIBasin']),str(row['ProducingReservoir']),\
                str(row['DrillType']),str(row['API14']),str(row['OperatorAlias']),\
                str(row['SpudDate']),str(row['CompletionDate']),str(row['FirstProdDate']),\
                str(row['LastProdDate']),str(row['WellStatus']),str(row['SurfaceHoleLatitudeWGS84']),\
                str(row['SurfaceHoleLongitudeWGS84']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['MeasuredDepth']),str(row['TrueVerticalDepth']),\
                str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),str(row['LowerPerforation']))
        else:
            cursor.execute("INSERT INTO [Public Data]\
                (ASSETNAME,LEASE,WELLNUM,MAJOR,FIELD,COUNTY,STATE,\
                DI_BASIN,RESERVOIR,DRILL_TYPE,API,OPERATOR,SPUD_DATE,\
                COMP_DATE,FIRST_PROD,LAST_PROD,DI_STATUS,LATITUDE,\
                LONGITUDE,OIL_GATHERER,GAS_GATHERER,MD,TVD,\
                PERF_INTERVAL,UPPER_PERF,LOWER_PERF) \
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",\
                AssetName,str(row['WellName']),str(row['WellNumber']),\
                str(row['ProductionType']),str(row['Field']),str(row['CountyParish']),\
                str(row['State']),str(row['DIBasin']),str(row['ProducingReservoir']),\
                str(row['DrillType']),str(row['API14']),str(row['OperatorAlias']),\
                str(row['SpudDate']),str(row['CompletionDate']),str(row['FirstProdDate']),\
                str(row['LastProdDate']),str(row['WellStatus']),str(row['SurfaceHoleLatitudeWGS84']),\
                str(row['SurfaceHoleLongitudeWGS84']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['MeasuredDepth']),str(row['TrueVerticalDepth']),\
                str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),str(row['LowerPerforation']))

    # Commit changes to databases
    if both == True:
        LC.SQLcnxn.commit()
        LC.Accesscnxn.commit()
    else:
        cnxn.commit()


## Updating information of all wells in database
def StatusUpdate(both, cursor, cnxn):
    APINumbers = []
    WellStatuses = []
    APICount = 0
    outputCount = 0

    # Collection all APIs found in the Public Data table
    print("Collecting public statuses for API Numbers.")
    # Internal command is SQL
    for APINum in cursor.execute("SELECT API FROM [Public Data]"):
        APINumbers.append(APINum)

    # Query run for all APIs found above
    for row in APINumbers:
        # Message printed to show user program is still running
        if APICount % 100 == 0:
            print("Collecting Data...")
            if admin == True:
                print(outputCount,":",datetime.now().strftime("%H:%M:%S"))
            outputCount += 1
        # Query control for DirectAccess Drilling Info
        for row in LC.d2.query('well-rollups', api14 = APINumbers[APICount][0], \
            fields='API14,OperatorAlias,ProducingReservoir,WellStatus,\
            LastProdDate,OilGatherer,GasGatherer,GrossPerforatedInterval,\
            UpperPerforation,LowerPerforation'):
            # Query added to Array
            WellStatuses.append(row)
        APICount += 1

    # Conversion of JSON Format through Pandas
    dfS = pd.DataFrame(WellStatuses)

    # Convert dates to Date format
    dfS['LastProdDate'] = pd.to_datetime(dfS['LastProdDate'])
    dfS['LastProdDate'] = dfS['LastProdDate'].dt.date

    dfSF = dfS.fillna("")

    print("Updating status for all wells.")
    for index,row in dfSF.iterrows():
        if both == True:
            # Internal command is SQL
            LC.SQLcursor.execute("UPDATE [Public Data] SET OPERATOR = ?, RESERVOIR = ?, DI_STATUS = ?, \
                LAST_PROD = ?, OIL_GATHERER = ?, GAS_GATHERER = ?, PERF_INTERVAL = ?, UPPER_PERF = ?, \
                LOWER_PERF = ? WHERE API = ?", str(row['OperatorAlias']),str(row['ProducingReservoir']),\
                str(row['WellStatus']),str(row['LastProdDate']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),\
                str(row['LowerPerforation']),str(row['API14']))
            LC.Accesscursor.execute("UPDATE [Public Data] SET OPERATOR = ?, RESERVOIR = ?, DI_STATUS = ?, \
                LAST_PROD = ?, OIL_GATHERER = ?, GAS_GATHERER = ?, PERF_INTERVAL = ?, UPPER_PERF = ?, \
                LOWER_PERF = ? WHERE API = ?", str(row['OperatorAlias']),str(row['ProducingReservoir']),\
                str(row['WellStatus']),str(row['LastProdDate']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),\
                str(row['LowerPerforation']),str(row['API14']))
        else:
            cursor.execute("UPDATE [Public Data] SET OPERATOR = ?, RESERVOIR = ?, DI_STATUS = ?, \
                LAST_PROD = ?, OIL_GATHERER = ?, GAS_GATHERER = ?, PERF_INTERVAL = ?, UPPER_PERF = ?, \
                LOWER_PERF = ? WHERE API = ?", str(row['OperatorAlias']),str(row['ProducingReservoir']),\
                str(row['WellStatus']),str(row['LastProdDate']),str(row['OilGatherer']),\
                str(row['GasGatherer']),str(row['GrossPerforatedInterval']),str(row['UpperPerforation']),\
                str(row['LowerPerforation']),str(row['API14']))

    # Commit changes to databases
    if both == True:
        LC.SQLcnxn.commit()
        LC.Accesscnxn.commit()
    else:
        cnxn.commit()


## Adding Production Data from DI to databases
def ProductionUpdate(choice, update, cursor, cnxn):
    ActiveAPINumbers = []
    MonthlyData = []
    EntryIDs = []
    inMonthly = False
    APICount = 0
    InputCount = 0
    outputCount = 0

    # Collection all APIs found in the Public Data table
    print("Collecting Monthly Production all wells.")
    # Internal command is SQL
    for APINum in cursor.execute("SELECT API FROM [Public Data]"):
        ActiveAPINumbers.append(APINum)

    # Query run for all APIs found above
    for row in ActiveAPINumbers:
        # Message printed to show user program is still running
        if APICount % 100 == 0:
            print("Collecting monthly data...")
            if admin == True:
                print(outputCount,":",datetime.now().strftime("%H:%M:%S"))
            outputCount += 1
        # Query control for DirectAccess Drilling Info
        for row in LC.d2.query('producing-entity-details', apino = ActiveAPINumbers[APICount][0], \
            fields='PdenProdId,ApiNo,ProdDate,ProdMonthNo,Gas,Liq,Wtr,UpdatedDate'):
            MonthlyData.append(row)
        APICount += 1

    outputCount = 0
    
    # Conversion of JSON Format through Pandas
    dfM = pd.DataFrame(MonthlyData)
    # Convert dates to Date format
    dfM['ProdDate'] = pd.to_datetime(dfM['ProdDate'])
    dfM['ProdDate'] = pd.DatetimeIndex(dfM['ProdDate']).to_period('M').to_timestamp('M')
    dfM['ProdDate'] = dfM['ProdDate'].dt.date
    dfM['UpdatedDate'] = pd.to_datetime(dfM['UpdatedDate'])
    dfM['UpdatedDate'] = dfM['UpdatedDate'].dt.date
    # Removes all NaNs from data
    dfMF = dfM.fillna(0)
    # Sorting Data to speed up program
    dfMS = dfMF.sort_values(by=['PdenProdId'])

    # Deletes all Production Data to download fresh from DI
    if update == True:
        if choice == '3':
            LC.SQLcursor.execute("DELETE FROM [Production Monthly]")
            LC.Accesscursor.execute("DELETE FROM [Production Monthly]")
        else:
            cursor.execute("DELETE FROM [Production Monthly]")

    # Retrieving the IDs of all records in Production Monthly table
    print("Uploading most recent well production data.")
    # Internal command is SQL
    for EntryID in cursor.execute("SELECT ENTRY_ID FROM [Production Monthly]"):
        EntryIDs.append(EntryID[0])

    # Sorting retrieved IDs
    EntryIDs.sort()

    # Writing all new Drilling Info Entries to Access database
    for index,row in dfMS.iterrows():
        inMonthly = False
        count = 0
        # Message printed to show user program is still running
        if InputCount % 10000 == 0:
            print("Updating Database...")
            if admin == True:
                print(outputCount,":",datetime.now().strftime("%H:%M:%S"))
            outputCount += 1
        for EntryID in EntryIDs:
            if update == True:
                break
            # Checking for ID match
            if row['PdenProdId'] == EntryID:
                # If a match is found the ID is deleted from EntryIDs
                inMonthly = True
                del EntryIDs[count]
                break
            count +=1
        if inMonthly == False:
            if choice == '3':
                # Internal command is SQL
                LC.SQLcursor.execute("INSERT INTO [Production Monthly]\
                    (ENTRY_ID,API_NO,PROD_DATE,PROD_MONTH_NO,GAS,OIL,WATER,UPDATE_DATE) \
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)", str(row['PdenProdId']),str(row['ApiNo']),\
                    row['ProdDate'],str(row['ProdMonthNo']),str(row['Gas']),str(row['Liq']),\
                    str(row['Wtr']),row['UpdatedDate'])
                # Internal command is SQL
                LC.Accesscursor.execute("INSERT INTO [Production Monthly]\
                    (ENTRY_ID,API_NO,PROD_DATE,PROD_MONTH_NO,GAS,OIL,WATER,UPDATE_DATE) \
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)", str(row['PdenProdId']),str(row['ApiNo']),\
                    row['ProdDate'],str(row['ProdMonthNo']),str(row['Gas']),str(row['Liq']),\
                    str(row['Wtr']),row['UpdatedDate'])
            else:
                # Internal command is SQL
                cursor.execute("INSERT INTO [Production Monthly]\
                    (ENTRY_ID,API_NO,PROD_DATE,PROD_MONTH_NO,GAS,OIL,WATER,UPDATE_DATE) \
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)", str(row['PdenProdId']),str(row['ApiNo']),\
                    row['ProdDate'],str(row['ProdMonthNo']),str(row['Gas']),str(row['Liq']),\
                    str(row['Wtr']),row['UpdatedDate'])
        InputCount += 1

    # Commit changes to databases
    if choice == '3':
        LC.SQLcnxn.commit()
        LC.Accesscnxn.commit()
    else:
        cnxn.commit()


## Main control
if __name__ == '__main__':
    import LoginCredentials as LC
    main()
else:
    from Modules import LoginCredentials as LC
    READMELocation = r'Modules\Documents\UpdateDatabases READ ME.txt'
