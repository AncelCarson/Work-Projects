### Ancel Carson
### 8/9/2019
### Windows 10
### Python command line, Notepad, IDLE
###
### This program will open or run anything
### that relates to the Databases

# Required Libraries: os
# Imported Modules: LoginCredentials, UpdateDatabase, StatusUpdate, ToFlatFile
import os
from Modules import LoginCredentials as LC

databaseLocation = LC.databaseLocation

## Main method of program
def main():
    print('This program can run all Database Applications.')
    print('Please choose a program to run.\n')
    menu()

    input('\nThe program has finished.\nPress enter to close...')


## Menu for chosing which program to run
def menu():
    print("Select the appropriate options from the menu.")
    print("|-----------------------------------|")
    print("|  1) Update from Drilling Info     |")
    print("|  2) Add new wells to databases    |")
    print("|  3) Download Production Flat File |")
    print("|  4) Open Access Database          |")
    print("|  5) Open Status Update workbook   |")
    print("|  6) Open package READ ME          |")
    print("|  7) Open program READ ME          |")
    print("|  8) Open database documentation   |")
    print("|  9) 1 & 2                         |")
    print("| 10) 1 & 3                         |")
    print("| 11) 1 - 3                         |")
    print("| 12) Exit Program                  |")
    print("|-----------------------------------|")
    selected = input("\nType selected number and press Enter...\n")      

    # Selection filter
    if selected == '1':
        from Modules import UpdateDatabases as UD
        UD.main()
    elif selected == '2':
        from Modules import StatusUpdate as SU
        SU.main()
    elif selected == '3':
        from Modules import ToFlatFile as TFF
        TFF.main()
    elif selected == '4':
        os.startfile(databaseLocation)
        menu()
    elif selected == '5':
        os.startfile(r'Modules\Monthly Status Update.xlsx')
        menu()
    elif selected == '6':
        os.startfile(r'READ ME.txt')
        menu()
    elif selected == '7':
        os.startfile(r'Modules\Documents\Database Control READ ME.txt')
        menu()
    elif selected == '8':
        os.startfile(r'Modules\Documents\Access Database Documentation.txt')
        menu()
    elif selected == '9':
        from Modules import StatusUpdate as SU
        SU.main()
        from Modules import UpdateDatabases as UD
        UD.main()
    elif selected == '10':
        from Modules import UpdateDatabases as UD
        UD.main()
        from Modules import ToFlatFile as TFF
        TFF.main()
    elif selected == '11':
        from Modules import StatusUpdate as SU
        SU.main()
        from Modules import UpdateDatabases as UD
        UD.main()
        from Modules import ToFlatFile as TFF
        TFF.main()

        
## Main control
if __name__ == '__main__':
    main()
