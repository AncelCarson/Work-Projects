import pyodbc
serial = '100170'
# cnxn = pyodbc.connect("DSN=SOTAMAS90;Directory=\\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90;SERVER=ntc-app1;UID=acarson;PWD=Carson123;CompanyCode=TST", autocommit=True)
# cnxn = pyodbc.connect("DSN=SOTAMAS90;Directory=\\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90;SERVER=ntc-app1;UID=acarson;PWD=Carson123;CompanyCode=NTC")
# cnxn = pyodbc.connect("DSN=SOTAMAS90; Directory=\\NTC-app1\SageSoftware\Sage 100 Standard V2021\MAS90\; Prefix=\\NTC-app1\SageSoftware\Sage 100 Standard V2021\MAS90\SY,\\NTC-app1\SageSoftware\Sage 100 Standard V2021\MAS90\==; ViewDLL=\\NTC-app1\SageSoftware\Sage 100 Standard V2021\MAS90\Home; LogFile=\PVXODBC.LOG; CacheSize=4; DirtyReads=1; BurstMode=1; StripTrailingSpaces=1; SERVER=NotTheServer;UID=acarson;PWD=Carson123;Company=NTC")
# cnxn = pyodbc.connect("DSN=SOTAMAS90;Directory=\\ntc-app1\SageSoftware\v530\MAS90;prefix=\\ntc-app1\SageSoftware\v530\MAS90\SY\, \\ntc-app1\SageSoftware\v530\MAS90\==\;viewdll=\\ntc-app1\SageSoftware\v530\MAS90\home;SERVER=ntc-app1;UID=acarson;PWD=Carson123;Company=NTC")
# cnxn = pyodbc.connct("DSN=eRemoteMas90; Directory=\\ntc-app1\SageSoftware\v530\MAS90; Prefix=\\ntc-app1\SageSoftware\v530\MAS90\SY\, \\ntc-app1\SageSoftware\v530\MAS90\==\; ViewDLL=\\ntc-app1\SageSoftware\v530\MAS90\HOME; LogFile=\PVXODBC.LOG; DirtyReads=1; BurstMode=1; StripTrailingSpaces=1; SERVER=ntc-app1;UID=acarson;PWD=Carson123;Company=NTC")

# cnxn = pyodbc.connect(r"DSN=RemoteMas90; Directory=\\ntc-app1\SageSoftware\v530\MAS90; Prefix=\\ntc-app1\SageSoftware\v530\MAS90\SY\, \\ntc-app1\SageSoftware\v530\MAS90\==\; ViewDLL=\\ntc-app1\SageSoftware\v530\MAS90\HOME; LogFile=\PVXODBC.LOG; DirtyReads=1; BurstMode=1; StripTrailingSpaces=1; SERVER=NotTheServer;UID=acarson;PWD=Carson123;Company=NTC", autocommit=True)
cnxn = pyodbc.connect(r"DSN=RemoteMas90; Directory=\\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90; Prefix=\\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90\SY\, \\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90\==\; ViewDLL=\\ntc-app1\SageSoftware\Sage 100 Standard V2021\MAS90\HOME; LogFile=\PVXODBC.LOG; DirtyReads=1; BurstMode=1; StripTrailingSpaces=1; SERVER=NotTheServer;UID=acarson;PWD=Carson123;Company=NTC", autocommit=True)

cnxn.close()

