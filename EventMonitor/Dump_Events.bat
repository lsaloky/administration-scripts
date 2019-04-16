@echo off

del Daily_Report_Errors.txt
del Daily_Report_Errors_not_filtered.txt
del Daily_Report_Security.txt
del Daily_Report_USB_Usage.txt
del Temp_*.txt
del Computers.txt
del Users_*.txt
del Security.txt

rem ----- get list of all computers from AD -----
csvde -f Temp_OU1.txt -d "OU=Servers,DC=domain,DC=com" -r "(objectClass=computer)" -l "DN"
csvde -f Temp_OU2.txt -d "OU=Desktops,DC=domain,DC=com" -r "(objectClass=computer)" -l "DN"

rem ----- remove all unnecessary characters from csvde output, output file: Computers.txt -----
copy Temp_OU1.txt+Temp_OU2.txt Temp_AllComputers.txt
for /F "tokens=1,2 delims==" %%i in (Temp_AllComputers.txt) do @echo %%j >>Temp_RemovedChars1.txt
for /F "tokens=1,2 delims=," %%i in (Temp_RemovedChars1.txt) do @if not %%j_==_ @echo %%i >>Computers.txt

rem ----- dump all logs, only last day, comma separated, output files: Temp_SEC_*.txt, Temp_APP_*.txt, Temp_SYS_*.txt -----
for /F %%i in (Computers.txt) do (dumpel -f Temp_SEC_%%i.txt -s %%i -l security -t -d %1) & (dumpel -f Temp_APP_%%i.txt -s %%i -l application -t -d %1 -c) & (dumpel -f Temp_SYS_%%i.txt -s %%i -l system -t -d %1 -c)

rem ----- Save only errors from APP, SYS logs + process USB usage -----
for %%q in (Temp_APP_*.txt) do @for /F "tokens=1,2,3,4,5,6,7,8* delims=," %%i in (%%q) do @if %%k==1 @echo %%i %%j %%k %%l %%m %%n %%o %%p %%q >>Daily_Report_Errors_not_filtered.txt
for %%q in (Temp_SYS_*.txt) do @for /F "tokens=1,2,3,4,5,6,7,8* delims=," %%i in (%%q) do @if %%k==1 @echo %%i %%j %%k %%l %%m %%n %%o %%p %%q >>Daily_Report_Errors_not_filtered.txt
for %%q in (Temp_SYS_*.txt) do @for /F "tokens=1,2,3,4,5,6,7,8* delims=," %%i in (%%q) do @if %%m==134 @echo %%i %%j %%p %%q >>Daily_Report_USB_Usage.txt
wscript FilterErrorLogs.vbs

rem ----- prepare for security log analysis: get list of all users from AD, -----
rem ----- create 2 lists for users with common logins during business hours / anytime. Other users are suspicious  -----
ldifde -f Temp_Department24x7_1.txt -d "OU=DepartmentName1,DC=domain,DC=com" -p subtree -r "(objectCategory=CN=Person,CN=Schema,CN=Configuration,DC=adroot,DC=com)" -l "sAMAccountName"
ldifde -f Temp_Department24x7_2.txt -d "OU=DepartmentName1,DC=domain,DC=com" -p subtree -r "(objectCategory=CN=Person,CN=Schema,CN=Configuration,DC=adroot,DC=com)" -l "sAMAccountName"
ldifde -f Temp_Department8x5_1.txt -d "OU=DepartmentName3,DC=domain,DC=com" -p subtree -r "(objectCategory=CN=Person,CN=Schema,CN=Configuration,DC=adroot,DC=com)" -l "sAMAccountName"
ldifde -f Temp_Department8x5_2.txt -d "OU=DepartmentName4,DC=domain,DC=com" -p subtree -r "(objectCategory=CN=Person,CN=Schema,CN=Configuration,DC=adroot,DC=com)" -l "sAMAccountName"
copy Temp_Department8x5_1.txt+Temp_Department8x5_2.txt Temp_Users_BusinessHoursLogon.txt
copy Temp_Department24x7_1.txt+Temp_Department24x7_2.txt Temp_Users_AnyTimeLogon.txt

rem ----- remove all unnecessary characters from ldifde output, create 2 files: for business hours / anytime logon -----
for /F "tokens=1,2 delims=:" %%i in (Temp_Users_BusinessHoursLogon.txt) do @if %%i==sAMAccountName echo %%j >>Users_BusinessHoursLogon.txt
for /F "tokens=1,2 delims=:" %%i in (Temp_Users_AnyTimeLogon.txt) do @if %%i==sAMAccountName echo %%j >>Users_AnyTimeLogon.txt

rem ----- Security logs: process with .vbs file -----
for %%i in (Temp_SEC_*.txt) do @type %%i >>Security.txt
wscript ProcessSecurityLogs.vbs

del Temp_*.txt 
del *.log
del Daily_Report_Errors_not_filtered.txt

rem ----- Output files: Daily_Report_Security.txt, Daily_Report_Errors.txt -----
rem Other files:
rem Computers.txt - list of computers
rem Users_AnyTimeLogon.txt - list of users where all logons are not suspicious - they can log on at anytime
rem Users_BusinessHoursLogon.txt - list of users where logons during business hours are not suspicious
rem logons of all other users are always suspicious
