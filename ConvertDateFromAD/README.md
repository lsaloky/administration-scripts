# ConvertDateFromAD

Note that tools 'csvde' and 'ldifde' must be present in current directory in order to execute the scripts.

To process 'password last set':

1. Execute batch file "ListOfUsersPwdLastSet.bat" to export all users
2. Execute script "ConvertDatePwdLastSet.vbs" to convert 'password last set' from numeric format to date

To process 'last logon date':

1. Execute batch file "ListOfUsers.bat" to export all users
2. Execute script "ConvertLogonDate.vbs" to convert 'last logon date' from numeric format to date
