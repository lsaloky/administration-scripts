@echo off
rem ----- usage: 1 command line parameter for Dump_Events.bat, enter number of days -----
runas /profile /env /user:company_administrator_account@domain.com "cmd /c Dump_Events.bat 1"