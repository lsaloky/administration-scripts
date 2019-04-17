@echo off
ldifde -d OU=Users,dc=domain,dc=com -r "(ObjectCategory=user)" -f Users.txt 
wscript ADCheck.vbs
rem del Users.txt



