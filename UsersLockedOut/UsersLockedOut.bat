@echo off
ldifde -d OU=Users,dc=domain,dc=com -r "(ObjectCategory=user)" -f Users.txt -l "badPwdCount, badPasswordTime, lockoutTime"
wscript UsersLockedOut.vbs
del Users.txt



