rem csvde -f Users.txt -d "OU=Users,DC=domain,DC=com" -r "(objectClass=user)" -l "lastLogon"
ldifde -d OU=Users,dc=domain,dc=com -l pwdLastSet -r "(ObjectCategory=user)" -f Users.txt
pause


