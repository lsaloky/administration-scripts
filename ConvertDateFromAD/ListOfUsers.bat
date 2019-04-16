csvde -f Users.txt -d "OU=Users,OU=KOS,OU=EUR,DC=molex,DC=com" -r "(objectClass=user)" -l "lastLogon"
ldifde -d OU=Users,OU=KOS,OU=EUR,dc=molex,dc=com -l lastLogon -r "(ObjectCategory=user)" -f Users.txt
pause


