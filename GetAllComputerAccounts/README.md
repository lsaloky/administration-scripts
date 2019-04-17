# GetAllComputerAccounts

Script retrieves all computer accounts from Active Directory, including date/time of last connection to network.

1. Update LDAP query "Select name,whenChanged from 'LDAP://DC=domain,DC=com'" with your domain name
2. Execute "GetAllComputerAccounts.vbs"
3. Review output file "Computer Accounts Last Changed.txt"