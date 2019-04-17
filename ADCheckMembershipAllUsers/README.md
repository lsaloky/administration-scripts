# ADCheckMembershipAllUsers

Check if all user accounts with mailbox, excluding service accounts, are members of 'all users' group

Note that 'ldifde' tool must be present in current directory in order to execute the scripts.

1. Update Select command in "ADCheck.vbs" to ensure that all groups in the domain, which are members of 'all users' group, are correctly processed. Each group should be mentioned in a row like this: 
```Case " CN=Group1," isMemberOfKosiceAllUsers = true```
2. Execute "ADCheck.bat"
3. Output file "ADCheckReportAllUsersMembership.txt" will contain user accounts, which are not members of 'all users' group.