# GetGPOsForComputers

Get all group policy objects applied to the computers listed in input file.

1. Enter computer names into Computers.txt
2. Update GetGPOsForComputers.vbs - GetObject("LDAP://cn=Policies,cn=System,dc=subdomain,dc=domain,dc=com") needs to be updated with the correct domain name.
3. Execute GetGPOsForComputers.vbs
4. Output will be stored in file GPOList.txt, format:

GPO name 1: COMPUTERNAME1 COMPUTERNAME2\
GPO name 2: COMPUTERNAME3 COMPUTERNAME4
