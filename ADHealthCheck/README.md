# ADHealthCheck

Scritp checks Adtive Directory health. Following issues are being reported:

 - list of computers, which are not members of any Active Directory group
 - list of computers, which have not connected to domain controller for the last 60 days. These accounts might be deleted
 - list of users with password never expires configured
 - list of homefolders without a user
 - list of users without a homefolder
 - list of roaming profiles without user
 - list of users with email address not in syntax firstname.lastname, including aliases
 - list of members of important security groups (Domain Admins, Administrators, Schema Admins, Enterprise Admins)
 - list of all user accounts who are not member of all users group

Steps:

1. Update "ADHealthCheck.bat" with your domain and user name.
2. Update LDAP questies in "ADHealthCheck.vbs": DC=domain,DC=com must match your domain
3. Update paths to shared folders: 
  - objShell.Namespace("\\FILESERVER1\Userhomes$") and objShell3.Namespace("\\FILESERVER2\Userhomes$") must match your file server shared folder for home folders
  - objShell2.Namespace("\\FILESERVER1\Profiles$") must match your file server shared folder for roaming profiles
4. Update part of script which works with all users group: "CN=ALL_USERS_GROUP,OU=Distribution groups,DC=domain,DC=com" must match distinguishedName for your all users group. Needs to be updated on two places
5. Execute "ADHealthCheck.bat"

