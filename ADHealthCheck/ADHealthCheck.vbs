' ----- ADOdb ----- 
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = "Select name, whenChanged, memberOf, userAccountControl   from 'LDAP://DC=domain,DC=com' Where objectCategory='computer'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

' ----- Output files ----- 
Set subor_computers_group = objFSO.CreateTextFile("AD HCh Computers Not in group Report.txt", True)
Set subor_old_computers = objFSO.CreateTextFile("AD HCh Old Computers Report.txt", True)
Set subor_Psswd_never_expires = objFSO.CreateTextFile("AD HCh Password never expires.txt", True)
Set subor_Unused_homefolder = objFSO.CreateTextFile("AD HCh Unused Homefolder.txt", True)
Set subor_User_without_homefolder = objFSO.CreateTextFile("AD HCh User Without homefolder.txt", True)
Set subor_Members_of_01 = objFSO.CreateTextFile("AD HCh Members of Domain Admin group.txt", True)
Set subor_Members_of_02 = objFSO.CreateTextFile("AD HCh Members of Schema Admin group.txt", True)
Set subor_Members_of_03 = objFSO.CreateTextFile("AD HCh Members of Enterprise Admin group.txt", True)
Set subor_Members_of_04 = objFSO.CreateTextFile("AD HCh Members of Administrator group.txt", True)
Set subor_Roaming_profiles = objFSO.CreateTextFile("AD HCh Roaming Profiles.txt", True)
Set subor_Mail_check = objFSO.CreateTextFile("AD HCh Mail Check.txt", True)
Set subor_Member_of_all = objFSO.CreateTextFile("AD HCh Member of all users.txt", True)

' ----- Home folders ----- 
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace("\\FILESERVER1\Userhomes$")
Set colItems = objFolder.Items

' ----- Roaming profiles ----- 
Set objShell2 = CreateObject("Shell.Application")
Set objFolder2 = objShell2.Namespace("\\FILESERVER1\Profiles$")
Set colItems2 = objFolder2.Items

' ----- 2nd file server, home folders ----- 
Set objShell3 = CreateObject("Shell.Application")
Set objFolder3 = objShell3.Namespace("\\FILESERVER2\Userhomes$")
Set colItems3 = objFolder3.Items

' ----- Computer accounts -----
Class clsComputer
  Public blnIsEnabled
  Public blnIsServer
  Public strName
  Public dateWhenChanged
  Public strMemberOf
End Class

' ----- User accounts -----
Class clsUser
  Public blnPsswdNeverExpires
  Public blnIsEnabled
  Public strName
  Public blnIsUser
  Public strCanonicalName
  Public strHomeDirectory
  Public strLastPartOfHomeDirectory
  Public strProfileDirectory
  Public strLastPartOfProfileDirectory
  Public strGivenName
  Public strSN
  Public strMail
  Public strMailNickName
  Public strFirstPartOfMail
  Public strDistinguishedName
  Public blnIsMemberOfAll
End Class 

' ----- Group accounts -----
Class clsGroup
	Public strName
	Public blnIsMember
	Public intCount

End Class

' ----- Initialize arrays -----
Dim arrComputer(2000)
Dim arrUser(3000)
Dim MyDate
Dim arrGroup(1000)
i = 0
j = 0
MyDate = Date

Function MemberOfAll(strEntryGroup)
	Dim objCommandGroup, objRecordSetGroup, ArrayOfMembers, k
	
	Set objCommandGroup = CreateObject("ADODB.Command")
	Set objCOmmandGroup.ActiveConnection = objConnection

	objCommandGroup.CommandText = "Select member from 'LDAP://"&strEntryGroup&"' Where objectCategory='group'"
	Set objRecordSetGroup = objCommandGroup.Execute
	objRecordSetGroup.MoveFirst

	Do Until objRecordSetGroup.EOF
	ArrayOfMembers = objRecordSetGroup.Fields("member").value

	If IsArray(ArrayOfMembers) Then
		For Each Member in ArrayOfMembers
			For i = 1 To usercount
				'user found, marked as Is member of ALL
				If(arrUser(i).strDistinguishedName = member) then arrUser(i).blnIsMemberOfAll = true		
			Next
			
			For k=0 TO maxGroup - 1
				if member = arrGroup(k).strName then 
					'group found, recursive call for this group
					arrGroup(k).blnIsMember = true
					
					MemberOfAll(member)

				End If
			Next
		Next
	End If	
		
objRecordSetGroup.MoveNext
Loop
End Function

' ----- First line in files : -----
subor_computers_group.WriteLine "Computers which are not member of any COMPUTER group : "
subor_old_computers.WriteLine "Computers which are not in Network for month : "
subor_Psswd_never_expires.WriteLine "Users with Password never Expires Set : "
subor_Unused_homefolder.WriteLine "Users with Inactive Homefolder : "
subor_User_without_homefolder.WriteLine "Users without Homefolder : "
subor_Members_of_01.WriteLine "List of Members of Domain Admin Group : "
subor_Members_of_02.WriteLine "List of Members of Schema Admin Group : "
subor_Members_of_03.WriteLine "List of Members of Enterprise Admin Group : "
subor_Members_of_04.WriteLine "List of Members of Administrator Group : "
subor_Roaming_profiles.WriteLine "List of Members with disabled Roaming Profile, but still active Profile Directory : "
subor_Mail_check.WriteLine "List of Members with bad eMail address : "
subor_Member_of_all.WriteLine "List of user who are NOT members of all users group : "


'----- Script for PC
Do Until objRecordSet.EOF
  i = i +1
  Set arrComputer(i) = New clsComputer

  with arrComputer(i)
	.strName = objRecordSet.Fields("name").value
		If LEFT(.strName, 1) = "Q" OR LEFT(.strName, 1) = "V" Then 
			.blnIsServer = true
		Else .blnIsServer = false
		End IF
	
	.dateWhenChanged = objRecordSet.Fields("whenChanged").value
	.strMemberOf = objRecordSet.Fields("memberOf").value
	
	If objRecordSet.Fields("userAccountControl").value = 4098 Then 
		.blnIsEnabled = false
	Else .blnIsEnabled = true
	End If
  End With

objRecordSet.MoveNext
Loop
computercount = i

For j = 1 To computercount
	with arrComputer(j)
	
		'All computer accouts which are not members of any AD computer group, are not servers and are not disabled
		If (NOT IsArray(.strMemberOf)) AND (.blnIsServer = false) AND (.blnIsEnabled = true) Then 
			subor_computers_group.WriteLine .strName
		End If
		
		'All computer accounts which have not logged on to network within the last 60 days
		If (MyDate - .dateWhenChanged) > 60 Then
			subor_old_computers.WriteLine .strName & " " & CInt(MyDate - .dateWhenChanged) & " days offline" 
		End If
	
	End with
Next

i = 0
Erase arrComputer
' ----- end of script for PC


'----- Script for Users
objCommand.CommandText = "Select name, userAccountControl, homeDirectory, canonicalName, userAccountControl, profilePath, givenName, sn, mail, mailNickname, distinguishedName  from 'LDAP://DC=domain,DC=com' Where objectClass='user' AND objectCategory='user'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	i = i+1
	Set arrUser(i) = new clsUser
	
	with arrUser(i)
		.strName = objRecordSet.Fields("name").value
		.blnPsswdNeverExpires = objRecordSet.Fields("userAccountControl").value AND 65536
		.strHomeDirectory = objRecordSet.Fields("homeDirectory").value
		.strCanonicalName = objRecordSet("canonicalName").value
		.strProfileDirectory = objRecordSet("profilePath").value
		
		If NOT IsNULL(.strHomeDirectory) Then 
			.strLastPartOfHomeDirectory = Mid(.strHomeDirectory, InStrRev(.strHomeDirectory,"\")+1)
		Else .strLastPartOfHomeDirectory = " "
		End If	
       
	   	If NOT IsNULL(.strProfileDirectory) Then 
			.strLastPartOfProfileDirectory = Mid(.strProfileDirectory, InStrRev(.strProfileDirectory,"\")+1)
		Else .strLastPartOfProfileDirectory = " "
		End If
	   
		If InStr(.strCanonicalName(0), "Distribution groups") = 0 AND _
		InStr(.strCanonicalName(0), "Microsoft Exchange System Objects") = 0 AND _
		InStr(.strCanonicalName(0), "Resources") = 0 AND _
		InStr(.strCanonicalName(0), "Users") = 0 AND _
		InStr(.strCanonicalName(0), "Visitors") = 0 AND _
		InStr(.strCanonicalName(0), "Servers") = 0 Then
			.blnIsUser = true
		Else .blnIsUser = false
		End If

		If objRecordSet.Fields("userAccountControl").value = 514 Then 
			.blnIsEnabled = false
		Else .blnIsEnabled = true
		End If
		
		.strGivenName = objRecordSet.Fields("givenName").value
		.strSN = objRecordSet.Fields("sn").value
		.strMail = objRecordSet.Fields("mail").value
		.strMailNickName = objRecordSet.Fields("mailNickname").value
		
		.strDistinguishedName = objRecordSet.Fields("distinguishedName").value
		.blnIsMemberOfAll = false

	End with
	
objRecordSet.MoveNext
Loop
usercount = i

'report all users with Password never expires
For j = 1 TO usercount
	with arrUser(j)
	
		If (.blnPsswdNeverExpires = 65536) Then
		subor_Psswd_never_expires.WriteLine .strName & " " & .blnPsswdNeverExpires
		End If

	End with
Next


'Report all homefolders without user, 1st file server
For Each objItem in colItems
	blnFoundInarrUser = false
	
	RealFolderName = Mid(objItem.Name, InStrRev(objItem.Name,"\")+1)
	
	For j = 1 TO usercount
		with arrUser(j)
			If StrComp(RealFolderName, .strLastPartOfHomeDirectory,1) = 0 AND (.blnIsEnabled = true) Then blnFoundInarrUser = true
		End with	
				
	Next

	If(blnFoundInarrUser = false) Then 
		subor_Unused_homefolder.WriteLine objItem.Name & " - No active user for this home folder"
	End If
	
Next

'Report all homefolders without user, 2nd file server
For Each objItem3 in colItems3
	blnFoundInarrUser = false
	
	RealFolderName = Mid(objItem3.Name, InStrRev(objItem3.Name,"\")+1)
	
	For j = 1 TO usercount
		with arrUser(j)
			If StrComp(RealFolderName, .strLastPartOfHomeDirectory,1) = 0 AND (.blnIsEnabled = true) Then blnFoundInarrUser = true
		End with	
				
	Next

	If(blnFoundInarrUser = false) Then 
		subor_Unused_homefolder.WriteLine objItem3.Name & " - No active user for this home folder"
	End If
	
Next

'Report all users without home folder
For j = 1 TO usercount
	with arrUser(j)
	
		If (.blnIsUser = true) AND (.strLastPartOfHomeDirectory = " ") AND (.blnIsEnabled = true) Then 
			subor_User_without_homefolder.WriteLine .strName & " - user does not have home folder"
		End If

	End with
Next



'Report all roaming profiles, where corresponding user does not have active roaming profile, or user does not exist, or is not active
For Each objItem in colItems2
	blnFoundInarrUser = false
	
	RealFolderName = Mid(objItem.Name, InStrRev(objItem.Name,"\")+1)
	
	For j = 1 TO usercount
		with arrUser(j)
			If StrComp(RealFolderName, .strLastPartOfProfileDirectory,1) = 0 AND (.blnIsEnabled = true) Then blnFoundInarrUser = true
		End with	
				
	Next

	If(blnFoundInarrUser = false) Then 
		subor_Roaming_profiles.WriteLine objItem.Name & " - No active user for this roaming profile, or roaming profile disabled"
	End If
	
Next


'All users with incorrect email address
For j = 1 TO usercount
	with arrUser(j)
	mail = " "
	mailalias = " "
	
	
		If(.blnIsUser = true) AND NOT IsNULL(.strMail) Then
			.strFirstPartOfMail = Mid(.strMail, 1, Instr(.strMail,"@")-1)
			
			If InStr(.strFirstPartOFMail, ".") <> 0 Then
				mail = split(.strFirstPartOfMail, ".", 2)
				
				If(mail(0) <> LCase(.strGivenName)) OR (mail(1) <> LCase(.strSN)) Then 
					subor_Mail_check.WriteLine .strName & " - incorrect email - " & .strMail
				End If
				
			Else subor_Mail_check.WriteLine .strName & " - suspected incorrect email - " & .strMail
			End If
			
			If InStr(.strMailNickName, ".") <> 0 Then
				mailalias = split(.strMailNickName, ".", 2)
				
				If(mailalias(0) <> LCase(.strGivenName)) OR (mail(1) <> LCase(.strSN)) Then
					subor_Mail_check.WriteLine .strName & " - incorrect email alias - " & .strMailNickName
				End If
				
			Else subor_Mail_check.WriteLine .strName & " - suspected incorrct email alias - " & .strMailNickName
			End If
						
		End If
		
	End with
Next


i = 0
' ----- end of script for Users


'----- Script for Member Of Domain Admins
objCommand.CommandText = "Select member from 'LDAP://CN=Domain Admins,CN=Users,DC=domain,DC=com' Where objectCategory='group'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	ArrayOfMembers = objRecordSet.Fields("member").value

	If IsArray(ArrayOfMembers) Then
		For Each Member in ArrayOfMembers
			NameOfMember = Mid(Member, 4, InStr(Member,",")-3)
			subor_Members_of_01.WriteLine NameOfMember
		Next
	End If	
		
objRecordSet.MoveNext
Loop
'----- End of Script


'----- Script for Member Of Schema Admins
objCommand.CommandText = "Select member from 'LDAP://CN=Schema Admins,CN=Users,DC=domain,DC=com' Where objectCategory='group'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	ArrayOfMembers = objRecordSet.Fields("member").value

	If IsArray(ArrayOfMembers) Then
		For Each Member in ArrayOfMembers
			NameOfMember = Mid(Member, 4, InStr(Member,",")-3)
			subor_Members_of_02.WriteLine NameOfMember
		Next
	End If	
		
objRecordSet.MoveNext
Loop
'----- End of Script


'----- Script for Member Of Enterprise Admins
objCommand.CommandText = "Select member from 'LDAP://CN=Enterprise Admins,CN=Users,DC=domain,DC=com' Where objectCategory='group'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	ArrayOfMembers = objRecordSet.Fields("member").value

	If IsArray(ArrayOfMembers) Then
		For Each Member in ArrayOfMembers
			NameOfMember = Mid(Member, 4, InStr(Member,",")-3)
			subor_Members_of_03.WriteLine NameOfMember
		Next
	End If	
		
objRecordSet.MoveNext
Loop
'----- End of Script


'----- Script for Member Of Administrators
objCommand.CommandText = "Select member from 'LDAP://CN=Administrators,CN=Builtin,DC=domain,DC=com' Where objectCategory='group'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
	ArrayOfMembers = objRecordSet.Fields("member").value

	If IsArray(ArrayOfMembers) Then
		For Each Member in ArrayOfMembers
			NameOfMember = Mid(Member, 4, InStr(Member,",")-3)
			subor_Members_of_04.WriteLine NameOfMember
		Next
	End If	
		
objRecordSet.MoveNext
Loop

'----- Script for reading group list and checking if group is member of all users
objCommand.CommandText = "Select distinguishedName from 'LDAP://DC=domain,DC=com' Where objectCategory='group'"
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

i=0

Do Until objRecordSet.EOF
		if Instr(objRecordSet.Fields("distinguishedName").value, "COMPUTERS") = 0 then
		
			Set arrGroup(i) = New clsGroup
			
			arrGroup(i).strName = objRecordSet.Fields("distinguishedName").value
			arrGroup(i).blnIsMember = false

			if(arrGroup(i).strName = "CN=ALL_USERS_GROUP,OU=Distribution groups,DC=domain,DC=com") then arrGroup(i).blnIsMember = true
			i=i+1

		End If
objRecordSet.MoveNext
Loop
maxGroup = i

MemberOfAll("CN=ALL_USERS_GROUP,OU=Distribution groups,DC=domain,DC=com")

'----- End of Script for all users group

for i=1 TO usercount
	
	if(arrUser(i).blnIsMemberOfAll = false) AND (arrUser(i).blnIsUser = true) AND (arrUser(i).strMail <> "") then subor_Member_of_all.WriteLine arrUser(i).strName
	
Next

Erase arrUser
'----- End of Script