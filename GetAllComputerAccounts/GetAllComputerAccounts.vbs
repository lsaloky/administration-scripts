Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("Computer Accounts Last Changed.txt", True)

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = "Select name,whenChanged from 'LDAP://DC=domain,DC=com' Where ((objectCategory='person')(objectClass='computer'))"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = 2 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
objOutputFile.WriteLine ("Computer Name:" & vbTab & "Last changed:")

Do Until objRecordSet.EOF
    objOutputFile.WriteLine objRecordSet.Fields("name").Value & vbTab & objRecordSet.Fields("whenChanged").Value
    objRecordSet.MoveNext
Loop
objOutputFile.Close