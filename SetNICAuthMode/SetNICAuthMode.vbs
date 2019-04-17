Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' ----- read content of %TEMP% environment variable -----
strTemp = WshShell.ExpandEnvironmentStrings( "%TEMP%" )

' ----- export NIC configuration -----
Const intHiddenWindow = 0
Const blnWaitForExit = True
WshShell.Run "netsh lan export profile folder=""" & strTemp &  """", intHiddenWindow, blnWaitForExit

' ----- read input file Local Area Connection.xml ----
Const intForReading = 1
Const intForWriting = 2
Set objFile = objFSO.OpenTextFile(strTemp & "\Local Area Connection.xml", intForReading)
strText = objFile.ReadAll
objFile.Close

' ----- add <authMode>machine</authMode> using find / replace ------ 
strNewText = Replace(strText, "<cacheUserData>true</cacheUserData>", _
			      "<cacheUserData>true</cacheUserData>" & _
			      VbCrLf & VbTab & VbTab & VbTab & VbTab & _
			      "<authMode>machine</authMode>")

' ----- replace file Local Area Connection.xml with new content----
Set objFile = objFSO.OpenTextFile(strTemp & "\Local Area Connection.xml", intForWriting)
objFile.WriteLine strNewText
objFile.Close

' ----- import NIC configuration -----
WshShell.Run "netsh lan add profile filename=""" & strTemp &  "\Local Area Connection.xml""", intHiddenWindow, blnWaitForExit
