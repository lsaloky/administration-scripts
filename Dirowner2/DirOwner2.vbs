' ----- important constants -----
outputDir = "D:\home\DirOwner"

' ----- input / output files -----
Set objFSI = CreateObject("Scripting.FileSystemObject")
Set objInputFile = objFSI.OpenTextFile ("dir.txt",1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objOutputFile (1000)

' ----- other variables -----
Dim UserList(1000)
maxUser = 0

s = objInputFile.Readline
s = objInputFile.Readline

' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  s = objInputFile.Readline
  If s <> "" Then

' ----- get current directory -----
    If Left(s,10) = " Directory" Then 
      currentdir = Right (s,Len(s)-14)
    Else
      if Left(s,1) <>" " Then

' ----- owner is at position 40, 23 chars. File name is at position 63 ----- 
        owner = Trim(Mid(s,40,23))
        If InStr(owner,"\") > 0 Then owner = Right(owner,Len(owner)-InStr(owner,"\"))
        filename = Trim(Mid(s,63,Len(s)-62))
        size = Trim(Mid(s,25,14))
        If (filename <> ".") and (filename<> "..") and (Mid(s,25,5) <> "<DIR>") Then
 
' ----- check if output file for user already exists -----
          foundInUserList = false
          For i = 0 To maxuser
            If userList(i) = owner Then 
              objOutputFile(i).Writeline (size & " * " & currentdir & "\" & filename)
              foundInUserList = true
              i = maxuser
            End If
          Next

' ----- new user name found: create new output file, write line -----
          If not FoundInUserList Then
            Set objOutputFile(maxuser)=objFSO.CreateTextFile(outputDir & "\" & owner & ".txt")
            objOutputFile(maxuser).Writeline (size & " * " & currentdir & "\" & filename)
            userList(maxuser) = owner
            maxuser = maxuser + 1
          End If
        End If
      End If
    End If
  End If
Loop