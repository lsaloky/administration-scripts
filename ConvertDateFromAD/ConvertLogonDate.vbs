' ----- input / output file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Users.txt", 1)
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("UsersConverted.txt")

' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  line = objInputFile.Readline
  If Left(line,2) = "dn" Then
    currentUser = line
  End If
  if Left(line,9) = "lastLogon" Then
    objOutputFile.WriteLine currentUser & " ; " & #1/1/1601# + Mid(line,11,255)/600000000/1440
  End If 
Loop
