' ----- input / output file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Users.txt", 1)
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("UsersLockedOut.txt")

' ----- main loop -----

objOutputFile.WriteLine ("Username; lockoutTime")
Do Until objInputFile.AtEndOfStream

  line = objInputFile.Readline

  If InStr (line,"MMS") = 0 Then ' ignore MMS OU
 
    If Left(line,2) = "dn" Then

' ----- if there is failure logon attempt, write user ----- 
      If badPwdCount = 5 Then objOutputFile.WriteLine (currentUser & "; " & lockoutTime) 
      badPwdCount = 0
      lockoutTime = 0
      currentUser = line
    End If

    If Left(line,11) = "badPwdCount" Then
      badPwdCount = Mid(line,13,255)
    End If 

    If Left(line,11) = "lockoutTime" Then
      lockoutTime = #1/1/1601# + Mid(line,13,255)/600000000/1440
    End If 

  Else
    Do Until line = "" ' find next user after ignoring MMS user
     line = objInputFile.Readline     
    Loop
  End If
Loop
If badPwdCount = 5 Then objOutputFile.WriteLine (currentUser & "; " & lockoutTime) 
