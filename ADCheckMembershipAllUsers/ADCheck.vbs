' ----- input / output file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Users.txt", 1)
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("ADCheckReportAllUsersMembership.txt")

' ----- main loop -----

objOutputFile.WriteLine ("Users with mailbox, who are not members of all users group:")
hasMailbox = false
isMemberOfKosiceAllUsers = false

Do Until objInputFile.AtEndOfStream

  line = objInputFile.Readline

' ----- get user name -----
  If Left(line,7) = "dn: CN=" Then
    userName = Mid(line, 8, InStr(line,",")-8)
  End If

' ----- check if user is member of KosiceAllUsers -----
  groupName = Left(line,InStr(line,",")) 
  Select Case groupName
    Case " CN=AllUsers," isMemberOfKosiceAllUsers = true
    Case " CN=Group1," isMemberOfKosiceAllUsers = true
    Case " CN=Group2," isMemberOfKosiceAllUsers = true
  End Select

' ----- check if user has a mailbox -----
  If Left(line,8) = "homeMDB:" Then
    hasMailbox=True
  End If 

' ----- end of user data, analyze and write if necessary -----
  If line = "" Then
    If hasMailbox Then
      count = count+1
      If not isMemberofKosiceAllUsers Then
        objOutputFile.WriteLine (userName) 
      End If
    End If
    hasMailbox = false
    isMemberOfKosiceAllUsers = false
  End If
Loop
objOutputFile.WriteLine ("Total mailbox count: " & count) 

