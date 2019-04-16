Dim arrOUName(100)     			' array of strings, contains OU names
Dim arrOUContent(100)			' list of computers found in this OU
Dim arrComputersToSearchFor(1000)	' array of strings, computer names to search for
Dim arrComputerFound(1000)		' indicates whether computer has been found

intOUCounter = 0
intComputerCounter = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objSearchedComputersFile = objFSO.OpenTextFile ("ComputersToSearchFor.txt", 1)
Set objOutputFile = objFSO.CreateTextFile("OUs.txt")
Set objAllComputersFile = objFSO.OpenTextFile ("AllComputers.txt", 1)
strLine = objAllComputersFile.Readline ' ignore first line
Set objErrorFile = objFSO.CreateTextFile("NotFound.txt")

' ----- read all computers to search for -----
Do Until objSearchedComputersFile.AtEndOfStream
  strLine = objSearchedComputersFile.Readline
  arrComputersToSearchFor(intComputerCounter) = strLine
  intComputerCounter = intComputerCounter + 1
Loop

' ----- main loop -----
Do Until objAllComputersFile.AtEndOfStream
 
  strLine = objAllComputersFile.Readline

  If strLine <> "" Then

    ' ----- get computer name and OU name from "CN=COMPUTERNAME,OU=OUCHILD,OU=OUPARENT,DC=subdomain,DC=domain,DC=com" -----
    strComputerName = Mid(strLine, 5, InStr(strLine, ",") - 5)     
    strOUName = ""
    If InStr(strLine, "OU=") = 0 Then
      strOUName = "Computers"
    Else
      strOUName = Mid(strLine, InStr(strLine, ",") + 1, InStr(strLine, ",DC=") - InStr(strLine, ",") - 1) 
    End If 

    ' ----- check if computer name is on the list of computers to search for -----
    blnFoundInComputersList = false 
    For i = 0 To intComputerCounter - 1 
      If arrComputersToSearchFor(i) = strComputerName Then 
        arrComputerFound(i) = "True"
        blnFoundInComputersList = true
      End If
    Next

    If blnFoundInComputersList = true Then

      ' ----- if yes, check if OU is already on the list -----
      blnFoundInOUList = false 
      For i = 0 To intOUCounter - 1 
        If arrOUName(i) = strOUName Then 
          blnFoundInOUList = true
          arrOUContent(i) = arrOUContent(i) & "," & strComputerName
        End If
      Next

      ' ----- add to list of OUs, if not found -----
      If blnFoundInOUList = false Then
        arrOUName(intOUCounter) = strOUName 
        arrOUContent(i) = strComputerName
        intOUCounter = intOUCounter + 1
      End If
  
    End If

  End If

Loop

' ----- write results to output file -----
For i = 0 To intOUCounter - 1 
  objOutputFile.WriteLine arrOUName(i) & vbTab & arrOUContent(i)
Next

For i = 0 To intComputerCounter - 1 
  If arrComputerFound(i) <> "True" Then
    objErrorFile.WriteLine arrComputersToSearchFor(i)
  End If
Next

MsgBox "Done"