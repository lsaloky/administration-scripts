strGroupNamePrefix = "TS_"
strOUForNewGroups = "OU=Test,OU=New"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objParentDirectory = objFS.GetFolder(".")
Set colFiles = objParentDirectory.Files

' ----- array for group names and array of booleans to remember if group name contains at least one computer account -----
Dim arrGroupName(5000)     ' string, contains group name
Dim arrGroupContainsComputerAccount(5000) ' boolean, informs if there is at least one computer account in group

For i = 0 To 5000 
  arrGroupContainsComputerAccount(i) = false
Next

intGroupNameCounter = 0

' ----- array for relevant GPOs from RelevantGPOs.txt -----
Dim arrRelevantGPO(5000)
intRelevantGPOCounter = 0

Set objInputFile = objFS.OpenTextFile ("RelevantGPOList.txt", 1)
Do Until objInputFile.AtEndOfStream
  arrRelevantGPO(intRelevantGPOCounter) = objInputFile.ReadLine
  intRelevantGPOCounter = intRelevantGPOCounter + 1
Loop


' ----- parse all text files, modify beginning of second line (add prefix with new group name) -----
For Each objFile in colFiles
  If Left(objFile.Name, 6) = "Report" Then

    Set objInputFile = objFS.OpenTextFile (objFile.Name, 1)

    strLine = objInputFile.ReadLine

    ' ----- if there are no members, do not create any .LDF file -----
    If strLine <> "DN,(null)" Then
  
      Set objOutputFile = objFS.CreateTextFile ("New" & objFile.Name & ".ldf")

      strLine = objInputFile.ReadLine
     
      ' ----- get group name and members from second line -----
      arrStringElements = Split(strLine, """")
      strGroupDN = arrStringElements(1)
      strMembers = arrStringElements(3)
      arrMember = Split(strMembers, ";")

      strDomainInGroupDN = Mid(strGroupDN, InStr(strGroupDN, "DC="))
      strGroupNameInGroupDN = Mid(strGroupDN, 4, InStr(strGroupDN, ",") - 4)

      arrGroupName(intGroupNameCounter) = strGroupNameInGroupDN

      strNewGroupName = "CN=" & strGroupNamePrefix & strGroupNameInGroupDN & "," & strOUForNewGroups & "," & strDomainInGroupDN
    
      objOutputFile.WriteLine "DN: " & strNewGroupName
      objOutputFile.WriteLine "changeType: modify"
      objOutputFile.WriteLine "add: member"
      For i = 0 To UBound(arrMember)
        objOutputFile.WriteLine "member: " & arrMember(i)
        ' ----- check if there is at least one computer account (15 characters name) -----
        If InStr(arrMember(i), ",") = 19 Then
          arrGroupContainsComputerAccount(intGroupNameCounter) = true
        End If 
      Next
      objOutputFile.WriteLine "-"
      
      Set objInputFile = Nothing
      Set objOutputFile = Nothing

      intGroupNameCounter = intGroupNameCounter + 1

    End If
  End If
Next

' ----- open GPOList, filter out GPOs without computer settings, with computer settings disabled, -----
' ----- without links to Servers or root OUs and without computer account in security filtering groups -----
Set objInputFile = objFS.OpenTextFile ("GPOList.txt", 1)
strLine = objInputFile.ReadLine ' ignore first line
Set objOutputFile = objFS.CreateTextFile ("MostRelevantGPOList.txt")
Do Until objInputFile.AtEndOfStream
  strLine = objInputFile.ReadLine
  arrElements = Split(strLine, ",")
  strGPOName = arrElements(0)
  
  If LCase(arrElements(1)) = "true" Then
    blnComputerSettingsEnabled = true
  Else
    blnComputerSettingsEnabled = false
  End If
  
  If LCase(arrElements(2)) = "true" Then
    blnAtLeastOneComputerSettingConfigured = true
  Else
    blnAtLeastOneComputerSettingConfigured = false
  End If

  If LCase(arrElements(3)) = "true" Then
    blnLinkedToServersOU = true
  Else
    blnLinkedToServersOU = false
  End If

  If LCase(arrElements(4)) = "true" Then
    blnLinkedToRootOU = true
  Else
    blnLinkedToRootOU = false
  End If

  blnAtLeastOneGroupContainsComputerAccount = false
  For i = 5 To UBound(arrElements)
    strGroupName = Mid(arrElements(i), InStr(arrElements(i),"\") + 1)
    For j = 0 To intGroupNameCounter - 1 
      If arrGroupName(j) = strGroupName Then
        If arrGroupContainsComputerAccount(j) = true Then
          blnAtLeastOneGroupContainsComputerAccount = true
        End If
      End If
    Next
  Next

  ' ----- write results to MostRelevantGPOList.txt -----
  If blnComputerSettingsEnabled = true And blnAtLeastOneComputerSettingConfigured = true And blnAtLeastOneGroupContainsComputerAccount = true Then
    If blnLinkedToServersOU = true or blnLinkedToRootOU = true Then
      
      ' ----- search for GPO name in RelevantGPOList.txt to copy links too -----
      For i = 0 To intRelevantGPOCounter - 1
        If InStr(arrRelevantGPO(i), strGPOName) = 1 Then
          objOutputFile.WriteLine arrRelevantGPO(i)
        End If
      Next
    End If
  End If

Loop

MsgBox "Done"