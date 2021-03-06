strNewGPOPrefix = "TS_"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objGPOList = objFSO.CreateTextFile("GPOList.txt")
Set objRelevantGPOList = objFSO.CreateTextFile("RelevantGPOList.txt")
strCurrentPath = objFSO.GetAbsolutePathName(".")

objGPOList.WriteLine "GPOName,ComputerSettingsEnabled,AtLeastOneComputerSettingConfigured,LinkedToServersOU,LinkedToRootOU,SecurityFilteringGroupsCommaDelimited"

' ----- array for group names (to ensure that one group will be processed only once) and integet to remember how many groups have been added -----
Dim arrGroupName(5000)     ' string, contains group name
Dim arrGroupNameUsed(5000) ' boolean, informs if there is at least one GPO with this group name enabled and with at least one computer setting

For i = 0 To UBound(arrGroupNameUsed)
  arrGroupNameUsed(i) = false
Next

intGroupNamesCounter = 0
intGPOWhereGroupIsUsedCounter = 0

Set objSecurityGroupsFile = objFSO.CreateTextFile("2 - Export.bat")

Set objCopyGPOsFile = objFSO.CreateTextFile("5 - ImportGPOs.bat")
objCopyGPOsFile.WriteLine "@ECHo OFF"
objCopyGPOsFile.WriteLine "C:"
objCopyGPOsFile.WriteLine "CD ""\Program Files\GPMC\Scripts"""

Set objParentDirectory = objFSO.GetFolder(".")
Set objSubfolders = objParentDirectory.SubFolders

' ----- main loop: parse all GPOs, each GPO is in separate subdirectory -----
For Each objSubfolder in objSubfolders
  Set objGPOFile = objFSO.OpenTextFile (".\" & objSubfolder.name & "\gpreport.xml", 1)

  ' ----- prepare variables which remember the current status of .XML file processing -----
  blnJustProcessingComputerSettings = false
  blnJustProcessingUserSettings = false
  blnJustProcessingSecurityDescriptor = false
  blnJustProcessingLinks = false
  blnAtLeastOneSettingInComputerHive = false
  blnLinkedToServersOU = false
  blnLinkedToRootOU = false
  strLinks = ""

  strComputerSettingsEnabled = "Not detected"
  strUserSettingsEnabled = "Not detected"
  strGPOName = ""
  strSecurityFilteringCurrentAccount = ""
  strSecurityFilteringAccounts = ""
 
  ' ----- read .XML file until GPO name found, or end of file reached -----
  Do Until objGPOFile.AtEndOfStream
    strUnicodeLine = objGPOFile.Readline

    ' ----- convert from Unicode to ASCII -----
    strCurrentLine = ""
    For j = 1 To Len (strUnicodeLine) / 2
      strCurrentLine = strCurrentLine + Mid (strUnicodeLine, j * 2, 1)
    Next

    ' ----- remember GPO name from first <Name> item -----
    If strGPOName = "" And InStr(strCurrentLine,"<Name>") > 0 Then
      strGPOName = Mid(strCurrentLine, InStr(strCurrentLine, "<Name>") + 6, InStr(strCurrentLine, "</Name>") - InStr(strCurrentLine, "<Name>") - 6)
    End If
    
    ' ----- remember that we are / are not in computer / user configuration hive, or in security descriptor, or links -----
    If InStr(strCurrentLine,"<Computer>") > 0 Then
      blnJustProcessingComputerSettings = true 
    End If
    If InStr(strCurrentLine,"</Computer>") > 0 Then
      blnJustProcessingComputerSettings = false 
    End If
    If InStr(strCurrentLine,"<User>") > 0 Then
      blnJustProcessingUserSettings = true 
    End If
    If InStr(strCurrentLine,"</User>") > 0 Then
      blnJustProcessingUserSettings = false 
    End If
    If InStr(strCurrentLine,"<SecurityDescriptor>") > 0 Then
      blnJustProcessingSecurityDescriptor = true 
    End If
    If InStr(strCurrentLine,"</SecurityDescriptor>") > 0 Then
      blnJustProcessingSecurityDescriptor = false 
    End If
    If InStr(strCurrentLine,"<LinksTo>") > 0 Then
      blnJustProcessingLinks = true 
    End If
    If InStr(strCurrentLine,"</LinksTo>") > 0 Then
      blnJustProcessingLinks = false 
    End If

    ' ----- check if computer hive is empty, or if there is at least one setting ----- 
    If blnJustProcessingComputerSettings = true And InStr(strCurrentLine,"<ExtensionData>") > 0 Then
      blnAtLeastOneSettingInComputerHive = true
    End If

    ' ----- if first <Enabled> found inside user / computer settings, remember if settings are enabled or not -----
    If InStr(strCurrentLine,"<Enabled>") > 0 Then
      If blnJustProcessingComputerSettings = true And strComputerSettingsEnabled = "Not detected" Then
        If InStr(strCurrentLine,"true") > 0 Then
          strComputerSettingsEnabled = "true"
        End If 
        If InStr(strCurrentLine,"false") > 0 Then
          strComputerSettingsEnabled = "false"
        End If 
      End If
      If blnJustProcessingUserSettings = true And strUserSettingsEnabled = "Not detected" Then
        If InStr(strCurrentLine,"true") > 0 Then
          strUserSettingsEnabled = "true"
        End If 
        If InStr(strCurrentLine,"false") > 0 Then
          strUserSettingsEnabled = "false"
        End If 
      End If
    End If

    ' ----- check if GPO is linked to root OU, or Servers under root ou -----
    If blnJustProcessingLinks = true Then
      If InStr(strCurrentLine,"Servers") > 0 Then
        blnLinkedToServersOU = true
        If InStr(strCurrentLine,"<SOMPath>") > 0 Then
          strLinks = strLinks & "," & Mid(strCurrentLine, 14, Len(strCurrentLine) - 24)
        End If
      End If
      arrCurrentLine = Split(strcurrentLine, "/")
      If UBound(arrCurrentLine) = 3 Then
        blnLinkedToRootOU = true
        strLinks = strLinks & "," & Mid(strCurrentLine, 14, Len(strCurrentLine) - 24)
      End If 
    End If

    ' ----- if "<Name xmlns=..." found while processing security descriptor, remember name -----
    If blnJustProcessingSecurityDescriptor = true Then
      If InStr(strCurrentLine,"<Name xmlns=") > 0 Then
        strSecurityFilteringCurrentAccount = Mid(strCurrentLine, InStr(strCurrentLine,"<Name xmlns=") + 57, InStr(strCurrentLine,"</Name>") - InStr(strCurrentLine,"<Name xmlns=") - 57)
      End If
      
      ' ----- if permission level "Apply Group Policy" found while processing security descriptor, remember group name -----
      If InStr(strCurrentLine,"<GPOGroupedAccessEnum>Apply Group Policy</GPOGroupedAccessEnum>") > 0 Then
        strSecurityFilteringAccounts = strSecurityFilteringAccounts & "," & strSecurityFilteringCurrentAccount

        ' ----- check if group name already added to the list of group names -----
        blnFoundInGroupNamesList = false
        For i = 0 To intGroupNamesCounter - 1
          If arrGroupName(i) = strSecurityFilteringCurrentAccount Then
            blnFoundInGroupNamesList = true
          End If
        Next
        If blnFoundInGroupNamesList = false Then
          arrGroupName(intGroupNamesCounter) = strSecurityFilteringCurrentAccount
          intGroupNamesCounter = intGroupNamesCounter + 1
        End If
      End If
    End If
  Loop
  
  If strGPOName <> "Default Domain Policy" And strGPOName <> "Default Domain Controllers Policy" And Left(strGPOName, 1) = "S" Then
    objGPOList.WriteLine strGPOName & "," & strComputerSettingsEnabled & "," & blnAtLeastOneSettingInComputerHive & "," & blnLinkedToServersOU & "," & blnLinkedToRootOU & strSecurityFilteringAccounts

    If strComputerSettingsEnabled = "true" And blnAtLeastOneSettingInComputerHive = true Then

      If blnLinkedToServersOU = true Or blnLinkedToRootOU = true Then
        objRelevantGPOList.WriteLine strGPOName & strLinks
      End If 

      ' ----- import GPO -----
      objCopyGPOsFile.WriteLine "cscript ImportGPO.wsf """ & strCurrentPath & """ """ & strGPOName & """ """ & strNewGPOPrefix & strGPOName & """"
      objCopyGPOsFile.WriteLine "cscript SetGPOPermissions.wsf """ & strNewGPOPrefix & strGPOName & """ ""Authenticated Users"" /Permission:None" 
      arrSecurityFilteringAccounts = Split(strSecurityFilteringAccounts, ",")
      For i = 1 To UBound(arrSecurityFilteringAccounts) 

        ' ----- set permissions for all accounts -----
        objCopyGPOsFile.WriteLine "cscript SetGPOPermissions.wsf """ & strNewGPOPrefix & strGPOName & """ """ & arrSecurityFilteringAccounts(i) & """ /Permission:Apply" 
 
        ' ----- remember that group name was used at least once in arrGroupNameUsed array -----
        For j = 0 To intGroupNamesCounter - 1
          If arrGroupName(j) = arrSecurityFilteringAccounts(i) Then
            If blnLinkedToServersOU = true Or blnLinkedToRootOU = true Then
              arrGroupNameUsed(j) = true
            End If
          End If
        Next
      Next
    End If
  End If

  Set objGPOFile = Nothing

Next

' ----- prepare group export batch file -----
For i = 0 To intGroupNamesCounter - 1

  If arrGroupNameUsed(i) = true Then

    ' ----- remove domain name from group name -----
    If InStr(arrGroupName(i), "\") > 0 Then
      arrGroupName(i) = Mid(arrGroupName(i), InStr(arrGroupName(i), "\") + 1)
    End If

    ' ----- prepare 2 - Export.bat file -----
    objSecurityGroupsFile.WriteLine "csvde -f Report"& i & ".csv -r ""(sAMAccountName=" & arrGroupName(i) & ")"" -l member"
  End If
Next

MsgBox "Done"