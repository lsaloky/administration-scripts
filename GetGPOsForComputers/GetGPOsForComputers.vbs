On Error Resume Next

' ----- class for GPO attributes -----
Class clsGPO
  Public strDisplayName
  Public strDistinguishedName   ' stores string in lowercase
  Public blnFound
  Public strComputerNames
End Class

' ----- function to get list of GPOs applied on OU (modify arrGPO) -----
Function GetListOfGPOs(strOU)
  Set objContainer = GetObject (strOU)

  strGpLink = "none"
  strGpLink = objContainer.Get("gPLink")

  If strGpLink <> "none" And strGpLink <> "" And strGpLink <> " " Then
    arrGpLinkItems = Split(strGpLink,"]")
    For i = UBound(arrGPLinkItems) To LBound(arrGpLinkItems) + 1 Step -1
      arrGPLink = Split(arrGpLinkItems(i-1),";")
      strDNGPLink = Mid(arrGPLink(0),9)
      strDNGPLink = LCase(strDNGPLink)

      ' ----- look for DN in list of GPOs, retrieve strDisplayName -----
      For j = 0 To intGPOCounter - 1
        If strDNGPLink = arrGPO(j).strDistinguishedName Then
          arrGPO(j).blnFound = True
          arrGPO(j).strComputerNames = arrGPO(j).strComputerNames & " " & strComputerName
        End If
      Next
    Next
  End If
End Function

' ----- function to get computer DN from name -----
Function GetComputerDN(strComputerName)
  Set objTrans = CreateObject("NameTranslate")
  objTrans.Init 1, strDomain
  objTrans.Set 3, strDomain & "\" & strComputerName & "$"
  strComputerDN = objTrans.Get(1) 
  GetComputerDN = strComputerDN
End Function

' ------------------------
' ----- Main program -----
' ------------------------

' -----initialize variables -----
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim arrGPO(5000)

Set objListOfGPOs = objFSO.CreateTextFile("GPOList.txt")
Set objListOfComputers = objFSO.OpenTextFile("computers.txt", 1)

Set objSystemInfo = CreateObject("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName

' ----- Get a list of Group Policy Objects -----
Set objGPOs = GetObject("LDAP://cn=Policies,cn=System,dc=subdomain,dc=domain,dc=com")

intGPOCounter = 0
For Each GPO in objGPOs
  Set arrGPO(intGPOCounter) = New clsGPO
  ' objListOfGPOs.WriteLine "Logging: Read GPO " & GPO.DisplayName
  With arrGPO(intGPOCounter)
    .strDisplayName = GPO.DisplayName
    .strDistinguishedName = LCase (GPO.DistinguishedName)
    .blnFound = False
    .strComputerNames = ""
  End With
  intGPOCounter = intGPOCounter + 1
Next

' MsgBox "List of GPOs read, click OK to continue."

' ----- main loop: parse list of computers -----
Do Until objListOfComputers.AtEndOfStream
  strComputerName = objListOfComputers.Readline
  
  ' ----- get distinguished name, get domain part of DN,
  strComputerDN = GetComputerDN(strComputerName)
  strDomainPartOfDN = Mid (strComputerDN, InStr(strComputerDN, "DC="))
  arrElements = Split (strComputerDN, ",")

  ' ----- find element where OU= sequence is finished and DC= starts
  For intElementCounter = UBound(arrElements) To 1 Step -1
    If Left(arrElements(intElementCounter), 3) = "DC=" And Left(arrElements(intElementCounter - 1), 3) = "OU=" Then
      intLastOUElement = intElementCounter - 1
      ' ----- Prepare DN to query for group policy objects for the given path, starting at OU level intElementCounter -----
      For intOULevel = 1 to intLastOUElement
        strPath ="LDAP://"
        For intCurrentOU = intOULevel to intLastOUElement
          strPath = strPath + arrElements(intCurrentOU) & ","
        Next
        strPath = strPath + strDomainPartOfDN
        GetListOfGPOs(strPath)
      Next
    End If
  Next
Loop

For i = 0 To intGPOCounter - 1
  If arrGPO(i).blnFound = True Then
    objListOfGPOs.WriteLine arrGPO(i).strDisplayName & ": " & arrGPO(i).strComputerNames
  End If
Next

objListOfGPOs.Close

MsgBox "Done"