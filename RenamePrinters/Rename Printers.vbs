' ----- list of old printer names -----
arrOldNames = Array ("OLDNAME1", "OLDNAME2") 
' ----- list of new printer names -----
arrNewNames = Array ("NEWNAME1", "NEWNAME2") 

Set objNetwork = WScript.CreateObject("WScript.Network")

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colInstalledPrinters =  objWMIService.ExecQuery ("Select * from Win32_Printer")

For Each objPrinter in colInstalledPrinters
  For i = 0 To UBound (arrOldNames) 

' ----- check if there is printer \\servername\printername, if yes then remove it and add new one -----
    If UCase(objPrinter.Name) = "\\OLDSERVERNAME\" & arrOldNames (i) Then
      objNetwork.RemovePrinterConnection "\\OLDSERVERNAME\" & arroldNames (i)
      objNetwork.AddWindowsPrinterConnection "\\NEWSERVERNAME\" & arrNewNames (i)
    End If

' ----- check if there is printer \\fullservername\printername, if yes then remove it and add new one -----
    If UCase(objPrinter.Name) = "\\oldservername.domain.com\" & arrOldNames (i) Then
      objNetwork.RemovePrinterConnection "\\oldservername.domain.com\" & arroldNames (i)
      objNetwork.AddWindowsPrinterConnection "\\NEWSERVERNAME\" & arrNewNames (i)
    End If

  Next
Next

