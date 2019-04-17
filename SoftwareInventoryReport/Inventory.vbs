On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile("Software.txt", True)
Set objFSI = CreateObject("Scripting.FileSystemObject")
Set objInputFile = objFSI.OpenTextFile("Computers.txt")
Set objFSFreeware = CreateObject("Scripting.FileSystemObject")
Set objFreewareFile = objFSFreeware.OpenTextFile("Freeware list.txt")

Dim freewareList (1000)

' ----- read list of freeware -----
lastRecord = 1
Do Until objFreewareFile.AtEndOfStream
  s = Trim(objFreewareFile.ReadLine)
  If (s <> "") and (Left (s,1) <> ";") Then freewareList (lastRecord) = s
  lastRecord = lastRecord+1
Loop

objOutputFile.WriteLine ("Computer Name" & vbtab & "Software Name" & vbtab & "Software Vendor" & vbtab & "Software Version")

' ----- Main loop -----
Do Until objInputFile.AtEndOfStream
  strComputer = Trim(objInputFile.ReadLine)

' ----- ignote blank lines and servers -----
  If (strComputer <> "") and (Left(strComputer,12) <> "SERVERPREFIX") Then
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colSoftware = objWMIService.ExecQuery ("Select * from Win32_Product")
    For Each objSoftware in colSoftware

' ----- search for software in freeware list -----
      foundInFreewareList = false
      For i = 1 To lastRecord - 1 
        If objSoftware.Name = freewareList (i) Then foundInFreewareList = true
      Next
      If Not foundInFreewareList Then
        objOutputFile.WriteLine (strComputer & vbtab & objSoftware.Name & vbtab & objSoftware.Vendor & vbtab & objSoftware.Version)
      End If
    Next
    Set colSoftware = nothing
    Set objWMIService = nothing
  End If
Loop

objOutputFile.Close
objInputFile.Close
