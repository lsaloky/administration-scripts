Dim arrGPOName(5000)     ' array of strings, contains GPO names
intGPOCounter = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objInputFile = objFSO.OpenTextFile ("Data.txt", 1)
Set objOutputFile = objFSO.CreateTextFile("GPOList.txt")
Set objErrorFile = objFSO.CreateTextFile("Errors.txt")

blnComputerSettings = false

' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  strLine = objInputFile.Readline

  ' ----- check if we are inside computer settings hive -----
  If InStr(strLine, "COMPUTER SETTINGS") > 0 Then
    blnComputerSettings = true
  End If

  ' ----- check if we are inside user settings hive -----
  If InStr(strLine, "USER SETTINGS") > 0 Then
    blnComputerSettings = false
  End If

  ' ----- report server names, where the previous script with GPRESULT was unable to connect -----
  If InStr(strLine, "ERROR: The RPC server is unavailable.") > 0 Then
    strLine = objInputFile.Readline
    objErrorFile.WriteLine strLine
  End If
  
  ' ----- remember all GPOs, if not already remebered -----
  If (InStr(strLine, "Applied Group Policy Objects") > 0) And (blnComputerSettings = true) Then
    strLine = objInputFile.Readline

    Do Until strLine = ""

      strLine = objInputFile.Readline

      If strLine <> "" Then

        ' ----- check if GPO is already on the list -----
        blnFoundInGPOList = false 
        For i = 0 To intGPOCounter -1 
          If arrGPOName(i) = strLine Then 
            blnFoundInGPOList = true
          End If
        Next

        ' ----- add to GPO list, if not found -----
        If blnFoundInGPOList = false Then
          arrGPOName(intGPOCounter) = strLine 
          intGPOCounter = intGPOCounter + 1
        End If
      End If
    Loop
  End If
Loop

' ----- write GPO names -----
For i = 0 To intGPOCounter -1
  objOutputFile.WriteLine LTrim(arrGPOName(i))
Next

MsgBox "Done"
