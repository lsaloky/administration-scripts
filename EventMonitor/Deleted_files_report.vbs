' ----- input file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Deleted_files_report.txt", 1)

' ----- output file -----
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("Deleted_files_filtered_report.txt")

currentDeletedFile = 0
dim deletedFileList (10000)
' ----- main loop -----
Do Until objInputFile.AtEndOfStream

  line = objInputFile.Readline
  s = Split(line , chr(9))
  s(9) = LCase(s(9))


' ----- Event ID 560, Success audit, not a .TMP file -----    
  If (s(4) = "560") and (s(2) = "8") and (UCase(Right(s(11),3)) <> "TMP") Then

' ----- do not audit D:\home\Apps\System -----
    If UCase(Left(s(11),19)) <> "D:\HOME\APPS\SYSTEM" Then

' ----- attributes = deleted: save to list, remove from list if it is already there -----
      If s(22) = "%%1537  " Then
        foundInDeletedFileList = false
        For i = 0 To currentDeletedFile - 1
          If s(11) = deletedFileList(i)(11) Then 
            foundInDeletedFileList = true
            indexInDeletedFileList = i
          End If 
        Next
        If foundInDeletedFileList Then 
          deletedFileList(indexInDeletedFileList)(0) = "_" 
        Else
          deletedFileList(currentDeletedFile) = s
          currentDeletedFile = currentDeletedFile + 1          
        End If
      End If
    End If
  End If
Loop

' ----- write output to the output file -----
For i = 0 To currentDeletedFile-1
  If deletedFileList(i)(0) <> "_" Then 

' ----- remove "_:\___\___\~$____.doc" - temporary Word files -----
    If (InStr(1, deletedFileList(i)(11), "~$", vbTextCompare) = 0) or (UCase(Right(deletedFileList(i)(11),3)) <> "DOC") Then
      objOutputFile.WriteLine (deletedFileList(i)(0) & "," & deletedFileList(i)(1) & "," & deletedFileList(i)(6) & "," & deletedFileList(i)(11))
    End If
  End If
Next

' ----- Structure of s string -----
' s(0) - date
' s(1) - time
' s(2) - Success/Failure audit - "8" or "16"
' s(3) - Category: always "2"
' s(4) - event ID
' s(5) - source: always "Security"
' s(6) - user name (or source computer name, if this is computer logon)
' s(7) - blank
' s(8) - destination computer name
' s(9) - user name (or source computer name, if this is computer logon)
' s(10) - domain: MOLEX, or local computer
' s(11) - logon ID (if present), for failure logons s(12) is here - all columns from now are shifted
' s(12) - event type - from "2" to "7"
' s(13) - logon process - User32, Kerberos, or no data
' s(14) ... - sometimes additional data
