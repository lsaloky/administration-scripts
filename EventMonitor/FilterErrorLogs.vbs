listOfIgnoredErrors = Array ( _
 "A socket operation was attempted to an unreachable host.", _
 "The specified domain either does not exist or could not be contacted.", _
 "Došlo k pokusu o operáciu so soketom v èase nedosiahnute¾nosti hostite¾a.", _
 "Windows cannot determine the user or computer name. Return value (1326).", _
 "The timeout waiting for the performance data collection function", _
 "The time provider NtpClient is configured to acquire time from one or more  time sources, however none of the sources are currently accessible.", _
 "No Domain Controller is available for domain MOLEX due to the following", _
 "The automatic certificate enrollment subsystem could not access local resources needed for enrollment.", _
 "TNT The RPC server is unavailable.", _ 
 "1 0 1030 Userenv", _
 "Automatic certificate enrollment for local system failed to contact the active directory (0x8007054b).")

' ----- input / output file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Daily_Report_Errors_not_filtered.txt", 1)
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("Daily_Report_Errors.txt")

' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  line = objInputFile.Readline
  foundInList = false
  For i = 0 to UBound(listOfIgnoredErrors)
    If InStr(line,listOfIgnoredErrors(i)) > 0 Then FoundInList = true
  Next
  If foundInList = false Then objOutputFile.WriteLine (line)
Loop

