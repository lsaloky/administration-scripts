' ----- constants -----

' thisScriptIsRunAs: enter here user name you use for collecting events, network logins produced by this user name will be 
' bypassed becouse script is producing many Kerberos events
thisScriptIsRunAs = "company_administrator_account"

' busimessHoursBegin, businessHoursEnd: interval of business hours. Logins outside this interval will be considered as
' suspicious for users from Users_BusinessHoursLogon.txt file
' remember to set time format to 24 hour - hh:mm:ss
businessHoursBegin = 7
businessHoursEnd = 17

' list of ignored computers. Network logons to this computers will be ignored. Array must have 10 elements, can leave blank
listOfIgnoredComputers = Array ("COMPUTERNAME1", " ", " ")

' list of logon events - used by one of the filters
' Event ID 538 is not listed intentionally - logoff is not suspicious
' Event ID 529 is not listed intentionally - failure logon is processed separately and other filters are not applied
listOfEvents = Array ("528", "530", "531", "532", "533", "534", "535", "536", "537", "539", "540", "682", "683")

' failure logons sensitivity. If found more than x failure logons for each computer&user name, report as suspicious.
' if user name is not in ATL or BHL list, report as suspicious even when less failure logons
failureLogonSensitivity = 4






' ----- function for writing to output file: for ID 529 there isn't Session ID, columns are shifted -----
function WriteToOutputFile
  If s(4) = "529" Then
    objOutputFile.WriteLine (s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(11) & "," & s(12))
  Else
    If (s(4) = "682") or (s(4) = "683") Then
      objOutputFile.WriteLine (s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & ",N/A," & s(13))
    Else
      objOutputFile.WriteLine (s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(12) & "," & s(13))
    End If
 End If
End Function

' ----- function for checking if username is listed in Any Time Logon list -----
function FoundInATLList (z)
  FoundInATLList = false
  For i = 0 To numUsersATL-1 
    If z = userATL(i) Then FoundInATLList = true
  Next
End Function

' ----- function for checking if username is listed in Business Hour Logon list -----
function FoundInBHLList (z)
  FoundInBHLList = false
  For i = 0 To numUsersBHL-1 
    If z = userBHL(i) Then FoundInBHLList = true
  Next
End Function






' ----- input file -----
Set objFSI=CreateObject("Scripting.FileSystemObject")
Set objInputFile=objFSI.OpenTextFile("Security.txt", 1)

' ----- data files: Users_BusinessHoursLogon.txt, Users_AnyTimeLogon.txt -----
Set objFSUser1=CreateObject("Scripting.FileSystemObject")
Set objUser1=objFSUser1.OpenTextFile("Users_BusinessHoursLogon.txt", 1)
Set objFSUser2=CreateObject("Scripting.FileSystemObject")
Set objUser2=objFSUser2.OpenTextFile("Users_AnyTimeLogon.txt", 1)

' ----- read all user names from Users_* files, BHL is BusinessHoursLogin, ATL is AnyTimeLogin -----
currentuser = 0
Dim userBHL(1000)
Do Until objUser1.AtEndOfStream
  userBHL(currentuser) = Trim(objUser1.Readline)
  currentuser = currentuser + 1
Loop
numUsersBHL = currentuser

currentuser = 0
Dim userATL(1000)
Do Until objUser2.AtEndOfStream
  userATL(currentuser) = Trim(objUser2.Readline)
  currentuser = currentuser + 1
Loop
numUsersATL = currentuser

' ----- output file -----
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objOutputFile=objFSO.CreateTextFile("Daily_Report_Security.txt")

' ----- array of user + computer names, number of failure logons -----
Dim userName(1000)
Dim computerName(1000)
Dim failureLogonsUC(1000)
maxName=0






' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  line = objInputFile.Readline
  s = Split(line , chr(9))
  If UBound (s)>8 Then 
    s(9) = LCase(s(9))

' ----- failure logon processing -----
    If s(4) = "529" Then 
      foundInUCList = false
      For i = 0 to maxName
        If (s(9) = userName(i)) and (s(8) = computerName (i)) Then
          failureLogonsUC(i) = failureLogonsUC(i) + 1
          foundInUCList = true
         End If 
      Next
      If not foundInUCList Then
        userName(maxName) = s(9)
        computerName (maxName) = s(8)
        failureLogonsUC(maxName) = 1
        maxName = maxName + 1
      End If
    End If

' ----- policy change processing ----- 
    If s(4)="608" or s(4)="609" Then
      objOutputFile.WriteLine ("Policy Change: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10))
    End If

' ----- account management processing -----
    If s(4)>=624 and s(4)<=668 Then
    Select Case s(4)
      Case "624" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10) & "," & s(12))
      Case "627" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10))
      Case "628" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10) & "," & s(12))
      Case "630" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10) & "," & s(12))
      Case "628" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(9) & "," & s(10) & "," & s(12))
      Case "633" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(10) & "," & s(11) & "," & s(14) & "," & s(15))
      Case "636" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(10) & "," & s(11) & "," & s(14) & "," & s(15))
      Case "637" objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8) & "," & s(10) & "," & s(11) & "," & s(14) & "," & s(15))
      Case "642" ' do nothing - avoid duplicity
      Case Else objOutputFile.WriteLine ("Account Management: " & s(0) & "," & s(1) & "," & s(4) & "," & s(8))
      End Select
    End If

' ----- bypass events not logged by auditing logon events -----
    foundInEventList = false
    For i = 0 To 12
      If s(4) = listOfEvents(i) Then foundInEventList = true
    Next
    If foundInEventList Then

' ----- bypass events logged by system -----
      If (s(6) <> "NT AUTHORITY\SYSTEM") and (s(6) <> "NT AUTHORITY\ANONYMOUS LOGON") Then 
  
' ----- bypass events with computer account as username -----
        If Right (s(9), 1) <> "$" Then

' ----- bypass network logins made by this script -----
          If (s(9) <> thisScriptIsRunAs) or (s(12) <> "3") Then

' ------ bypass network logons to listed computers -----
            foundInComputersList = false
            For i = 0 To 9 
              If s(8) = listOfIgnoredComputers(i) Then foundInComputersList = true
            Next
            If not foundInComputersList Then

' ----- bypass events logged by AdvApi - produced by Windows XP -----
              If s(13) <> "Advapi  " Then
 
' ----- bypass user logon when found in ATL list -----
                If not FoundInATLList(s(9)) Then

' ----- bypass user logon if found in business hours logon list, time is inside business hours, no weekend -----
' ----- WARNING: result of Split command may depend on Regional options - time settings, s(1) should be in HH:mm format (without AM/PM) -----
                  t = Split(s(1),":")

' ----- WARNING: result of DatePart command may depend on Regional options - date settings, s(0) should be in mm/dd/yyyy format -----
                  dayofweek = DatePart("w", s(0))
                  If (not FoundInBHLList(s(9))) or (dayofweek=1) or (dayofweek=7) or (t(0)*1<businessHoursBegin) or (t(0)*1>businessHoursEnd) Then

' ----- bypass logons to the user's computer - first four characters of username is part of computer name -----
                    u = Left(s(9),4)
                    If Instr(1,s(8),u,vbTextCompare) = 0 then

' ----- output to file -----
                      WriteToOutputFile()
                    End If 
                  End If
                End If
              End If
            End If 
          End If 
        End If 
      End If
    End If
  End If 
Loop

' ----- write failure logon counts to output file -----
objOutputFile.WriteLine ("---------------------------------------------------------------")
objOutputFile.WriteLine ("Number of failure logons - higher than " & failureLogonSensitivity)
For i = 0 to maxName - 1 
  If (failureLogonsUC(i) > failureLogonSensitivity) Then objOutputFile.WriteLine (userName(i) & "," & computerName(i) & ": " & failureLogonsUC(i))
Next
objOutputFile.WriteLine ("---------------------------------------------------------------")
objOutputFile.WriteLine ("Suspicious usernames failure logon counts:")
For j = 0 to maxName - 1 
  If not FoundInBHLList(userName(j)) Then
    If not FoundInATLList(userName(j)) Then
      If FailureLogonsUC(j) <= FailureLogonSensitivity Then
        objOutputFile.WriteLine (userName(j) & "," & computerName(j) & ": " & failureLogonsUC(j))
      End If
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

' ----- list of event IDs for logon events -----
' 528	A user successfully logged on to a computer.
' 529	The logon attempt was made with an unknown user name or a known user name with a bad password.
' 530	The user account tried to log on outside the allowed time.
' 531	A logon attempt was made by using a disabled account.
' 532	A logon attempt was made by using an expired account.
' 533	The user is not allowed to log on at this computer.
' 534	The user attempted to log on with a logon type that is not allowed, such as network, interactive, batch, service, or remote interactive.
' 535	The password for the specified account has expired.
' 536	The Netlogon service is not active.
' 537	The logon attempt failed for other reasons.
' 538	A user logged off. 
' 539	The account was locked out at the time the logon attempt was made. This event is logged when a user or computer attempts to authenticate with an account that has been previously locked out.
' 540	Network logon succeeded.
' 682	A user has reconnected to a disconnected Terminal Services session.
' 683	A user disconnected a Terminal Services session without logging off.

' ----- list of event types -----
' 2	Console logon - interactive
' 3	Network logon - network mapping
' 4	Batch logon - scheduler
' 5	Service logon - service uses an account
' 7	Unlock workstation

' ----- list of logon processes -----
' NtLmSsp - NTLM authentication protocol Security Support Provider (SSP) initiated the logon
' IIS - Microsoft IIS initiated the logon (this situation occurs when you use anonymous access or basic authentication on the IIS level)
' DCOMSCM - ???
' Advapi - application called the LogonUser function to initiate a logon process
' User32 - console logon
' seclogon - secondary logon = runas
' Kerberos - network logon

' ----- list of logon IDs for policy change events
' 608	User right assigned
' 609	User right removed
' 612	Audit Policy change - not in the report, is logged after each restart

' ----- list of logon IDs for account management events
' 624	User account created
' 625	User account type changed
' 626	User account enabled
' 627	NT AUTHORITY\ANONYMOUS is trying to change password - occurs when password expires and user tries to change it
' 628	User account password set
' 629	User account disabled
' 630	User account deleted
' 631	Security enabled Global Group created
' 632	Security enabled Global Group member added
' 633	Security enabled Global Group member removed
' 634	Security enabled Global Group deleted
' 635	Security enabled Local Group created
' 636	Security enabled Local Group member added
' 637	Security enabled Local Group member removed
' 638	Security enabled Local Group deleted
' 639	Security enabled Local Group changed
' 641	Security enabled Global Group changed
' 642	User account changed
' 644	User account locked
' 648 - 667	Security disabled group events
' 668	Group type changed

