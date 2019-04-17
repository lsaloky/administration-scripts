' ----- DHCP Export: reconstruct the table of DHCP leases from DHCP logs, prepare output for import -----

strTargetDHCPServerIP = "<<IPADDRESS>>"			' address of DHCP server where output will be imported

Set objFSO=CreateObject("Scripting.FileSystemObject")

' ----- list of log files to process -----
arrLogfile = Array ("C:\Windows\System32\dhcp\DhcpSrvLog-Mon.log", "C:\Windows\System32\dhcp\DhcpSrvLog-Tue.log", _
		   "C:\Windows\System32\dhcp\DhcpSrvLog-Wed.log", "C:\Windows\System32\dhcp\DhcpSrvLog-Thu.log", _
		   "C:\Windows\System32\dhcp\DhcpSrvLog-Fri.log", "C:\Windows\System32\dhcp\DhcpSrvLog-Sat.log", _
		   "C:\Windows\System32\dhcp\DhcpSrvLog-Sun.log")

' ----- translation between subnets and scopes -----
arrSubnetScopeMapping = Array ("1", "10.0.1.0", "2", "10.0.2.0", "3", "10.0.3.0", "4", "10.0.4.0", _
			       "5", "10.0.5.192", "6", "10.0.6.128", "7", "10.0.7.192")

Set objOutputFile=objFSO.CreateTextFile("DHCPImport.bat")

intMaxSubnet = 0					' number of subnets remembered
Dim arrSubnet (10) 					' third octet of subnet 
Dim arrLease (10,256)					' array of leases for particular subnet

' ----- data type for leases -----
Class clsLease						' declaration of data type for lease
  Public strDate
  Public strTime
  Public strMACAddress
  Public strIPAddress
  Public strDNSName
End Class

For i = 0 To 9						' initialization of array of leases
  For j = 0 To 255
    Set arrLease (i,j) = New clsLease
  Next
Next

' ----- parse all log files -----
For intLogfile = 0 To 6

  Set objInputFile = objFSO.OpenTextFile(arrLogfile(intLogfile), 1)
' ----- ignore initial lines from log file -----
  Do 
    Do 
      strLine = objInputFile.ReadLine
    Loop Until strLine <> "" 				' ignore blank lines at the beginning of file
    arrValues = Split (strLine, ",")
  Loop Until arrValues(0) = "ID"			' until "ID" found			

  Do Until objInputFile.AtEndOfStream			' do until end of file
    Do 
      strLine = objInputFile.ReadLine
    Loop Until strLine <> "" 				' ignore blank lines
    arrValues = Split (strLine, ",")
    
    If arrValues(0) = "10" Or arrValues(0) = "11" Then	' if there is an event Renew or Assign
      arrIPAddress = Split (arrValues(4), ".")		' obtain IP address
      strSubnet = arrIPAddress (2)			' obtain subnet number
      blnFoundInSubnetList = False

      For i = 0 To 9 					' search subnet in the list of subnets
        If arrSubnet(i) = strSubnet Then		' if found
          blnFoundInSubnetList = True

          With arrLease(i,arrIPAddress(3))		' check if date is newer 
            If arrValues(1) > .strDate Then		' (assuming that time is always newer if found)
              .strDate = arrValues (1)			' store values
              .strTime = arrValues (2)
              .strMACAddress = arrValues (6)
              .strIPAddress = arrValues (4)
              .strDNSName = arrValues (5)
            End If
          End With

        End If
      Next

      If blnFoundInSubnetList = False Then		' if not found
        arrSubnet(intMaxSubnet) = strSubnet		' remember new subnet

        With arrLease(intMaxSubnet,arrIPAddress(3))	' store values from lease
          .strDate = arrValues (1)
          .strTime = arrValues (2)
          .strMACAddress = arrValues (6)
          .strIPAddress = arrValues (4)
          .strDNSName = arrValues (5)
        End With
        intMaxSubnet = intMaxSubnet + 1

      End If
    End If
  Loop
Next
' ----- filter out MAC addresses which can appear under 2 or more IPs -----
For i = 0 To 9
  For j = 0 To 254
    For k = 0 to j - 1
      If (arrLease(i,j).strMACAddress = arrLease(i,k).strMACAddress) AND (arrLease(i,j).strDate<>"") AND (arrLease(i,k).strDate<>"") Then
        If arrLease(i,j).strDate < arrLease(i,k).strDate Then	' only date is compared. Assuming that it the date
          arrLease(i,j).strIPAddress = ""			' is equal, time is always newer for "j"
        Else
          arrLease(i,k).strIPAddress = ""
        End If
      End If
    Next
  Next
Next

' ----- write output to file ----- 
objOutputFile.WriteLine "@echo off" & vbCrLf

For i = 0 To 9
  If arrSubnet(i) <> "" Then

    blnFoundInScopeMappingList = False
    For j = 0 To UBound (arrSubnetScopeMapping)-1		' translate subnet to scope
      If arrSubnet(i) = arrSubnetScopeMapping(j) Then
        strScope = arrSubnetScopeMapping (j+1)
        blnFoundInScopeMappingList = True
      End If
    Next

' ----- if scope is not found, all leases from affected subnet will be ignored, other subnets will be processed -----

    For j = 0 To 254					' export directly to batch of netsh commands
      If arrLease (i,j).strIPAddress <> "" Then
        With arrLease(i,j)  
          If .strIPAddress <> "" Then
            objOutputFile.WriteLine "netsh dhcp server " & strTargetDHCPServerIP & " scope " & strScope & " add reservedip " & _
                                    .strIPAddress & " " & .strMACAddress & " " & .strDNSName
          End If 
        End With
      End If
    Next
  End If
Next
