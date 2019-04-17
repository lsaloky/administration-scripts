Set objFS = CreateObject("Scripting.FileSystemObject")
set objInputFile = objFS.OpenTextFile("Temp.txt", 1)
set objOutputFile = objFS.CreateTextFile("PhoneList.html")
Dim riadok (1000)

objOutputFile.WriteLine("<HTML>" & vbCrLf & "<HEAD>" & vbCrLf & "  <TITLE>Phone List</TITLE>" & vbCrLf _
  & "</HEAD>" & vbCrLf & "<BODY>" & vbCrLf & "  <FONT FACE='Courier'>" & vbCrLf & "    <TABLE>" & vbCrLf _
  & "      <TR><TD><B>Meno</B>&nbsp;&nbsp;</TD><TD><B>Miestnost</B></TD><TD><B>Oddelenie</B></TD>" _
  & "<TD><B>Budova&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</B></TD><TD><B>Telef&oacute;n</B></TD></TR>" & vbCrLf)

' ----- read and process source data -----
count = 0
Do Until objInputFile.AtEndOfStream 
  s = objInputFile.ReadLine
  If s <> "" Then
    t = Split (s,":") 							' split input line
    if t(0) = "displayName" Then username = t(1)
    if t(0) = "physicalDeliveryOfficeName" Then useroffice = t(1)
    if t(0) = "telephoneNumber" Then userphone = t(1)
    if t(0) = "department" Then userdepartment = t(1)
    if t(0) = "streetAddress" Then useraddress = t(1)
  Else									' blank line, process loaded data
    If userphone <> "" Then
      riadok (count) = username & "</TD><TD>" & useroffice & "</TD><TD>" & userdepartment _
        & "</TD><TD>" & useraddress  & "</TD><TD>" & userphone		' store data into riadok variable
      count = count + 1							' number of records
    End If
    username = ""
    useroffice = ""
    userphone = ""
    userdepartment = ""
    useraddress = ""
  End If
Loop

' ----- sort data -----
For i = 0 To count-1 
  For j = 0 to i-1 
    If riadok(i)<riadok(j) Then
      riadokTemp = riadok(i)
      riadok(i) = riadok(j)
      riadok(j) = riadokTemp
    End If
  Next
Next

' ----- check for duplicate records -----
For i = 0 to count - 2
  If Left(riadok(i),InStr(riadok(i), "</TD>")) = Left(riadok(i+1),InStr(riadok(i+1), "</TD>")) Then 
    riadok(i) = ""
  End If
Next
' ----- write data -----
For i = 0 To count - 1
  If riadok(i) <>"" Then
    objOutputFile.WriteLine("      <TR><TD>" & riadok(i) & "</TD></TR>" & vbCrLf )
  End If
Next

objoutputFile.WriteLine("    </TABLE>" & vbCrLf & "  </FONT>" & vbCrLf & "</BODY>" & vbCrLf & "</HTML>")
