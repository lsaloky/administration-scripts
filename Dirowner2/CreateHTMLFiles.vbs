inputPath = "D:\home\DirOwner"
outputPath = "D:\home\DirOwner"

' ----- paths to be replaced -----
absolutePath = Array ("D:\home\Dir1, "D:\home\Dir2", "D:\home\Dir3")
relativePath = Array ("R:", "F:", "H:")

Set objFS = CreateObject("Scripting.FileSystemObject")
Dim userName(1000)   ' user names
Dim fileName(100000) ' file names of current user, max 100000
Dim fileSize(100000) ' file sizes of current user

' ----- get list of files in folder, remove extension -----
Set objFolder=objFS.GetFolder(inputPath) 
maxFiles = 0
For Each objFile in objFolder.Files 
  If Right(objFile.Name,3) = "txt" Then
    userName(maxFiles) = Left(objFile.Name,Len(objFile.Name)-4)
    maxFiles = maxfiles+1
  End If
Next

' ----- main loop -----
For mainLoop = 0 to maxFiles-1 
  set objInputFile = objFS.OpenTextFile(inputPath & "\" & userName(mainLoop) & ".txt", 1)
  set objOutputFile = objFS.CreateTextFile(outputPath & "\" & userName(mainLoop) & ".html")
  objOutputFile.WriteLine("<HTML>" & vbCrLf _
    & "<HEAD>" & vbCrLf _
    & "  <TITLE>" & userName(mainLoop) & "</TITLE>" & vbCrLf _
    & "</HEAD>" & vbCrLf _
    & "<BODY>" & vbCrLf _
    & "  <FONT FACE='Courier'>" & vbCrLf _
    & "    <TABLE>" & vbCrLf _
    & "      <TR><TD><B>Veækosù</B>&nbsp;&nbsp;</TD><TD><B>Meno s˙boru</B></TD></TR>")

' ---- inside file loop: read file names -----
  insideLoop = 0
  Do Until objInputFile.AtEndOfStream 
    s = objInputFile.ReadLine
    fileSize(insideLoop) = Trim(Left(s,InStr(s,"*")-2)) * 1
    fileName(insideLoop) = Trim(Right(s,Len(s)-InStr(s,"*")))

' ----- replace path in file name to relative path (network drive) -----
    For i = 0 to UBound(absolutePath)
      t = fileName(insideLoop)
      If Left(t,Len(absolutePath(i))) = absolutePath(i) Then
        fileName(insideLoop) = relativePath(i) & Right(t,Len(t)-Len(absolutePath(i)))
      End If

' ----- if username is in path, assume it is Q: drive ----- 
      If InStr(t,"\" & userName(MainLoop) & "\")>0 Then
        filename(insideLoop) = "Q:" & Right(t,Len(t)-InStr(t,"\" & userName(MainLoop) & "\") - Len(userName(mainLoop)))
      End If 
    Next
    insideLoop = insideLoop + 1
  Loop

' ----- sort array of file names, biggest files first -----
  For i = 0 To insideLoop-1 
    For j = 0 to i-1 
      If fileSize(i)>fileSize(j) Then
        fileSizeTemp = fileSize(i)
        fileSize(i) = fileSize(j)
        fileSize(j) = fileSizeTemp
        fileNameTemp = fileName(i)
        fileName(i) = fileName(j)
        fileName(j) = fileNameTemp
      End If
    Next
  Next
  For i = 0 to insideLoop-1 
    objOutputFile.WriteLine("      <TR><TD>" & fileSize(i) & "</TD><TD>" & fileName(i) & "</TD></TR>")
  Next
  objoutputFile.WriteLine("    </TABLE>" & vbCrLf _
    & "  </FONT>" & vbCrLf _
    & "</BODY>" & vbCrLf _
    & "</HTML>")
Next
