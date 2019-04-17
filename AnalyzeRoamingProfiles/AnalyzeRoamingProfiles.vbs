' ----- important constants -----
strDir = "D:\Profiles"

' if found under special dirs list, ignore subdirectories. Used when user/application is creating per-user customized subdirs.
arrSpecialDirs = Array ("\Application Data\Macromedia\Flash Player", "\Application Data\Mozilla\Firefox", _
  "\Application Data\Sun\Java\Deployment\cache", "\Application Data\ICQLite", "\My Documents", "\Desktop", _
  "\Application Data\Macromedia\Shockwave Player", "\NetHood", "\QIP portable\Users", "\Favorites", _
  "\Application Data\Talkback\MozillaOrg", "\Application Data\Opera", "\Application Data\.purple\logs\jabber", _
  "\Application Data\GTek\GTUpdate\AUpdate\Channels", "\workspace", "\Application Data\Microsoft\MSN Messenger", _
  "\.nx", "\Gaim", "\GTK", "\Application Data\Google\Google Earth", "\Application Data\dvdcss", "\UserData", _
  "\Recent", "\Application Data\AR System\HOME", "\.tmeconsole", "\Application Data\Adobe\Flash Player\AssetCache", _
  "\Application Data\Identities", "\Application Data\Microsoft\Installer", "\Application Data\Real", _
  "\Application Data\Microsoft\Internet Explorer\UserData", "\Application Data\Skype", "\SapWorkDir")

' ----- input / output files -----
Set objFSI = CreateObject("Scripting.FileSystemObject")
Set objInputFile = objFSI.OpenTextFile ("dir.txt",1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutputFile = objFSO.CreateTextFile ("directories.txt")

' ----- initialize data structure for Directory -----
Class clsDirectory
  Public strName
  Public dblFileCount
  Public dblFileSizeTotal
End Class

Dim arrDirectory(5000)

For i = 0 To UBound(arrDirectory) - 1
  Set arrDirectory(i) = New clsDirectory
  arrDirectory(i).dblFileCount = 0
  arrDirectory(i).dblFileSizeTotal = 0
Next

arrDirectory (0).strName = "."
For i = 1 to UBound(arrSpecialDirs) + 1
  arrDirectory (i).strName = arrSpecialDirs(i-1)
Next
  
intCount = UBound(arrSpecialDirs) + 2

s = objInputFile.Readline
s = objInputFile.Readline

' ----- main loop -----
Do Until objInputFile.AtEndOfStream
  s = objInputFile.Readline							' read line from input file
  If s <> "" Then

' ----- directory -----
    If Left(s,10) = " Directory" Then 						' if name of directory found
      If Len(s)-Len(strDir)-15 > 0 Then						' if not root dir
        strCurrentDir = Right (s,Len(s)-15-Len(strDir))				' obtain name of directory

	If InStr(strCurrentDir, "\") > 0 Then 					' if not first-level subdir
	  strCurrentDir = Mid (strCurrentDir, InStr(strCurrentDir, "\"),255)	' obtain current dir except profile name

	  blnFoundInSpecialDirList = false					' search in special directory list
	  For i = 0 To UBound(arrSpecialDirs) 
	    If InStr(strCurrentDir, arrSpecialDirs(i)) = 1 Then			' if found 
	      intCurrentDir = i							' remember index
	      BlnFoundInSpecialDirList = True
	    End If	                	    
	  Next

	  blnFoundInDirList = false						' search in current directory list
	  For i = 0 To intCount
	    If strCurrentDir = arrDirectory(i).strName Then 			' if found 
	      intCurrentDir = i							' remember index
	      BlnFoundInDirList = True
	    End If	                	    
	  Next

	  If (blnFoundInDirList = False) and (blnFoundInSpecialDirList = False) Then	' if not found
	    arrDirectory(intCount).strName = strCurrentDir			' create new record for dir
	    intCurrentDir = intCount						' set index
	    intCount = intCount + 1						' increase count of dirs
	  End If  	  

        Else									' if first-level subdir
	  strCurrentDir = "."							' set current dir to "."
	  intCurrentDir = 0
        End If

      Else									' if root dir
        strCurrentDir = "" 							' set to "" - will be ignored later
      End If

    Else

' ----- file -----
      If Left(s,1) <>" " Then							' if not directory - row with file information or empty row
	strFileName = Trim(Mid(s,40,Len(s)-39))					' obtain file name
	dblFileSize = Trim(Mid(s,25,14))					' obtain file size

	If (strFileName <> ".") and (strFileName<> "..") and (Mid(s,25,5) <> "<DIR>") and (strCurrentDir <> "") Then
	  With arrDirectory(intCurrentDir)
	    .dblFileCount = .dblFileCount + 1					' increase file count
	    .dblFileSizeTotal = .dblFileSizeTotal + dblFileSize			' increase total file size 
	  End With
	End If

      End If
    End If
  End If
Loop

' ----- write to output file -----
For i = 0 to intCount - 1
  objOutputFile.WriteLine arrDirectory(i).strName & ", " & arrDirectory(i).dblFileCount & ", " & arrDirectory(i).dblFileSizeTotal
Next

objInputFile.Close
objOutputFile.Close
