Const HKCU = &H80000001

' ----- connect to the registry on local computer -----
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

' ----- obtain user name from registry: HKCU, Registry Path, Key Name, Value -----
objReg.GetStringValue HKCU, "Software\Microsoft\Windows\CurrentVersion\Explorer", "Logon User Name", strValue

' ----- convert string to array of string values, each character to 2 strings: "A" to "65","0"; "B" to "66","0" ... ----- 
Dim arrValue ()
ReDim arrValue(2 * Len (strValue) + 1)
For i = 0 To Len (strValue) - 1
  arrValue (2 * i) = CStr( Asc (Mid (strValue,i+1,1)))
  arrValue (2 * i + 1) = "0"
Next
arrValue (2 * Len (strValue)) = "0" ' 2 x zero at the end of string
arrValue (2 * Len (strValue) + 1) = "0"

' ----- Set registry value: HKCU, Registry Path, Key Name, Value -----
objReg.SetBinaryValue HKCU,"Software\Microsoft\Office\11.0\Common\UserInfo","UserName",arrValue
 
	
