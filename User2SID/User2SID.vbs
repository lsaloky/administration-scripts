Option Explicit
Dim objComputer, strSid, arrbytSid

' enter Hexadecimal value from ADSI edit 
Wscript.Echo "Decimal SID: " & HexStrToDecStr("010500000000000515000000D10A949F4483166BDF2882206F100000")

Function HexStrToDecStr(strSid)
Dim arrbytSid, lngTemp, j

ReDim arrbytSid(Len(strSid)/2 - 1)
For j = 0 To UBound(arrbytSid)
  arrbytSid(j) = CInt("&H" & Mid(strSid, 2*j + 1, 2))
Next

HexStrToDecStr = "S-" & arrbytSid(0) & "-" & arrbytSid(1) & "-" & arrbytSid(8)

lngTemp = arrbytSid(15)
lngTemp = lngTemp * 256 + arrbytSid(14)
lngTemp = lngTemp * 256 + arrbytSid(13)
lngTemp = lngTemp * 256 + arrbytSid(12)

HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

lngTemp = arrbytSid(19)
lngTemp = lngTemp * 256 + arrbytSid(18)
lngTemp = lngTemp * 256 + arrbytSid(17)
lngTemp = lngTemp * 256 + arrbytSid(16)

HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

lngTemp = arrbytSid(23)
lngTemp = lngTemp * 256 + arrbytSid(22)
lngTemp = lngTemp * 256 + arrbytSid(21)
lngTemp = lngTemp * 256 + arrbytSid(20)

HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

lngTemp = arrbytSid(25)
lngTemp = lngTemp * 256 + arrbytSid(24)

HexStrToDecStr = HexStrToDecStr & "-" & CStr(lngTemp)

End Function

