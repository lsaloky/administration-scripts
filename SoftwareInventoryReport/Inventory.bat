@echo off

del Temp*.txt

rem ----- get list of all computers from AD -----
csvde -f Temp_Computers.txt -r "(objectClass=computer)" -l "DN"

rem ----- remove all unnecessary characters from csvde output, output file: Computers.txt -----
for /F "tokens=1,2 delims==" %%i in (Temp_Computers.txt) do @echo %%j >>Temp_RemovedChars1.txt
for /F "tokens=1,2 delims=," %%i in (Temp_RemovedChars1.txt) do @if not %%j_==_ @echo %%i >>Computers.txt

echo This process could take several hours, please wait. This window will be closed automatically.

wscript Inventory.vbs
del Temp*.txt 
del Computers.txt
DEL *.log
