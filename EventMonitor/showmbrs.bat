@echo off
del Temp_*.txt
for /F %%i in (Computers.txt) do @showmbrs \\%%i\Administrators >>Temp_Administrators.txt
for /F %%i in (Computers.txt) do @showmbrs \\%%i\Power Users >>Temp_PowerUsers.txt

pause
