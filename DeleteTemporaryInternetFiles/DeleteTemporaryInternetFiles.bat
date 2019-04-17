@echo off
for /D %%I in (DOCUME~1\*.*) do echo %%I
for /D %%I in (DOCUME~1\*.*) do del /F /S /Q C:\%%I\LOCALS~1\TEMPOR~1\*.* >>C:\a.txt
pause
