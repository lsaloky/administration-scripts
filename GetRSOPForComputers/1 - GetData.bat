@ECHO OFF
DEL Data.txt
DEL Errors.txt
FOR /F %%I IN (Computers.txt) DO (
  GPRESULT /S %%I >>Data.txt 2>&1
  ECHO %%I >>Data.txt
)