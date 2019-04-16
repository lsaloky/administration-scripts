@ECHO OFF 
FOR %%I IN (NewReport*.csv.ldf) DO ldifde -i -f %%I
pause