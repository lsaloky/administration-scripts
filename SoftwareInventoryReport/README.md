# SoftwareInventoryReport

Script creates a report with list of all software.

1. Modify Inventory_runas.bat to run Inventory.bat file under your account with domain admin rights
2. Update exclusions in "Freeware list.txt". Software listed here will not appear in the report.
3. If you want to exclude servers, update the command 
```Left(strComputer,12) <> "SERVERPREFIX")``` 
in .vbs file
4. Run Inventory_runas.bat, enter password
5. You can find results in tab-delimited file Software.txt. You can import it into Excel. There is no progress bar, wait until command prompt window closes automatically

