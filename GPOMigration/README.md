# GPOMigration

Migrate all groups and GPOs to the new domain

1. Execute initial step "0 - BackupAllGPOs.bat". Group Policy Management Console must be installed in C:\Program Files\GPMC\Scripts. Output will be stored in E:\Backup
2. Execute script to get GPOs in domain: "1 - Analyze.vbs". This script will create a batch file to export all policies: "2 - Export.bat" and to import policies: "5 - ImportGPOs.bat". Script will also create lists of GPOs: "GPOList.txt" and "RelevantGPOList.txt"
3. Execute "2 - Export.bat". This file will create output files "Report<<number>>.csv"
4. Execute "3 - ModifyGroupNames.vbs" to update group names
5. Execute "4 - ImportNewGroups.bat" to import groups into new domain
6. Execute "5 - ImportGPOs.bat" to import GPOs