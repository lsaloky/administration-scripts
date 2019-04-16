# EventMonitor

Export all events for the last day from all computers and extract errors. Analyze security logs - get all users working 8 x 5 and 24 x 7 and search for user accounts working 8 x 5 logging on out of business hours. Report also unsuccessful logins.

Note that tools csvde, ldifde, dumpel and showmbrs are needed in order to execute this script.

1. Execute "Dump_Events_Runas.bat"
2. To get daily report of errors, execute "FilterErrorLogs.vbs". Output file: "Daily_Report_Errors.txt"
3. To analyze security logs, execute "ProcessSecurityLogs.vbs". Output file: "Daily_Report_Security.txt"
4. Review output file "Daily_Report_USB_Usage.txt" to see USB sticks usage.

# Deleted_files_report

Get list of all files deleted on file server for the last three days.

1. Execute "Deleted_files_report.bat"
2. Execute "Deleted_files_report.vbs"
3. Review output file "Deleted_files_filtered_report.txt"

# showmbrs report

Get list of users in local Administrators and Power Users groups from all computers.

1. Execute "showmbrs.bat"
2. Review output files "Temp_Administrators.txt" and "Temp_PowerUsers.txt"