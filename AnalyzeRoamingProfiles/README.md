# AnalyzeRoamingProfiles

Analyze directory sizes in directory with roaming profiles. Create a report of directory total size across all roaming profiles.

1. Copy .bat and .vbs file into the root folder with roaming profiles
2. Update path in AnalyzeRoamingProfiles.vbs, line 2: strDir = "D:\Profiles" to match your roaming profiles root folder
3. Add / remove directories to exclude - see arrSpecialDirs in .vbs file
4. Execute "AnalyzeRoamingProfiles.bat"
5. Review results stored in output file "directories.txt"