@echo off

rem ----- Output directory should exist - see outputdir variable in .vbs file -----
echo Step 1/3: Creating list of files ...
dir /s /q /-c >dir.txt

rem ----- Process output from dir command -----
echo Step 2/3: Processing list of files ...
wscript DirOwner2.vbs

rem ----- Delete listings for Administrators group and unknown users -----
del DirOwner\Administrators.txt
del dirOwner\S-1-5-*.txt

rem ----- Create HTML files with sorted output -----
echo Step 3/3: Creating .HTML files, sorting ...
wscript CreateHTMLFiles.vbs
