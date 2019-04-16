# GetOUForComputers

Get all OUs for the computers listed in input file.

1. Enter computer names into ComputersToSearchFor.txt
2. Execute script to export all computers: "1 - ExportComputers.bat". Output file "AllComputers.txt" will be created
3. Execute script to search for computers: "2 - SearchForComputers.vbs"
4. Output will be stored in files OUs.txt, with the syntax "OUName<tab>comma-delimited list of computer names located in this OU"
5. NotFound.txt output file will contain computer names, which have not been found
