# Shrink rtf files in folder structure

This simple python script walks your file system 
 - finds rtf files (larger than x - set the size in the source code)
 - saves a list of the files in a sqlite database
 - runs through the list of rtfs in the db and opens them in MS Word and saves them

 Unfortunately pywin32 is really picky about file and foldernames. So this script copies the found file to a tmp file within the scripts location und after shrinking it copies it back.
 So make sure the script runs from a "simple path" with no white spaces or non-ascii characters. Use at your own risk.

 ## Usage

````
Usage: main.py [OPTIONS]

Options:
  --top-dir TEXT  Defines the top of the tree..
  --scan          Scans the directory recursively to find RTF files matching
                  set criteria
  --shrink        Finds rtfs from db and reduces their file size
  --dry-run       don't actually shrink them
  --help          Show this message and exit.
````


 ## MS Word ExportPictureWithMetafile

 Make sure that your MS Word instance is configured to "write small RTF files".
 This is accomplished with the following registry key (Office 2010 was used, use 15.0 and 16.0 for newer versions):

 ````
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Word\Options]
"ExportPictureWithMetafile"="0"

 ````

 As this script uses pywin32 and MS Word it is **windows only**
 No warranties, as this was written quickly to solve a urgent problem. Use with care.