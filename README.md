# SongInventory
_Creates an Excel spreadsheet inventory of songs for Yarg or Clone Hero_

1. Download and execute the Powershell script, or download the .exe file and double-click  

_All info is read from the **song.ini** or **songs.dta** in each song subfolder._

If you attempt to edit or run the script in Powershell, please note:

* _This script requires the **Export-Excel** cmdlet which is installed with the ImportExcel module.This module allows you to export data directly to .xlsx files without needing Excel installed on the system._  
  Open an elevated PowerShell prompt and run:
```
Install-Module -Name ImportExcel -Scope CurrentUser
```

* _If you are unable to execute scripts and are receiving an error similar to "Cannot be loaded because running scripts is disabled on this system" or "...execution of scripts is disabled on this system" you will need to enable PowerShell script execution on your PC.  More information is available at:_
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.5  
