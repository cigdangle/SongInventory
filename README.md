# SongInventory
Creates an Excel spreadsheet inventory of songs for Yarg or Clone Hero

* _This script requires the **Export-Excel** cmdlet which is installed with the ImportExcel module.This module allows you to export data directly to .xlsx files without needing Excel installed on the system._  
  
* _If you are unable to execute scripts and are receiving an error similar to "Cannot be loaded because running scripts is disabled on this system" or "...execution of scripts is disabled on this system" you will need to enable PowerShell script execution on your PC.  More information is available at:_
https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.5

## 1. Install the ImportExcel Module
First, you need to install the ImportExcel module from the PowerShell Gallery (if not already installed). Open an elevated PowerShell prompt and run:

```
Install-Module -Name ImportExcel -Scope CurrentUser
```

## 2. Modify the Script
Open the script in PowerShell ISE (you will not need elevated privledges to modify or execute the script), modify the following variables:  
* **$songDirectory**  
* **$outputExcelFile**  

## 3. Execute the script
Run the script an enjoy the output  
_All info is read from the **song.ini** in each song subfolder._
