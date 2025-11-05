# SongInventory
Creates an Excel spreadsheet inventory of songs for Yarg or Clone Hero

This script requires the Export-Excel comdlet which is installed with the ImportExcel module.This module allows you to export data directly to .xlsx files without needing Excel installed on the system.

## Install the ImportExcel Module:
First, you need to install the ImportExcel module from the PowerShell Gallery. Open an elevated PowerShell prompt and run:

```
Install-Module -Name ImportExcel -Scope CurrentUser
```

Once installed, open the script in PowerShell ISE, modify the following variables and execute the script:
$rootDir
$excelOutputPath
