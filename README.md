# SongInventory
Creates an Excel spreadsheet inventory of songs for Yarg or Clone Hero

This script requires the Export-Excel comdlet which is installed with the ImportExcel module.This module allows you to export data directly to .xlsx files without needing Excel installed on the system.

## 1. Install the ImportExcel Module
First, you need to install the ImportExcel module from the PowerShell Gallery. Open an elevated PowerShell prompt and run:

```
Install-Module -Name ImportExcel -Scope CurrentUser
```

## 2. Modify the Script
Open the script in PowerShell ISE (you will not need elevated privledges to modify or execute the script), modify the following variables:  
* **$rootDir**  
* **$excelOutputPath**  

## 3. Execute the script
Run the script an enjoy the output  
_All info is read from the **song.ini** in each song subfolder._
