# Requires the ImportExcel module. Install with:
# Install-Module -Name ImportExcel -Scope CurrentUser
Add-Type -AssemblyName System.Windows.Forms

#Select Song Path
$songBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$songBrowser.Description = "Please select your songs folder"
$songBrowser.SelectedPath = [Environment]::GetFolderPath("Desktop") # Set initial path to My Documents
$songDirectory = $songBrowser.ShowDialog()
$songDirectory = $songBrowser.SelectedPath

#Select Output Path
$outputBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$outputBrowser.Description = "Please select your output folder"
$outputBrowser.SelectedPath = [Environment]::GetFolderPath("Desktop") # Set initial path to My Documents
$outputExcelFile = $outputBrowser.ShowDialog()
$outputExcelFile = $songBrowser.SelectedPath + "\SongData.xlsx"

function Parse-SongIni ($filePath) {
    $data = @{
        Name    = $null
        Artist  = $null
        Album   = $null
        Genre   = $null
        Year    = $null
        Guitar  = $null
        Vocals  = $null
        Drums   = $null
        Bass    = $null
        Keys    = $null
    }

    # Use a regex-based switch to parse key-value pairs
    switch -regex -file $filePath {
        '^name\s*=\s*(.*)'   { $data.Name = $matches[1].Trim() }
        '^artist\s*=\s*(.*)' { $data.Artist = $matches[1].Trim() }
        '^album\s*=\s*(.*)'  { $data.Album = $matches[1].Trim() }
        '^genre\s*=\s*(.*)'  { $data.Genre = $matches[1].Trim() }
        '^year\s*=\s*(.*)'   { $data.Year = $matches[1].Trim() }
        # Check for instrument difficulty keys to determine presence
        '^diff_guitar\s*=\s*(.*)'  { $data.Guitar = $matches[1] }
        '^diff_vocals\s*=\s*(.*)'  { $data.Vocals = $matches[1] }
        '^diff_drums\s*=\s*(.*)'   { $data.Drum = $matches[1] }
        '^diff_bass\s*=\s*(.*)'    { $data.Bass = $matches[1] }
        '^diff_keys\s*=\s*(.*)'    { $data.Keys = $matches[1] }
    }
    # If no diff keys are found, check for the 'frets' or 'pro_drums' key which often implies presence
    if (-not $data.Guitar -and (Get-Content $filePath | Select-String -Pattern '^frets\s*=')) {$data.Guitar = "Yes"}
    if (-not $data.Drum -and (Get-Content $filePath | Select-String -Pattern '^pro_drums\s*=')) {$data.Drum = "Yes"}

    return New-Object PSObject -Property $data
}

function Parse-SongsDta ($filePath) {
    $content = Get-Content -Path $filePath -Raw
    $data = @{
        Name    = $null
        Artist  = $null
        Album   = $null
        Genre   = $null
        Year    = $null
        Guitar  = $null
        Vocals  = $null
        Drums   = $null
        Bass    = $null
        Keys    = $null
    }

    # Use regex to extract data. The '(?s)' makes the '.' match newlines
    if ($content -match '(?s)\(\s*''name''\s*"(.+?)"\s*\)') {
        $data.Name = $matches[1]
    }
    if ($content -match '(?s)\(\s*''artist''\s*"(.+?)"\s*\)') {
        $data.Artist = $matches[1]
    }
    # Look for 'album_name' or a general album entry
    if ($content -match '(?s)\(\s*''album_name''\s*"(.+?)"\s*\)') {
        $data.Album = $matches[1]
    } elseif ($content -match '(?s)\(\s*''album''\s*"(.+?)"\s*\)') {
        $data.Album = $matches[1]
    }
    # Look for 'genre'
    if ($content -match '(?s)\(\s*''genre''\s*''?(.+?)''''?\s*\)') {
        $data.Genre = $matches[1]
    }
    # Look for 'year_released' or 'year'
    if ($content -match '(?s)\(\s*''year_released''\s*(\d+)\s*\)') {
        $data.Year = $matches[1]
    } elseif ($content -match '(?s)\(\s*''year''\s*(\d+)\s*\)') {
        $data.Year = $matches[1]
    }

# Instruments to extract
$instruments = @('drum', 'guitar', 'bass', 'vocals', 'keys')

foreach ($instrument in $instruments) {
    # Define a simple regex pattern that captures the number using named group 'value'
    # This specifically targets the format ('instrument' number)
    $regexPattern = "^\s*\('$instrument'\s+(?<value>\d+)\)"

    # Iterate line by line to find the match within the file content
    $value = $null
    foreach ($line in $content -split "`n") {
        if ($line -match $regexPattern) {
            # If the line matches, the value is captured in the $Matches hashtable under the group name 'value'
            $value = $Matches['value']
            $data.$instrument = $value
        }
    }
}
    return New-Object PSObject -Property $data
}

# --- Main Script ---

# Check if the output directory exists, create if not
$outputDir = Split-Path -Path $outputExcelFile -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

# If output file already exists, delete it in prepreation for a new one
if (Test-Path -Path $outputExcelFile) {
    Remove-Item -Path $outputExcelFile -Force
    Write-Host "File '$outputExcelFile' removed successfully."
}

$foundFiles = Get-ChildItem -Path $songDirectory -Recurse -File | Where-Object { $_.Name -eq 'song.ini' -or $_.Name -eq 'songs.dta' }

$allSongData = foreach ($file in $foundFiles) {
    if ($file.Name -eq 'song.ini') {
        Write-Host "Parsing INI file: $($file.FullName)"
        Parse-SongIni -filePath $file.FullName | Select-Object Name, Artist, Album, Genre, Year, Guitar, Vocals, Drum, Bass, Keys
    } elseif ($file.Name -eq 'songs.dta') {
        Write-Host "Parsing DTA file: $($file.FullName)"
        Parse-SongsDta -filePath $file.FullName | Select-Object Name, Artist, Album, Genre, Year, Guitar, Vocals, Drum, Bass, Keys
    }
}

if ($allSongData) {
    # Sort Data by Artist and Song
    $sortedSongData = $allSongData | Sort-Object -Property Artist, Name
    
    # Export to Excel spreadsheet
    $sortedSongData | Export-Excel -Path $outputExcelFile  -AutoSize -TableStyle:'Medium2'-Show -CellStyleSB {
        param($worksheet)

        # Left-align cells in columns A thru D
        $worksheet.Cells["A:D"].Style.HorizontalAlignment = "Left"

        # Center-align all cells in columns E thru J
        $worksheet.Cells["E:J"].Style.HorizontalAlignment = "Center"
}
    Write-Host "Successfully exported data to $outputExcelFile" -ForegroundColor Green
} else {
    Write-Host "No song.ini or songs.dta files found in $songDirectory." -ForegroundColor Yellow
}
