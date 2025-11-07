# Requires the ImportExcel module. Install with:
# Install-Module -Name ImportExcel -Scope CurrentUser

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
        '^diff_drums\s*=\s*(.*)'   { $data.Drums = $matches[1] }
        '^diff_bass\s*=\s*(.*)'    { $data.Bass = $matches[1] }
        '^diff_keys\s*=\s*(.*)'    { $data.Keys = $matches[1] }
    }
    # If no diff keys are found, check for the 'frets' or 'pro_drums' key which often implies presence
    if (-not $data.Guitar -and (Get-Content $filePath | Select-String -Pattern '^frets\s*=')) {$data.Guitar = "Yes"}
    if (-not $data.Drums -and (Get-Content $filePath | Select-String -Pattern '^pro_drums\s*=')) {$data.Drums = "Yes"}

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

    # Check for instrument presence in the 'rank' section
    #if ($content -match "(?s)\('rank'(.*?)\)\s*\('genre'")  {
    #    $rank = $matches[1].Trim()
    #    Write-Host "match"
        if ($content -match "'rank'\s*\((?:[^)]*?\s*'guitar'\s*(\d+)|[^)]*)\)") { $data.Guitar = $matches[1]}
        #if ($tracks -match 'guitar''?\s*\(([\d\s-]+)\)') { $data.Guitar = "poop" }
        if ($rank -match 'vocals') { $data.Vocals = "Yes" }
        if ($rank -match 'drum') { $data.Drums = "Yes" }
        if ($rank -match 'bass') { $data.Bass = "Yes" }
        if ($rank -match 'keys') { $data.Keys = "Yes" }
    #}

    return New-Object PSObject -Property $data
}

# --- Main Script ---
$startDirectory = "E:\Clone Hero Assets\Songs" # <--- Change this to your starting directory
$outputFile = "E:\Clone Hero Assets\Song_Data.xlsx"   # <--- Change this to your desired output path

# Check if the output directory exists, create if not
$outputDir = Split-Path -Path $outputFile -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

# If output file already exists, delete it in prepreation for a new one
if (Test-Path -Path $outputFile) {
    Remove-Item -Path $outputFile -Force
    Write-Host "File '$outputFile' removed successfully."
}

$foundFiles = Get-ChildItem -Path $startDirectory -Recurse -File | Where-Object { $_.Name -eq 'song.ini' -or $_.Name -eq 'songs.dta' }

$allSongData = foreach ($file in $foundFiles) {
    if ($file.Name -eq 'song.ini') {
        Write-Host "Parsing INI file: $($file.FullName)"
        Parse-SongIni -filePath $file.FullName | Select-Object Name, Artist, Album, Genre, Year, Guitar, Vocals, Drums, Bass, Keys
    } elseif ($file.Name -eq 'songs.dta') {
        Write-Host "Parsing DTA file: $($file.FullName)"
        Parse-SongsDta -filePath $file.FullName | Select-Object Name, Artist, Album, Genre, Year, Guitar, Vocals, Drums, Bass, Keys
    }
}

if ($allSongData) {
    # Export to Excel spreadsheet
    $allSongData | Export-Excel -Path $outputFile  -AutoSize -TableStyle:'Medium2'-Show 
    Write-Host "Successfully exported data to $outputFile" -ForegroundColor Green
} else {
    Write-Host "No song.ini or songs.dta files found in $startDirectory." -ForegroundColor Yellow
}
