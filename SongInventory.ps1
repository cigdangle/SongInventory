# Change these values as needed. Set the directory to search and the output file path
$songDirectory = "E:\Clone Hero Assets\Songs"
$outputExcelFile = "E:\Clone Hero Assets\SongData.xlsx"

# If output file already exists, delete it in prepreation for a new one
if (Test-Path -Path $outputExcelFile) {
    Remove-Item -Path $outputExcelFile -Force
    Write-Host "File '$outputExcelFile' removed successfully."
}


# This array will store all the custom objects created from the .ini files.
$songDataCollection = @()

# Use Get-ChildItem to find all files named 'song.ini' in subdirectories.
$iniFiles = Get-ChildItem -Path $songDirectory -Filter "song.ini" -Recurse

# Loop through each found file.
foreach ($file in $iniFiles) {
    # Get the file's content as a single string.
    $content = Get-Content -Path $file.FullName | Out-String

    # Use a hashtable to hold the key-value pairs from the .ini file.
    # The default values are set to empty strings.
    $data = @{
        'Title' = ''
        'Artist' = ''
        'Album' = ''
        'Genre' = ''
        'Year' = ''
        'Guitar' = ''
        'Vocals' = ''
        'Drums' = ''
        'Bass' = ''
        'Keys' = ''
    }

    # Define the regular expression pattern to match 'key = value'.
    # This pattern captures the key and its corresponding value.
    $pattern = '(?<key>[a-zA-Z_]+)\s*=\s*(?<value>.*)'

    # Process each line of the file content.
    $content.Split([System.Environment]::NewLine) | ForEach-Object {
        if ($_ -match $pattern) {
            $key = $Matches.key.Trim()
            $value = $Matches.value.Trim()

            # Map the .ini keys to the desired output columns.
            switch ($key) {
                'artist' { $data['Artist'] = $value }
                'name' { $data['Title'] = $value }
                'album' { $data['Album'] = $value }
                'genre' { $data['Genre'] = $value }
                'year' { $data['Year'] = $value }
                'diff_guitar' { $data['Guitar'] = $value }
                'diff_vocals' { $data['Vocals'] = $value }
                'diff_drums' { $data['Drums'] = $value }
                'diff_bass' { $data['Bass'] = $value }
                'diff_keys' { $data['Keys'] = $value }
            }
        }
    }

    # Create a custom PowerShell object from the collected data and add it to the collection.
    $songDataObject = [PSCustomObject]$data
    $songDataCollection += $songDataObject
}

# Export the final collection of objects to an Excel file.
# The `-AutoSize` parameter automatically adjusts column widths for better readability, and TableStyle makes it pretty!
$sortedSongData = $songDataCollection | Sort-Object -Property Artist
$sortedSongData| Select-Object Artist, Title, Album, Year, Genre, Guitar, Bass, Vocals, Drums, Keys | Export-Excel -Path $outputExcelFile -WorksheetName "Song Data" -AutoSize -TableStyle:'Medium2'-Show -CellStyleSB {
    param($workSheet)
    # Align Artist, Title, and Album left
    $workSheet.Cells["A:C"].Style.HorizontalAlignment = "Left"
    # Align Year (except the header) right
    $workSheet.Cells["D2:D"].Style.HorizontalAlignment = "Right"
    # Align Genre left
    $workSheet.Cells["E:E"].Style.HorizontalAlignment = "Left"
    # Align instrument difficulties center
    $workSheet.Cells["F:J"].Style.HorizontalAlignment = "Center"
}
# Acknowledge file created
Write-Host "Successfully exported data to $outputExcelFile"
