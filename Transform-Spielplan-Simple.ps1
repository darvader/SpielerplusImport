# Transform-Spielplan-Simple.ps1
# Script to transform Spielplan CSV data to Excel format
# Only imports data up to 'Geschlecht' column, ignoring the problematic duplicate columns

param(
    [string]$OutputPath = ".\Transformed_Spielplan_ExcelFormat.xlsx"
)

# Import required modules
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Force -Scope CurrentUser
    Import-Module ImportExcel
}

Write-Host "Reading CSV file..." -ForegroundColor Green

# Get the current directory and find CSV file
$currentDir = Get-Location
$csvFiles = Get-ChildItem -Path $currentDir -Filter "*Spielplan*.csv"

if ($csvFiles.Count -eq 0) {
    Write-Error "Could not find CSV file matching pattern '*Spielplan*.csv'"
    exit 1
}

$csvFullPath = $csvFiles[0].FullName
Write-Host "Found CSV file at: $csvFullPath" -ForegroundColor Green

# Read the raw CSV content
$csvContent = Get-Content -Path $csvFullPath -Encoding UTF8

# Function to fix encoding issues with German umlauts and ß
function Fix-GermanEncoding {
    param([string]$text)
    
    # Fix common German characters that appear as '�'
    $text = $text -replace "Oberwei�bach", "Oberweißbach"
    $text = $text -replace "Th�ringenliga", "Thüringenliga"
    $text = $text -replace "Th�ringen", "Thüringen"
    $text = $text -replace "Reinhard-He�", "Reinhard-Heß"
    $text = $text -replace "Nordstra�e", "Nordstraße"
    
    return $text
}

# Define the columns we want to keep (up to 'Geschlecht')
$desiredColumns = @(
    "Datum", "Uhrzeit", "Wochentag", "#", "ST", "Mannschaft 1", "Mannschaft 2", 
    "Schiedsgericht", "Gastgeber", "Austragungsort/Ergebnis", "Austragungsort", 
    "Ergebnis", "Saison", "Spielrunde", "Geschlecht"
)

Write-Host "Processing CSV with desired columns: $($desiredColumns -join ', ')" -ForegroundColor Cyan

# Parse each line manually to extract only the desired columns
$transformedData = @()

foreach ($line in $csvContent[1..($csvContent.Count - 1)]) {
    if ([string]::IsNullOrWhiteSpace($line)) {
        continue
    }
    
    # Split the line by semicolon and clean quotes
    $fields = $line -split ';' | ForEach-Object { $_.Trim('"') }
    
    # Skip if we don't have enough fields
    if ($fields.Count -lt 15) {
        continue
    }
    
    # Extract only the first 15 fields (up to Geschlecht)
    $datum = $fields[0]
    $uhrzeit = $fields[1]
    $wochentag = $fields[2]
    $nummer = $fields[3]
    $st = $fields[4]
    $mannschaft1 = Fix-GermanEncoding $fields[5]
    $mannschaft2 = Fix-GermanEncoding $fields[6]
    $schiedsgericht = Fix-GermanEncoding $fields[7]
    $gastgeber = Fix-GermanEncoding $fields[8]
    $austragungsortErgebnis = Fix-GermanEncoding $fields[9]
    $austragungsort = Fix-GermanEncoding $fields[10]
    $ergebnis = $fields[11]
    $saison = $fields[12]
    $spielrunde = Fix-GermanEncoding $fields[13]
    $geschlecht = $fields[14]
    
    # Skip rows without essential data
    if ([string]::IsNullOrWhiteSpace($datum) -or [string]::IsNullOrWhiteSpace($mannschaft1)) {
        continue
    }
    
    # Filter only games with '1. VSV Jena II'
    if ($mannschaft1 -ne "1. VSV Jena II" -and $mannschaft2 -ne "1. VSV Jena II") {
        continue
    }
    
    # Parse date and time
    $gameDate = $null
    $gameTime = $null
    
    try {
        if (![string]::IsNullOrWhiteSpace($datum)) {
            $gameDate = [DateTime]::ParseExact($datum, "dd.MM.yyyy", $null)
        }
        if (![string]::IsNullOrWhiteSpace($uhrzeit)) {
            $gameTime = [DateTime]::ParseExact($uhrzeit, "HH:mm:ss", $null)
        }
    } catch {
        Write-Warning "Could not parse date/time for row: $nummer"
        continue
    }
    
    # Calculate dates for deadlines
    $oneWeekBefore = if ($gameDate) { $gameDate.AddDays(-7) } else { $null }
    $twoWeeksBefore = if ($gameDate) { $gameDate.AddDays(-14) } else { $null }
    
    # Create info text with referee and game ID
    $gameInfo = "$st - $spielrunde | Spiel-ID: $nummer | Schiedsrichter: $schiedsgericht"
    
    # Determine home game and opponent based on which position '1. VSV Jena II' is in
    $isHomeGame = $false
    $opponent = ""
    
    if ($mannschaft1 -eq "1. VSV Jena II") {
        # VSV Jena II is Mannschaft 1, so it's a home game
        $isHomeGame = $true
        $opponent = $mannschaft2
    } else {
        # VSV Jena II is Mannschaft 2, so it's an away game
        $isHomeGame = $false
        $opponent = $mannschaft1
    }
    
    # Create transformed record matching Excel template columns exactly
    $transformedRecord = [PSCustomObject]@{
        'Spieltyp (Opptional)' = "Spiel"  # Default to "Spiel" for all volleyball games
        'Gegner' = $opponent  # The other team (opponent)
        'Start-Datum' = if ($gameDate) { $gameDate } else { $null }
        'End-Datum' = if ($gameDate) { $gameDate } else { $null }  # Same as start date for volleyball games
        'Start-Zeit' = if ($gameTime) { $gameTime.TimeOfDay.TotalDays } else { $null }  # Excel time format
        'End-Zeit (Optional)' = if ($gameTime) { ($gameTime.AddHours(2)).TimeOfDay.TotalDays } else { $null }  # Estimated 2-hour duration
        'Treffen (Optional)' = if ($gameTime) { ($gameTime.AddMinutes(-30)).TimeOfDay.TotalDays } else { $null }  # 30 min before game
        'Heimspiel' = if ($isHomeGame) { "TRUE" } else { "FALSE" }  # True when VSV Jena II is Mannschaft 1
        'Gelände / Räumlichkeiten' = $austragungsort
        'Adresse (Optional)' = ""  # Not available in source data
        'Infos zum Spiel (Optional)' = $gameInfo  # Round, league, game ID and referee info
        'Nominierung (Optional)' = ""  # Not available in source data
        'Teilname (Optional)' = ""  # Not available in source data
        'Zu-/Absagen bis (Stunden vor dem Termin)' = if ($oneWeekBefore) { $oneWeekBefore } else { $null }  # One week before game
        'Erinnerung zum Zu-/Absagen (Stunden vor dem Termin)' = if ($twoWeeksBefore) { $twoWeeksBefore } else { $null }  # Two weeks before game
        # Additional fields for reference (not in Excel template but useful)
        '_Season' = $saison
        '_Gender' = $geschlecht
        '_Result' = $ergebnis
    }
    
    $transformedData += $transformedRecord
}

Write-Host "Transformed $($transformedData.Count) records" -ForegroundColor Green

# Export to Excel
Write-Host "Exporting to Excel..." -ForegroundColor Green

try {
    $transformedData | Export-Excel -Path $OutputPath -WorksheetName "Games" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    Write-Host "Successfully exported to: $OutputPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to export to Excel: $($_.Exception.Message)"
    
    # Fallback: export to CSV
    $csvOutputPath = $OutputPath -replace "\.xlsx$", ".csv"
    $transformedData | Export-Csv -Path $csvOutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to CSV instead: $csvOutputPath" -ForegroundColor Yellow
}

# Display summary
Write-Host "`nTransformation Summary:" -ForegroundColor Cyan
Write-Host "- Transformed records: $($transformedData.Count)" -ForegroundColor White
Write-Host "- Output file: $OutputPath" -ForegroundColor White

# Show sample of transformed data
Write-Host "`nSample of transformed data:" -ForegroundColor Cyan
$transformedData | Select-Object -First 5 | Format-Table -Property 'Spieltyp (Opptional)', 'Gegner', 'Start-Datum', 'Start-Zeit', 'Heimspiel', 'Gelände / Räumlichkeiten' -AutoSize

Write-Host "`nColumns included in the output:" -ForegroundColor Cyan
$transformedData[0].PSObject.Properties.Name | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
