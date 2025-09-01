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

# Configuration
$homeTeamName = "1. VSV Jena II"
$homeTeamVenue = "SH Lobdeburgschule (07747 Jena)"  # Home venue for travel calculations

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

# Function to extract postal code and city from venue string (from brackets)
function Get-PostalCodeAndCity {
    param([string]$venue)
    
    # Extract content from brackets like "(07747 Jena)" or "(99096 Erfurt)"
    if ($venue -match '\((\d{5})\s+([^)]+)\)') {
        $postalCode = $matches[1]
        $city = $matches[2]
        return "$postalCode $city, Germany"
    }
    
    # Fallback - try to extract city name without postal code
    if ($venue -like "*Jena*") { return "07747 Jena, Germany" }
    if ($venue -like "*Erfurt*") { return "99096 Erfurt, Germany" }
    if ($venue -like "*Weimar*") { return "99427 Weimar, Germany" }
    if ($venue -like "*Gera*") { return "07546 Gera, Germany" }
    if ($venue -like "*Meiningen*") { return "98617 Meiningen, Germany" }
    if ($venue -like "*Suhl*") { return "98529 Suhl, Germany" }
    if ($venue -like "*Altenburg*") { return "04600 Altenburg, Germany" }
    if ($venue -like "*Bleicherode*") { return "95752 Bleicherode, Germany" }
    if ($venue -like "*Oberweißbach*") { return "98744 Oberweißbach, Germany" }
    
    return "Unknown Location, Germany"
}

# Function to estimate travel time from origin to destination (in minutes)
function Get-TravelTime {
    param(
        [string]$originVenue,
        [string]$destinationVenue
    )
    
    # Extract postal codes and cities from both venues
    $origin = Get-PostalCodeAndCity $originVenue
    $destination = Get-PostalCodeAndCity $destinationVenue
    
    # If origin and destination are the same, it's a local game (0 travel time)
    if ($origin -eq $destination) {
        return 0
    }
    
    # Try to get travel time from Google Maps API (placeholder for future implementation)
    try {
        # Future Google Maps API integration would go here
        # For now, use fallback static values based on cities
        
        # Extract city names for fallback lookup
        $originCity = if ($origin -match '\d{5}\s+([^,]+)') { $matches[1] } else { "Unknown" }
        $destCity = if ($destination -match '\d{5}\s+([^,]+)') { $matches[1] } else { "Unknown" }
        
        # Static travel time matrix (in minutes) from various cities
        $travelMatrix = @{
            "Jena-Erfurt" = 60
            "Jena-Weimar" = 45
            "Jena-Gera" = 45
            "Jena-Meiningen" = 90
            "Jena-Suhl" = 75
            "Jena-Altenburg" = 60
            "Jena-Bleicherode" = 120
            "Jena-Oberweißbach" = 60
            # Add reverse routes
            "Erfurt-Jena" = 60
            "Weimar-Jena" = 45
            "Gera-Jena" = 45
            "Meiningen-Jena" = 90
            "Suhl-Jena" = 75
            "Altenburg-Jena" = 60
            "Bleicherode-Jena" = 120
            "Oberweißbach-Jena" = 60
        }
        
        $routeKey = "$originCity-$destCity"
        if ($travelMatrix.ContainsKey($routeKey)) {
            return $travelMatrix[$routeKey]
        }
        
        # Default travel time for unknown routes
        return 90
        
    } catch {
        Write-Warning "Could not calculate travel time from $origin to $destination. Using default."
        return 90
    }
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
    
    # Filter only games with the home team
    if ($mannschaft1 -ne $homeTeamName -and $mannschaft2 -ne $homeTeamName) {
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
            # If Uhrzeit is "00:00:00", use "11:00:00" instead
            $timeToUse = if ($uhrzeit -eq "00:00:00") { "11:00:00" } else { $uhrzeit }
            $gameTime = [DateTime]::ParseExact($timeToUse, "HH:mm:ss", $null)
        }
    } catch {
        Write-Warning "Could not parse date/time for row: $nummer"
        continue
    }
    
    # Calculate hours before game for deadlines
    $hoursOneWeekBefore = 7 * 24  # 7 days = 168 hours
    $hoursTwoWeeksBefore = 14 * 24  # 14 days = 336 hours
    
    # Create info text with referee and game ID
    $gameInfo = "$st - $spielrunde | Spiel-ID: $nummer | Schiedsrichter: $schiedsgericht"
    
    # Determine home game and opponent based on which position the home team is in
    $isHomeGame = $false
    $opponent = ""
    
    if ($mannschaft1 -eq $homeTeamName) {
        # Home team is Mannschaft 1, so it's a home game
        $isHomeGame = $true
        $opponent = $mannschaft2
    } else {
        # Home team is Mannschaft 2, so it's an away game
        $isHomeGame = $false
        $opponent = $mannschaft1
    }
    
    # Calculate meeting time (Treffen)
    $treffenTime = ""
    if ($gameTime) {
        if ($isHomeGame) {
            # Home game: 2 hours before start time
            $treffenTime = ($gameTime.AddHours(-2)).ToString("HH:mm:ss")
        } else {
            # Away game: travel time + 60 min buffer before start time
            $travelMinutes = Get-TravelTime $homeTeamVenue $austragungsort
            $totalMinutesEarly = $travelMinutes + 60
            $treffenTime = ($gameTime.AddMinutes(-$totalMinutesEarly)).ToString("HH:mm:ss")
        }
    }
    
    # Create transformed record matching Excel template columns exactly
    $transformedRecord = [PSCustomObject]@{
        'Spieltyp (Opptional)' = "Spiel"  # Default to "Spiel" for all volleyball games
        'Gegner' = $opponent  # The other team (opponent)
        'Start-Datum' = if ($gameDate) { $gameDate.ToString("dd.MM.yyyy") } else { "" }
        'End-Datum' = if ($gameDate) { $gameDate.ToString("dd.MM.yyyy") } else { "" }  # Same as start date for volleyball games
        'Start-Zeit' = if ($gameTime) { $gameTime.ToString("HH:mm:ss") } else { "" }  # Time format hh:mm:ss from Uhrzeit
        'End-Zeit (Optional)' = if ($gameTime) { ($gameTime.AddHours(8)).ToString("HH:mm:ss") } else { "" }  # Game time plus 8 hours as hh:mm:ss
        'Treffen (Optional)' = $treffenTime  # Meeting time based on home/away game logic
        'Heimspiel' = if ($isHomeGame) { "TRUE" } else { "FALSE" }  # True when home team is Mannschaft 1
        'Gelände / Räumlichkeiten' = $austragungsort
        'Adresse (Optional)' = ""  # Not available in source data
        'Infos zum Spiel (Optional)' = $gameInfo  # Round, league, game ID and referee info
        'Nominierung (Optional)' = ""  # Not available in source data
        'Teilname (Optional)' = ""  # Not available in source data
        'Zu-/Absagen bis (Stunden vor dem Termin)' = $hoursOneWeekBefore  # 168 hours (7 days) before game
        'Erinnerung zum Zu-/Absagen (Stunden vor dem Termin)' = $hoursTwoWeeksBefore  # 336 hours (14 days) before game
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
    # Simple Excel export without complex formatting to avoid corruption
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
