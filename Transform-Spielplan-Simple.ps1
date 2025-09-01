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

# Add System.Web assembly for URL encoding
Add-Type -AssemblyName System.Web

# Function to load environment variables from .env file
function Get-EnvironmentConfig {
    $envFile = ".\.env"
    $config = @{}
    
    if (Test-Path $envFile) {
        Get-Content $envFile | ForEach-Object {
            if ($_ -match "^\s*([^#=]+)\s*=\s*(.*)$") {
                $key = $matches[1].Trim()
                $value = $matches[2].Trim()
                $config[$key] = $value
            }
        }
        Write-Host "Loaded configuration from .env file" -ForegroundColor Green
    } else {
        Write-Host "No .env file found. Copy .env.example to .env and configure your settings." -ForegroundColor Yellow
    }
    
    return $config
}

# Load environment configuration
$envConfig = Get-EnvironmentConfig

# Get configuration values with fallbacks
$GoogleMapsApiKey = if ($envConfig["GOOGLE_MAPS_API_KEY"]) { $envConfig["GOOGLE_MAPS_API_KEY"] } else { "" }
$homeTeamName = if ($envConfig["HOME_TEAM_NAME"]) { $envConfig["HOME_TEAM_NAME"] } else { "1. VSV Jena II" }
$homeTeamVenue = if ($envConfig["HOME_TEAM_VENUE"]) { $envConfig["HOME_TEAM_VENUE"] } else { "SH Lobdeburgschule (07747 Jena)" }
$responseDeadlineHours = if ($envConfig["RESPONSE_DEADLINE_HOURS"]) { [int]$envConfig["RESPONSE_DEADLINE_HOURS"] } else { 168 }  # 7 days default
$reminderHours = if ($envConfig["REMINDER_HOURS"]) { [int]$envConfig["REMINDER_HOURS"] } else { 336 }  # 14 days default

Write-Host "Reading CSV file..." -ForegroundColor Green

Write-Host "Configuration:" -ForegroundColor Cyan
Write-Host "  Home Team: $homeTeamName" -ForegroundColor White
Write-Host "  Home Venue: $homeTeamVenue" -ForegroundColor White
Write-Host "  Google Maps API: $(if ($GoogleMapsApiKey) { 'Configured' } else { 'Not configured (using static estimates)' })" -ForegroundColor White
Write-Host "  Response Deadline: $responseDeadlineHours hours ($([math]::Round($responseDeadlineHours / 24, 1)) days)" -ForegroundColor White
Write-Host "  Reminder Time: $reminderHours hours ($([math]::Round($reminderHours / 24, 1)) days)" -ForegroundColor White

if ($GoogleMapsApiKey -and $GoogleMapsApiKey -ne "your_google_maps_api_key_here") {
    Write-Host "`nGoogle Maps API Setup:" -ForegroundColor Cyan
    Write-Host "  To use Google Maps for real-time travel calculations, ensure:" -ForegroundColor Gray
    Write-Host "  1. Distance Matrix API is enabled in Google Cloud Console" -ForegroundColor Gray
    Write-Host "  2. Billing is set up for your Google Cloud project" -ForegroundColor Gray
    Write-Host "  3. Your API key has no restrictions blocking this usage" -ForegroundColor Gray
    Write-Host "  4. You haven't exceeded your quota limits" -ForegroundColor Gray
} elseif (!$GoogleMapsApiKey -or $GoogleMapsApiKey -eq "your_google_maps_api_key_here") {
    Write-Host "`nGoogle Maps API not configured. Using static travel time estimates." -ForegroundColor Yellow
    Write-Host "  To enable real-time calculations:" -ForegroundColor Gray
    Write-Host "  1. Copy .env.example to .env" -ForegroundColor Gray
    Write-Host "  2. Get API key from: https://console.cloud.google.com/apis/credentials" -ForegroundColor Gray
    Write-Host "  3. Enable Distance Matrix API in Google Cloud Console" -ForegroundColor Gray
    Write-Host "  4. Set up billing for your Google Cloud project" -ForegroundColor Gray
}

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
    
    # Try Google Maps API if API key is provided
    if (![string]::IsNullOrWhiteSpace($GoogleMapsApiKey)) {
        try {
            Write-Host "Calculating travel time from $origin to $destination using Google Maps..." -ForegroundColor Yellow
            
            # URL encode the addresses for the API
            $encodedOrigin = [System.Web.HttpUtility]::UrlEncode($origin)
            $encodedDestination = [System.Web.HttpUtility]::UrlEncode($destination)
            
            # Construct Google Maps Distance Matrix API URL
            $apiUrl = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=$encodedOrigin&destinations=$encodedDestination&mode=driving&language=de&key=$GoogleMapsApiKey"
            
            # Make API request
            $response = Invoke-RestMethod -Uri $apiUrl -Method Get
            
            # Check if API call was successful
            if ($response.status -eq "OK" -and $response.rows.Count -gt 0 -and $response.rows[0].elements.Count -gt 0) {
                $element = $response.rows[0].elements[0]
                
                if ($element.status -eq "OK") {
                    $durationSeconds = $element.duration.value
                    $durationMinutes = [Math]::Ceiling($durationSeconds / 60.0)
                    $distance = $element.distance.text
                    
                    Write-Host "Google Maps result: $distance, $durationMinutes minutes" -ForegroundColor Green
                    return $durationMinutes
                } else {
                    Write-Warning "Google Maps API element status: $($element.status)"
                }
            } elseif ($response.status -eq "REQUEST_DENIED") {
                Write-Warning "Google Maps API REQUEST_DENIED. Please check:"
                Write-Warning "  1. API key is valid and not expired"
                Write-Warning "  2. Distance Matrix API is enabled in Google Cloud Console"
                Write-Warning "  3. Billing is set up for your Google Cloud project"
                Write-Warning "  4. API key restrictions allow this domain/IP"
                Write-Warning "  5. You haven't exceeded your quota limits"
            } else {
                Write-Warning "Google Maps API response status: $($response.status)"
                if ($response.error_message) {
                    Write-Warning "Error message: $($response.error_message)"
                }
            }
        } catch {
            Write-Warning "Failed to query Google Maps API: $($_.Exception.Message)"
        }
    }
    
    # Fallback to static travel time estimates
    Write-Host "Using static travel time estimates for $origin to $destination" -ForegroundColor Cyan
    
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
        Write-Host "Found static route: $routeKey = $($travelMatrix[$routeKey]) minutes" -ForegroundColor Gray
        return $travelMatrix[$routeKey]
    }
    
    # Default travel time for unknown routes
    Write-Host "Using default travel time (90 minutes) for unknown route: $routeKey" -ForegroundColor Gray
    return 90
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
    
    # Calculate hours before game for deadlines (using configured values)
    # $responseDeadlineHours = hours before game for final response deadline (Zu-/Absagen bis)
    # $reminderHours = hours before game for reminder notification (Erinnerung zum Zu-/Absagen)
    
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
        'Zu-/Absagen bis (Stunden vor dem Termin)' = $responseDeadlineHours  # Configurable hours before game for response deadline
        'Erinnerung zum Zu-/Absagen (Stunden vor dem Termin)' = $reminderHours  # Configurable hours before game for reminder
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
