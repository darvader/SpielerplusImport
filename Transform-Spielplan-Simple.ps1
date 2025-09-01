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

# Add System.Web as# Display summary
Write-Host "`nTransformation Summary:" -ForegroundColor Cyan
Write-Host "- Transformed records: $($transformedData.Count)" -ForegroundColor White
Write-Host "- Output file: $OutputPath" -ForegroundColor White

# Show OpenAI usage statistics
if ($global:openAICallCount -gt 0 -or $global:germanTextCache.Count -gt 0) {
    Write-Host "`nOpenAI German Text Correction Summary:" -ForegroundColor Cyan
    Write-Host "- API calls made: $global:openAICallCount" -ForegroundColor White
    Write-Host "- Cached corrections: $($global:germanTextCache.Count)" -ForegroundColor White
    if ($global:germanTextCache.Count -gt 0) {
        Write-Host "- Cache contents:" -ForegroundColor White
        $global:germanTextCache.GetEnumerator() | ForEach-Object {
            Write-Host "  '$($_.Key)' -> '$($_.Value)'" -ForegroundColor Gray
        }
    }
}

# Show sample of transformed data
Write-Host "`nSample of transformed data:" -ForegroundColor Cyan
$transformedData | Select-Object -First 5 | Format-Table -Property 'Spieltyp (Opptional)', 'Gegner', 'Start-Datum', 'Start-Zeit', 'Heimspiel', 'Gelände / Räumlichkeiten' -AutoSize

Write-Host "`nColumns included in the output:" -ForegroundColor Cyan
$transformedData[0].PSObject.Properties.Name | ForEach-Object { Write-Host "  $_" -ForegroundColor White }URL encoding
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
$GoogleApiKey = if ($envConfig["GOOGLE_API_KEY"]) { $envConfig["GOOGLE_API_KEY"] } else { "" }
$OpenAIApiKey = if ($envConfig["OPENAI_API_KEY"]) { $envConfig["OPENAI_API_KEY"] } else { "" }
$homeTeamName = if ($envConfig["HOME_TEAM_NAME"]) { $envConfig["HOME_TEAM_NAME"] } else { "1. VSV Jena II" }
$homeTeamVenue = if ($envConfig["HOME_TEAM_VENUE"]) { $envConfig["HOME_TEAM_VENUE"] } else { "SH Lobdeburgschule (07747 Jena)" }
$responseDeadlineHours = if ($envConfig["RESPONSE_DEADLINE_HOURS"]) { [int]$envConfig["RESPONSE_DEADLINE_HOURS"] } else { 168 }  # 7 days default
$reminderHours = if ($envConfig["REMINDER_HOURS"]) { [int]$envConfig["REMINDER_HOURS"] } else { 336 }  # 14 days default

# Global cache for German text corrections to avoid repeated API calls
$global:germanTextCache = @{}
$global:lastOpenAICall = [DateTime]::MinValue
$global:openAICallCount = 0

Write-Host "Reading CSV file..." -ForegroundColor Green

Write-Host "Configuration:" -ForegroundColor Cyan
Write-Host "  Home Team: $homeTeamName" -ForegroundColor White
Write-Host "  Home Venue: $homeTeamVenue" -ForegroundColor White
Write-Host "  Google API: $(if ($GoogleApiKey) { 'Configured' } else { 'Not configured (using static estimates)' })" -ForegroundColor White
Write-Host "  Response Deadline: $responseDeadlineHours hours ($([math]::Round($responseDeadlineHours / 24, 1)) days)" -ForegroundColor White
Write-Host "  Reminder Time: $reminderHours hours ($([math]::Round($reminderHours / 24, 1)) days)" -ForegroundColor White

if ($GoogleApiKey -and $GoogleApiKey -ne "your_google_api_key_here") {
    Write-Host "`nGoogle API Setup:" -ForegroundColor Cyan
    Write-Host "  To use Google services for travel calculations and text encoding, ensure:" -ForegroundColor Gray
    Write-Host "  1. Distance Matrix API is enabled in Google Cloud Console" -ForegroundColor Gray
    Write-Host "  2. Cloud Translation API is enabled in Google Cloud Console" -ForegroundColor Gray
    Write-Host "  3. Billing is set up for your Google Cloud project" -ForegroundColor Gray
    Write-Host "  4. Your API key has no restrictions blocking this usage" -ForegroundColor Gray
} elseif (!$GoogleApiKey -or $GoogleApiKey -eq "your_google_api_key_here") {
    Write-Host "`nGoogle API not configured. Using static estimates and local text fixes." -ForegroundColor Yellow
    Write-Host "  To enable Google services:" -ForegroundColor Gray
    Write-Host "  1. Copy .env.example to .env" -ForegroundColor Gray
    Write-Host "  2. Get API key from: https://console.cloud.google.com/apis/credentials" -ForegroundColor Gray
    Write-Host "  3. Enable Distance Matrix API and Cloud Translation API" -ForegroundColor Gray
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
    
    # Early return if text is empty or null
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $text
    }
    
    # Try OpenAI API for intelligent German character restoration if API key is available
    if (![string]::IsNullOrWhiteSpace($OpenAIApiKey) -and $text -match '�') {
        
        # Check cache first
        if ($global:germanTextCache.ContainsKey($text)) {
            $cachedResult = $global:germanTextCache[$text]
            Write-Host "Using cached result: '$text' -> '$cachedResult'" -ForegroundColor Cyan
            return $cachedResult
        }
        
        # Rate limiting: max 3 calls per minute for free tier
        $now = Get-Date
        if ($now.Subtract($global:lastOpenAICall).TotalSeconds -lt 60) {
            if ($global:openAICallCount -ge 3) {
                $waitTime = 60 - $now.Subtract($global:lastOpenAICall).TotalSeconds
                Write-Host "Rate limit reached. Waiting $([math]::Ceiling($waitTime)) seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds ([math]::Ceiling($waitTime))
                $global:openAICallCount = 0
                $global:lastOpenAICall = Get-Date
            }
        } else {
            # Reset counter if more than a minute has passed
            $global:openAICallCount = 0
            $global:lastOpenAICall = $now
        }
        
        try {
            Write-Host "Using OpenAI API to fix German encoding for: $text (Call #$($global:openAICallCount + 1))" -ForegroundColor Yellow
            
            $apiUrl = "https://api.openai.com/v1/chat/completions"
            
            $prompt = @"
Fix the corrupted German text by replacing the � symbols with the correct German characters (ä, ö, ü, ß). 
This is volleyball schedule data from Thuringia, Germany. Context: team names, venue names, city names.

Original text: "$text"

Rules:
1. Only replace � symbols with correct German characters
2. Keep all other text exactly the same
3. Consider German place names, street names, and common German words
4. Return ONLY the corrected text, no explanations

Corrected text:
"@

            $requestBody = @{
                model = "gpt-3.5-turbo"
                messages = @(
                    @{
                        role = "user"
                        content = $prompt
                    }
                )
                max_tokens = 100
                temperature = 0.1
            } | ConvertTo-Json -Depth 3

            $headers = @{
                "Authorization" = "Bearer $OpenAIApiKey"
                "Content-Type" = "application/json"
            }
            
            Write-Host "Sending request to OpenAI API..." -ForegroundColor Cyan
            
            $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Body $requestBody -Headers $headers
            $global:openAICallCount++
            
            if ($response.choices -and $response.choices.Count -gt 0) {
                $correctedText = $response.choices[0].message.content.Trim()
                
                # Remove any quotes that might be added by the API
                $correctedText = $correctedText -replace '^"', '' -replace '"$', ''
                
                # Cache the result
                $global:germanTextCache[$text] = $correctedText
                
                Write-Host "OpenAI result: '$text' -> '$correctedText'" -ForegroundColor Green
                return $correctedText
            } else {
                Write-Warning "OpenAI API returned unexpected response format"
            }
        } catch {
            Write-Warning "OpenAI API failed: $($_.Exception.Message)"
            if ($_.ErrorDetails.Message) {
                Write-Host "API Error Details: $($_.ErrorDetails.Message)" -ForegroundColor Red
            }
            # Fall back to local replacement rules
        }
    }

    # Try smart German character restoration using context and common patterns
    if ($text -match '�') {
        Write-Host "Using smart German character restoration for: $text" -ForegroundColor Yellow
        
        # Create a comprehensive mapping based on common German volleyball terms and places
        $restoredText = $text
        
        # City names and places (most common in Thuringia volleyball)
        $restoredText = $restoredText -replace "Oberwei�bach", "Oberweißbach"
        $restoredText = $restoredText -replace "Th�ringenliga", "Thüringenliga"
        $restoredText = $restoredText -replace "Th�ringen", "Thüringen"
        
        # Street names and venues
        $restoredText = $restoredText -replace "Nordstra�e", "Nordstraße"
        $restoredText = $restoredText -replace "Stra�e", "Straße"
        $restoredText = $restoredText -replace "Gro�", "Groß"
        
        # Person names (common German surnames)
        $restoredText = $restoredText -replace "Reinhard-He�", "Reinhard-Heß"
        $restoredText = $restoredText -replace "Fr�bel", "Fröbel"
        $restoredText = $restoredText -replace "M�ller", "Müller"
        $restoredText = $restoredText -replace "Sch�fer", "Schäfer"
        $restoredText = $restoredText -replace "Kr�ger", "Krüger"
        
        # Common German words in sports context
        $restoredText = $restoredText -replace "Sporthalle", "Sporthalle"  # Already correct
        $restoredText = $restoredText -replace "Turnhalle", "Turnhalle"    # Already correct
        $restoredText = $restoredText -replace "M�nchen", "München"
        $restoredText = $restoredText -replace "N�rnberg", "Nürnberg"
        $restoredText = $restoredText -replace "W�rzburg", "Würzburg"
        $restoredText = $restoredText -replace "D�sseldorf", "Düsseldorf"
        
        # Pattern-based replacements for common German character combinations
        # ß patterns (most � in German text are ß)
        $restoredText = $restoredText -replace "wei�", "weiß"      # Oberweißbach, weißt, etc.
        $restoredText = $restoredText -replace "gro�", "groß"      # groß, große, etc.
        $restoredText = $restoredText -replace "hei�", "heiß"      # heiß, heißt, etc.
        $restoredText = $restoredText -replace "wei�", "weiß"      # Already covered above
        $restoredText = $restoredText -replace "stra�", "straß"    # Straße, etc.
        $restoredText = $restoredText -replace "fu�", "fuß"        # Fuß, etc.
        
        # ü patterns 
        $restoredText = $restoredText -replace "�ringen", "üringen"  # Thüringen
        $restoredText = $restoredText -replace "�r", "ür"            # für, über, etc.
        $restoredText = $restoredText -replace "�n", "ün"            # grün, München, etc.
        
        # ä patterns
        $restoredText = $restoredText -replace "�r", "är"            # Bär, wär, etc. (if not caught by ü)
        $restoredText = $restoredText -replace "�h", "äh"            # näher, etc.
        
        # ö patterns  
        $restoredText = $restoredText -replace "�l", "öl"            # Köln, etc.
        $restoredText = $restoredText -replace "�n", "ön"            # schön, etc. (if not caught by ü)
        
        # Final fallback: if we still have � characters, assume they're ß (most common case)
        $restoredText = $restoredText -replace "�", "ß"
        
        if ($restoredText -ne $text) {
            Write-Host "Smart restoration: '$text' -> '$restoredText'" -ForegroundColor Green
            return $restoredText
        } else {
            Write-Host "No changes needed for: $text" -ForegroundColor Cyan
        }
    }
    
    # Fallback: Local replacement rules for common German encoding issues
    Write-Host "Using local encoding fixes for: $text" -ForegroundColor Cyan
    
    # Fix common German characters that appear as '�' due to encoding issues
    $text = $text -replace "Oberwei�bach", "Oberweißbach"
    $text = $text -replace "Th�ringenliga", "Thüringenliga"
    $text = $text -replace "Th�ringen", "Thüringen"
    $text = $text -replace "Reinhard-He�", "Reinhard-Heß"
    $text = $text -replace "Nordstra�e", "Nordstraße"
    $text = $text -replace "Wei�", "Weiß"
    $text = $text -replace "Stra�e", "Straße"
    $text = $text -replace "Gro�", "Groß"
    $text = $text -replace "F��e", "Füße"
    $text = $text -replace "M�nchen", "München"
    $text = $text -replace "N�rnberg", "Nürnberg"
    $text = $text -replace "W�rzburg", "Würzburg"
    
    # Fix common patterns where � appears
    $text = $text -replace "�", "ß"  # Most common case: ß becomes �
    
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
    if (![string]::IsNullOrWhiteSpace($GoogleApiKey)) {
        try {
            Write-Host "Calculating travel time from $origin to $destination using Google Maps..." -ForegroundColor Yellow
            
            # URL encode the addresses for the API
            $encodedOrigin = [System.Web.HttpUtility]::UrlEncode($origin)
            $encodedDestination = [System.Web.HttpUtility]::UrlEncode($destination)
            
            # Construct Google Maps Distance Matrix API URL
            $apiUrl = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=$encodedOrigin&destinations=$encodedDestination&mode=driving&language=de&key=$GoogleApiKey"
            
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
