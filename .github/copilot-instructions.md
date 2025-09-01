# GitHub Copilot Instructions for SpielerplusImport

## Project Overview
This project transforms volleyball schedule data from CSV format (German volleyball league) to Excel format compatible with SpielerPlus team management software.

## Key Context
- **Sport**: Volleyball team schedule management for "1. VSV Jena II"
- **Source Format**: CSV with German text encoding and semicolon delimiters
- **Target Format**: Excel with specific column mapping for SpielerPlus import
- **Language**: PowerShell with ImportExcel module
- **Geography**: German cities in Thuringia region (travel time calculations from Jena)

## Technical Stack
- **PowerShell**: Main scripting language
- **ImportExcel Module**: For Excel file generation and formatting
- **OpenAI API**: gpt-4o-mini model for intelligent German character restoration
- **Google Maps API**: Distance Matrix API for real-time travel calculations
- **CSV Processing**: Manual parsing with AI-powered encoding fixes
- **Caching System**: Global cache to prevent duplicate API calls
- **Environment Configuration**: Secure API key management via .env files

## Code Patterns & Conventions

### AI-Powered German Character Handling
```powershell
# OpenAI API integration with caching
function Fix-GermanEncoding {
    param([string]$text)
    
    # Check cache first
    if ($global:germanTextCache.ContainsKey($text)) {
        return $global:germanTextCache[$text]
    }
    
    # Use OpenAI API with rate limiting
    if ($openAIApiKey) {
        # Rate limiting logic
        # API call to gpt-4o-mini model
        # Cache result
    } else {
        # Fallback to local pattern matching
        $text = $text -replace "Oberwei�bach", "Oberweißbach"
        $text = $text -replace "Th�ringenliga", "Thüringenliga"
    }
}
```

### Google Maps Integration
```powershell
# Real-time travel time calculation
function Get-GoogleMapsDistance {
    param([string]$origin, [string]$destination)
    
    # Extract postal codes, make API call
    # Return actual travel time with traffic
}
```

### Time Formatting
```powershell
# Always use HH:mm:ss format for time fields
$gameTime.ToString("HH:mm:ss")
# Handle "00:00:00" as default 11:00:00 start time
$timeToUse = if ($uhrzeit -eq "00:00:00") { "11:00:00" } else { $uhrzeit }
```

### Travel Time Logic
```powershell
# Home games: 2 hours early
$treffenTime = ($gameTime.AddHours(-2)).ToString("HH:mm:ss")
# Away games: travel time + 60 minutes buffer
$totalMinutesEarly = $travelMinutes + 60
```

## Data Structure Mapping

### Input CSV Columns (German)
- Datum, Uhrzeit, Mannschaft 1, Mannschaft 2, Austragungsort, etc.

### Output Excel Columns (SpielerPlus format)
- 'Spieltyp (Opptional)', 'Gegner', 'Start-Datum', 'Start-Zeit', 'Heimspiel', etc.

## Business Rules
1. **Team Filter**: Only process games involving "1. VSV Jena II"
2. **Home Game Logic**: Home game when "1. VSV Jena II" is "Mannschaft 1"
3. **Time Defaults**: Use 11:00:00 when original time is 00:00:00
4. **Meeting Times**: 
   - Home: 2 hours before game
   - Away: Travel time + 60 min buffer before game
5. **Deadlines**: 
   - Zu-/Absagen: 168 hours (7 days) before
   - Erinnerung: 336 hours (14 days) before

## Travel Time Reference (from Jena)
- Erfurt: 60 minutes
- Weimar: 45 minutes
- Gera: 45 minutes
- Meiningen: 90 minutes
- Suhl: 75 minutes
- Altenburg: 60 minutes
- Bleicherode: 120 minutes
- Oberweißbach: 60 minutes
- Default: 90 minutes

## File Handling
- **Input**: `*Spielplan*.csv` (UTF-8 encoding)
- **Output**: `Transformed_Spielplan_ExcelFormat.xlsx`
- **Environment**: `.env` file for API keys and configuration
- **Cache**: Global variables for API response caching
- **Gitignore**: Excel temp files (~$*.xlsx), generated outputs, and environment files

## Error Handling Patterns
```powershell
try {
    # Date/time parsing with German format
    $gameDate = [DateTime]::ParseExact($datum, "dd.MM.yyyy", $null)
} catch {
    Write-Warning "Could not parse date/time for row: $nummer"
    continue
}
```

## Excel Export Considerations
- Use text format (@) for hour fields to prevent date conversion
- Apply AutoSize, AutoFilter, FreezeTopRow, BoldTopRow
- Fallback to CSV export if Excel fails

## Testing Approach
- Verify German character encoding fixes
- Check home/away game detection
- Validate time calculations and formatting
- Ensure Excel columns match SpielerPlus requirements
- Test with real volleyball schedule data

## Common Tasks
When working on this project, you might need to:
- Add new cities to travel time calculations
- Modify meeting time logic for different game types
- Update German character encoding mappings
- Adjust Excel column formatting
- Handle new CSV format variations
- Configure OpenAI API integration and caching
- Set up Google Maps API for real-time travel calculations
- Optimize API rate limiting and error handling
- Update environment configuration settings
