# SpielerplusImport

A PowerShell script that transforms German volleyball league schedule data from CSV format to Excel format compatible with SpielerPlus team management software.

## Overview

This tool is specifically designed for **1. VSV Jena II** volleyball team to convert official league schedule CSV files into the format required by SpielerPlus for team management and logistics planning.

## Features

### 🏐 Volleyball-Specific Processing
- Filters games involving "1. VSV Jena II" team
- Automatically detects home vs. away games
- Calculates realistic meeting times based on travel distances

### 🌍 AI-Powered German Data Handling
- **OpenAI API Integration**: Uses gpt-4o-mini model for intelligent German character restoration
- **Smart Caching System**: Prevents duplicate API calls for repeated text corrections
- **Fallback Protection**: Local pattern matching when API is unavailable
- Processes semicolon-delimited CSV files
- Handles German date formats (dd.MM.yyyy)

### ⏰ Intelligent Time Management
- Converts "00:00:00" times to default 11:00:00 start time
- Calculates meeting times:
  - **Home games**: 2 hours before start time
  - **Away games**: Travel time + 60-minute buffer
- Formats all times in HH:mm:ss format

### 🗺️ Real-Time Travel Calculations
**Google Maps Distance Matrix API** integration for precise travel times:
- Real-time traffic-aware calculations
- Automatic postal code extraction from venue addresses
- Intelligent meeting time adjustments based on actual distances

**Built-in fallback estimates** from Jena to common venues:
- Erfurt: 60 minutes
- Weimar: 45 minutes  
- Gera: 45 minutes
- Meiningen: 90 minutes
- Suhl: 75 minutes
- Altenburg: 60 minutes
- Bleicherode: 120 minutes
- Oberweißbach: 60 minutes

### 📊 Excel Output Features
- SpielerPlus-compatible column mapping
- Automatic deadline calculations (7 and 14 days before games)
- Professional Excel formatting with filters and frozen headers
- Fallback to CSV if Excel export fails

## Requirements

- **PowerShell 5.1** or higher
- **ImportExcel module** (automatically installed if missing)
- Input CSV file matching pattern `*Spielplan*.csv`

### Optional API Keys
- **OpenAI API Key**: For AI-powered German character restoration (recommended)
- **Google Maps API Key**: For real-time travel time calculations (optional)

## Installation

1. Clone this repository:
```powershell
git clone https://github.com/darvader/SpielerplusImport.git
cd SpielerplusImport
```

2. The script will automatically install the required ImportExcel module on first run.

3. **Set up API keys** (optional but recommended):
   - Copy `.env.example` to `.env`
   - Add your OpenAI API key for intelligent German text correction
   - Add your Google Maps API key for real-time travel calculations
   - See [API Setup Guide](#api-setup) below for detailed instructions

## Usage

### Basic Usage
```powershell
.\Transform-Spielplan-Simple.ps1
```

### Custom Output Path
```powershell
.\Transform-Spielplan-Simple.ps1 -OutputPath ".\MySchedule.xlsx"
```

## Input Format

The script expects a CSV file with German volleyball league data containing these columns:
- `Datum` - Game date (dd.MM.yyyy format)
- `Uhrzeit` - Game time (HH:mm:ss format)
- `Mannschaft 1` / `Mannschaft 2` - Team names
- `Austragungsort` - Venue location
- Additional columns for referee, league info, etc.

## Output Format

The Excel output contains columns compatible with SpielerPlus:

| Column | Description | Example |
|--------|-------------|---------|
| Spieltyp (Opptional) | Game type | "Spiel" |
| Gegner | Opponent team | "Geraer VC I" |
| Start-Datum | Game date | 20.09.2025 |
| Start-Zeit | Start time | 11:00:00 |
| End-Zeit (Optional) | End time | 19:00:00 |
| Treffen (Optional) | Meeting time | 09:00:00 |
| Heimspiel | Home game flag | TRUE/FALSE |
| Gelände / Räumlichkeiten | Venue | "SH Lobdeburgschule (07747 Jena)" |
| Zu-/Absagen bis | Response deadline | 168 |
| Erinnerung zum Zu-/Absagen | Reminder hours | 336 |

## File Structure

```
SpielerplusImport/
├── Transform-Spielplan-Simple.ps1    # Main transformation script
├── Spielplan_Thüringenliga_Damen.csv # Sample input file
├── .env.example                      # Environment configuration template
├── .gitignore                        # Git ignore rules
├── README.md                         # This file
└── .github/
    └── copilot-instructions.md       # GitHub Copilot context
```

## Examples

### Sample Input (CSV)
```csv
"Datum";"Uhrzeit";"Mannschaft 1";"Mannschaft 2";"Austragungsort"
"20.09.2025";"11:00:00";"1. VSV Jena II";"Geraer VC I";"SH Lobdeburgschule (07747 Jena)"
"27.09.2025";"00:00:00";"VV70 Meiningen I";"1. VSV Jena II";"Multihalle Meiningen (98617 Meiningen)"
```

### Sample Output (Excel)
| Gegner | Start-Datum | Start-Zeit | Treffen | Heimspiel |
|--------|-------------|------------|---------|-----------|
| Geraer VC I | 20.09.2025 | 11:00:00 | 09:00:00 | TRUE |
| VV70 Meiningen I | 27.09.2025 | 11:00:00 | 08:30:00 | FALSE |

## Configuration

### Environment Settings
Create a `.env` file (copy from `.env.example`) to configure:

```properties
# OpenAI API for German character restoration (recommended)
OPENAI_API_KEY=your_openai_api_key_here

# Google Maps API for real-time travel calculations (optional)
GOOGLE_API_KEY=your_google_api_key_here

# Team Settings
HOME_TEAM_NAME=1. VSV Jena II
HOME_TEAM_VENUE=SH Lobdeburgschule (07747 Jena)

# Deadline Configuration (hours before game)
RESPONSE_DEADLINE_HOURS=168  # 7 days for response deadline
REMINDER_HOURS=336           # 14 days for reminder notification
```

### Available Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `OPENAI_API_KEY` | (empty) | API key for AI-powered German character restoration |
| `GOOGLE_API_KEY` | (empty) | API key for real-time travel calculations |
| `HOME_TEAM_NAME` | "1. VSV Jena II" | Name of your team in the CSV |
| `HOME_TEAM_VENUE` | "SH Lobdeburgschule (07747 Jena)" | Your home venue |
| `RESPONSE_DEADLINE_HOURS` | 168 | Hours before game for final response (7 days) |
| `REMINDER_HOURS` | 336 | Hours before game for reminder (14 days) |

## API Setup

### OpenAI API (Recommended)
For intelligent German character restoration:

1. **Create OpenAI Account**: Visit [platform.openai.com](https://platform.openai.com)
2. **Generate API Key**: Go to API Keys section and create a new key
3. **Add to .env file**: `OPENAI_API_KEY=your_api_key_here`
4. **Benefits**: 
   - Intelligent context-aware German character fixes
   - Handles complex encoding issues automatically
   - Faster than manual pattern matching
   - Comprehensive caching prevents duplicate API calls

### Google Maps API (Optional)
For real-time travel calculations:

1. **Create Google Cloud Project**: Visit [console.cloud.google.com](https://console.cloud.google.com)
2. **Enable Distance Matrix API**: In APIs & Services
3. **Create API Key**: In Credentials section
4. **Add to .env file**: `GOOGLE_API_KEY=your_api_key_here`
5. **Benefits**:
   - Real-time traffic-aware travel times
   - Automatic venue address parsing
   - More accurate meeting time calculations

### Google Maps Integration
For real-time travel calculations, see [GOOGLE_MAPS_SETUP.md](GOOGLE_MAPS_SETUP.md)

### Adding New Cities
To add travel times for new venues, modify the `Get-TravelTime` function:

```powershell
if ($venue -like "*NewCity*") { return 75 }  # 75 minutes to NewCity
```

### Adjusting Meeting Times
- **Home games**: Modify line with `$gameTime.AddHours(-2)`
- **Away games**: Modify `$travelMinutes + 60` for different buffer time

## Advanced Features

### 🤖 AI-Powered Text Correction
- **OpenAI gpt-4o-mini Model**: Latest, fastest, and most cost-effective model
- **Intelligent Context**: Understands German sports terminology and locations
- **Smart Caching**: Remembers corrections to avoid duplicate API calls
- **Rate Limiting**: Respects API limits with intelligent throttling
- **Fallback System**: Works without API using local pattern matching

### 📊 Performance Optimizations
- **Global Caching**: Prevents repeated API calls for same text
- **Batch Processing**: Efficiently processes large schedule files
- **Memory Management**: Handles large datasets without performance issues
- **Error Recovery**: Continues processing even if individual records fail

### 🔒 Security & Configuration
- **Environment Files**: Secure API key storage in `.env` files
- **Git Ignore**: Prevents accidental API key commits
- **Fallback Modes**: Works without any API keys using local data
- **Configurable Settings**: All deadlines and team settings customizable

## Troubleshooting

### Common Issues

**Q: Script can't find CSV file**
A: Ensure your CSV file name contains "Spielplan" and is in the same directory

**Q: German characters appear as question marks**
A: The script uses OpenAI API for intelligent character restoration. If you don't have an API key, it falls back to local pattern matching. For new characters, add them to the `Fix-GermanEncoding` function

**Q: Excel file shows dates instead of hours for deadlines**
A: The script uses text formatting to prevent this. Ensure you're using a recent version of the ImportExcel module

**Q: Meeting times seem incorrect**
A: Check that the venue name matches patterns in `Get-TravelTimeFromJena` function or ensure Google Maps API is configured for real-time calculations

**Q: OpenAI API rate limits reached**
A: The script automatically handles rate limiting. For heavy usage, the gpt-4o-mini model provides much higher limits than free tier models

**Q: Google Maps API returns REQUEST_DENIED**
A: Ensure Distance Matrix API is enabled in Google Cloud Console and billing is set up for your project

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Created for **1. VSV Jena II** volleyball team
- Built for **Thüringenliga Damen** schedule processing
- Designed to work with **SpielerPlus** team management software

---

*Made with ❤️ for volleyball team logistics in Thuringia, Germany*
