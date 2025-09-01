# SpielerplusImport

A PowerShell script that transforms German volleyball league schedule data from CSV format to Excel format compatible with SpielerPlus team management software.

## Overview

This tool is specifically designed for **1. VSV Jena II** volleyball team to convert official league schedule CSV files into the format required by SpielerPlus for team management and logistics planning.

## Features

### 🏐 Volleyball-Specific Processing
- Filters games involving "1. VSV Jena II" team
- Automatically detects home vs. away games
- Calculates realistic meeting times based on travel distances

### 🌍 German Data Handling
- Fixes encoding issues with German umlauts (ä, ö, ü) and ß
- Processes semicolon-delimited CSV files
- Handles German date formats (dd.MM.yyyy)

### ⏰ Intelligent Time Management
- Converts "00:00:00" times to default 11:00:00 start time
- Calculates meeting times:
  - **Home games**: 2 hours before start time
  - **Away games**: Travel time + 60-minute buffer
- Formats all times in HH:mm:ss format

### 🗺️ Travel Time Estimation
Built-in travel time estimates from Jena to common venues:
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

## Installation

1. Clone this repository:
```powershell
git clone https://github.com/darvader/SpielerplusImport.git
cd SpielerplusImport
```

2. The script will automatically install the required ImportExcel module on first run.

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

### Adding New Cities
To add travel times for new venues, modify the `Get-TravelTimeFromJena` function:

```powershell
if ($venue -like "*NewCity*") { return 75 }  # 75 minutes to NewCity
```

### Adjusting Meeting Times
- **Home games**: Modify line with `$gameTime.AddHours(-2)`
- **Away games**: Modify `$travelMinutes + 60` for different buffer time

## Troubleshooting

### Common Issues

**Q: Script can't find CSV file**
A: Ensure your CSV file name contains "Spielplan" and is in the same directory

**Q: German characters appear as question marks**
A: The script automatically fixes common encoding issues. For new characters, add them to the `Fix-GermanEncoding` function

**Q: Excel file shows dates instead of hours for deadlines**
A: The script uses text formatting to prevent this. Ensure you're using a recent version of the ImportExcel module

**Q: Meeting times seem incorrect**
A: Check that the venue name matches patterns in `Get-TravelTimeFromJena` function

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
