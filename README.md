# Reftown-to-Refsix

A Python utility to convert soccer officiating assignments from Reftown platform to Refsix system format.

## Purpose

This tool automates the conversion of officiating assignment data between two soccer officiating management platforms:
- **Source**: Reftown (download format)
- **Target**: Refsix (upload format)

Perfect for soccer officials and assignors who need to transfer data between these systems.

## Features

- Automatic detection of the latest Reftown download file
- Age group mapping with configurable game parameters
- Team name extraction and abbreviation
- Team color assignment using hex codes
- Official role mapping (Referee, AR1, AR2, 4th Official)
- Date/time format conversion to ISO 8601 standard
- Template-based conversion using Excel lookup tables

## Requirements

- Python 3.8 or higher
- pandas
- openpyxl

## Installation

1. Clone this repository:
```bash
git clone https://github.com/veikkoautoja-cyber/Reftown-to-Refsix.git
cd Reftown-to-Refsix
```

2. Create a virtual environment (recommended):
```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Setup

Create the required folder structure:

```
Reftown-to-Refsix/
├── 1 Download_from_Reftown/      # Place your Reftown downloads here
├── 2 Conversion_file/             # Contains template with lookup tables
│   ├── REFSIXUploadMatchesTemplate.xlsx
│   └── New_colors.xlsx
└── 3 Upload_to_Refsix/            # Converted files will be saved here
```

### Folders Explained

1. **`1 Download_from_Reftown/`**
   - Place your Excel files downloaded from Reftown here
   - The tool automatically finds the most recent file

2. **`2 Conversion_file/`**
   - Contains the Refsix template with lookup tables
   - `REFSIXUploadMatchesTemplate.xlsx` - Main template with age group mappings
   - `New_colors.xlsx` - Team color assignments
   - These files are included in the repository

3. **`3 Upload_to_Refsix/`**
   - Output folder for converted files
   - Files are named with date stamp: `Export_assignments_toRefsix_DDMMMYYYY.xlsx`
   - Created automatically if it doesn't exist

## Usage

### Basic Usage

Run the conversion from the command line:

```bash
python main.py
```

The tool will:
1. Find the latest Reftown download file
2. Load conversion lookup tables
3. Convert data to Refsix format
4. Save output to `3 Upload_to_Refsix/` folder
5. Display a sample of converted data

### Specify Your Name

To identify your official role in the assignments, use the `--name` parameter:

```bash
python main.py --name "Your Name"
```

This will populate the "Official Role" column with your assigned role (Referee, Assistant Referee, etc.) for each game.

### Customization

You can customize the conversion by editing the lookup tables in `2 Conversion_file/REFSIXUploadMatchesTemplate.xlsx`:

- **Age Group Settings**: Team size, bench size, periods, period length, interval
- **Team Colors**: Hex color codes for each team abbreviation

## Example Output

The converted file contains columns for:
- Match details (ID, competition, venue, date, time)
- Team information (names, abbreviations, colors)
- Game parameters (team size, periods, lengths)
- Official assignments (Referee, AR1, AR2, 4th Official)
- Tags and notes

## File Format Reference

### Input: Reftown Download Format
Expected columns in the Reftown Excel file:
- `GameID` - Unique game identifier
- `Date` - Game date (MM/DD/YYYY or Excel serial)
- `Time` - Game time (HH:MM AM/PM or Excel serial)
- `Level_1` - Age group (e.g., "U15 Girls")
- `Home` - Home team name
- `Visitor_1` - Away team name
- `Location` - Venue name
- `League` - Competition name
- `Official`, `Official.1`, `Official.2`, `Official.3` - Official names
- Additional metadata columns

### Output: Refsix Upload Format
Standard Refsix import format with all required columns for match and official data.

## Troubleshooting

**No Excel files found in download folder**
- Ensure your Reftown downloads are in `1 Download_from_Reftown/`
- Check that files have `.xlsx` extension
- Avoid temporary Excel files starting with `~$`

**Age group not found**
- Edit the lookup table in `REFSIXUploadMatchesTemplate.xlsx`
- Add missing age groups to the "Tables_for_vlookups" sheet

**Incorrect team colors**
- Update color mappings in `New_colors.xlsx` or the template

## Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest features
- Submit pull requests

## License

MIT License - See LICENSE file for details

## Acknowledgments

Designed to streamline officiating workflow between Reftown and Refsix platforms.
