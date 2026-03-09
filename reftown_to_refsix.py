"""
Reftown to Refsix Assignment Converter

Reads the latest Reftown download and converts it to Refsix upload format.
"""

from pathlib import Path
import pandas as pd
from datetime import datetime


def get_latest_reftown_file(download_folder: Path) -> Path:
    """Find the latest Excel file in the download folder."""
    files = list(download_folder.glob("*.xlsx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No Excel files found in {download_folder}")
    latest_file = max(files, key=lambda f: f.stat().st_mtime)
    print(f"Latest Reftown file: {latest_file.name}")
    return latest_file


def load_conversion_tables(conversion_file: Path) -> dict:
    """Load lookup tables from the conversion template."""
    xl = pd.ExcelFile(conversion_file)

    # Read Tables_for_vlookups sheet without header to preserve structure
    df_lookup = pd.read_excel(conversion_file, sheet_name="Tables_for_vlookups",
                              header=None)

    # Extract age group lookup table (rows 7-27 approximately)
    # Look for the row starting with "Age group"
    age_start = None
    for idx, row in df_lookup.iterrows():
        if str(row.iloc[0]).strip() == "Age group (steers vlookup)":
            age_start = idx + 1  # Data starts after header
            break

    if age_start:
        age_data = df_lookup.iloc[age_start:age_start + 25].copy()
        age_data.columns = [f"col_{i}" for i in range(len(age_data.columns))]
        # Filter to valid age group rows (first column has U11, U12, etc.)
        age_data = age_data[age_data["col_0"].notna()]
        age_data = age_data[age_data["col_0"].str.match(r"^(U\d+|\d+[BG])", na=False)]

        # Create lookup dictionary
        age_lookup = {}
        for _, row in age_data.iterrows():
            age_group = str(row["col_0"]).strip()
            age_lookup[age_group] = {
                "team_size": int(row["col_3"]) if pd.notna(row["col_3"]) else 11,
                "bench_size": int(row["col_4"]) if pd.notna(row["col_4"]) else 7,
                "no_of_periods": int(row["col_5"]) if pd.notna(row["col_5"]) else 2,
                "period_length": int(row["col_6"]) if pd.notna(row["col_6"]) else 35,
                "interval_length": int(row["col_7"]) if pd.notna(row["col_7"]) else 10,
                "notes": str(row["col_2"]).strip() if pd.notna(row["col_2"]) else "",
            }
    else:
        age_lookup = {}
        print("Warning: Could not find age group lookup table")

    # Extract team color lookup table (rows 38-41 approximately)
    # Look for team abbreviations
    color_lookup = {}
    for idx in range(30, 50):
        if idx >= len(df_lookup):
            break
        row = df_lookup.iloc[idx]
        team_abbr = str(row.iloc[0]).strip()
        hex_color = str(row.iloc[1]).strip()

        # Check if it's a valid hex color pattern
        if team_abbr and hex_color.startswith("#") and len(hex_color) in [7, 4]:
            color_lookup[team_abbr] = hex_color

    print(f"Loaded age lookup for {len(age_lookup)} age groups")
    print(f"Loaded colors for {len(color_lookup)} teams")

    return {"age_groups": age_lookup, "team_colors": color_lookup}


def parse_reftown_date(date_str):
    """Convert MM/DD/YYYY to YYYY-MM-DD."""
    if pd.isna(date_str):
        return ""
    date_str = str(date_str).strip()
    try:
        # Handle Excel date serial
        if isinstance(date_str, (int, float)) or date_str.isdigit():
            dt = pd.to_datetime(float(date_str), origin="1899-12-30", unit="D")
        else:
            dt = pd.to_datetime(date_str)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return ""


def parse_reftown_time(time_str):
    """Convert time to HH:MM 24-hour format."""
    if pd.isna(time_str):
        return ""
    time_str = str(time_str).strip()

    try:
        # Handle Excel time serial
        if isinstance(time_str, (int, float)):
            # Excel stores time as fraction of day
            dt = pd.to_datetime(float(time_str), unit="D", origin="1899-12-30")
            return dt.strftime("%H:%M")
        else:
            # Parse "3:00 PM" format
            dt = pd.to_datetime(time_str)
            return dt.strftime("%H:%M")
    except Exception:
        return ""


def extract_team_short_name(team_name: str) -> str:
    """Extract short name from full team name."""
    if pd.isna(team_name):
        return ""
    team_name = str(team_name).strip()

    # Try to extract abbreviation (e.g., "PCU 11G Red 3" -> "PCU")
    parts = team_name.split()
    if parts:
        # First 3-4 characters of first part, or the whole abbreviation
        abbrev = parts[0]
        if len(abbrev) <= 4:
            return abbrev.upper()
        else:
            return abbrev[:3].upper()

    return team_name[:3].upper()


def get_age_group(level_1: str) -> str:
    """Extract age group from Level_1 field (e.g., 'U15' from 'U15 Girls')."""
    if pd.isna(level_1):
        return ""
    level_1 = str(level_1).strip()

    # Look for U## pattern
    import re
    match = re.search(r"U\d{2}", level_1)
    if match:
        return match.group(0)

    # Look for ##B or ##G pattern (e.g., "15B")
    match = re.search(r"\d{2}[BG]", level_1)
    if match:
        age = match.group(0)[:2]
        return f"U{age}"

    return ""


def convert_reftown_to_refsix(reftown_df: pd.DataFrame,
                               lookup_tables: dict,
                               my_name: str = "") -> pd.DataFrame:
    """
    Convert Reftown data to Refsix format.

    Args:
        reftown_df: DataFrame with Reftown assignment data
        lookup_tables: Dictionary with age_groups and team_colors
        my_name: Optional official name to identify their role
    """

    output_rows = []

    for _, row in reftown_df.iterrows():
        age_group = get_age_group(row.get("Level_1", ""))
        age_data = lookup_tables["age_groups"].get(age_group, {})

        # Parse date and time
        date_str = parse_reftown_date(row.get("Date"))
        time_str = parse_reftown_time(row.get("Time"))

        # Extract team names
        home_team = str(row.get("Home", "")).strip()
        away_team = str(row.get("Visitor_1", "")).strip()

        home_short = extract_team_short_name(home_team)
        away_short = extract_team_short_name(away_team)

        # Lookup team colors
        home_color = lookup_tables["team_colors"].get(home_short, "#FFFFFF")
        away_color = lookup_tables["team_colors"].get(away_short, "#FFFFFF")

        # Determine official roles - populate all officials from Reftown
        officials = {
            "Referee": "",
            "Assistant Referee": "",
            "Assistant Referee2": "",
            "4th Official": ""
        }

        # Track user's role for the Official Role column (if my_name provided)
        my_official_role = ""

        # Map Reftown official columns to Refsix roles
        official_mapping = {
            "Official": "Referee",
            "Official.1": "Assistant Referee",
            "Official.2": "Assistant Referee2",
            "Official.3": "4th Official"
        }

        for reftown_col, refsix_role in official_mapping.items():
            official_name = str(row.get(reftown_col, "")).strip()
            if official_name and official_name != "nan":
                officials[refsix_role] = official_name
                # Check if this is the user's role (if my_name provided)
                if my_name and my_name.lower() in official_name.lower():
                    my_official_role = refsix_role

        # Build output row
        output_row = {
            "Match ID": row.get("GameID", ""),
            "Competition": row.get("League", ""),
            "Venue": row.get("Location", ""),
            "Date (YYYY-MM-DD)": date_str,
            "Time (HH:MM)": time_str,
            "Time Zone": "",
            "Official Role": my_official_role,
            "Home Team Name": home_team,
            "Home Team Short Name": home_short,
            "Home Team Colour": home_color,
            "Away Team Name": away_team,
            "Away Team Short Name": away_short,
            "Away Team Colour": away_color,
            "Team Size": age_data.get("team_size", 11),
            "Bench Size": age_data.get("bench_size", 7),
            "No of Periods": age_data.get("no_of_periods", 2),
            "Period Length": age_data.get("period_length", 35),
            "Interval Length": age_data.get("interval_length", 10),
            "Extra Time": "Yes",
            "Extra Time Length": 10,
            "Penalty Kicks": "Yes",
            "Record Goal Scorers": "No",
            "Misconduct Codes": "USSF",
            "Temporary Dismissals": "No",
            "Dismissal Length": 10,
            "Referee": officials["Referee"],
            "Assistant Referee": officials["Assistant Referee"],
            "Assistant Referee2": officials["Assistant Referee2"],
            "4th Official": officials["4th Official"],
            "Observer": 0,
            "Fees": 0,
            "Expenses": 0,
            "Notes": age_data.get("notes", ""),
            "Tag 1 ": age_data.get("notes", "") if "PK" in age_data.get("notes", "") else "",
            "Tag 2": row.get("Organization", ""),
            "Tag 3": f"{age_data.get('period_length', 35)}min",
            "Tag 4": row.get("CrewType", ""),
            "Tag 5": age_group,
        }

        output_rows.append(output_row)

    # Create DataFrame and replace NaN with empty strings to avoid Excel issues
    df = pd.DataFrame(output_rows)
    df = df.fillna("")
    return df


def main():
    """Main conversion function."""
    # Define paths
    base_dir = Path.cwd()
    download_folder = base_dir / "1 Download_from_Reftown"
    conversion_file = base_dir / "2 Conversion_file" / "REFSIXUploadMatchesTemplate.xlsx"
    output_folder = base_dir / "3 Upload_to_Refsix"

    # Ensure output folder exists
    output_folder.mkdir(exist_ok=True)

    # Get latest Reftown file
    reftown_file = get_latest_reftown_file(download_folder)

    # Load Reftown data
    print(f"Reading Reftown data from: {reftown_file.name}")
    reftown_df = pd.read_excel(reftown_file)
    print(f"Found {len(reftown_df)} assignments")

    # Load conversion lookup tables
    print(f"Loading conversion tables from: {conversion_file.name}")
    lookup_tables = load_conversion_tables(conversion_file)

    # Convert data
    print("Converting data to Refsix format...")
    refsix_df = convert_reftown_to_refsix(reftown_df, lookup_tables)

    # Generate output filename with today's date
    today = datetime.now().strftime("%d%b%Y")
    output_filename = f"Export_assignments_toRefsix_{today}.xlsx"
    output_path = output_folder / output_filename

    # Write output
    print(f"Writing output to: {output_filename}")
    refsix_df.to_excel(output_path, index=False)

    print(f"\nConversion complete!")
    print(f"- Processed {len(refsix_df)} assignments")
    print(f"- Output saved to: {output_path}")

    # Show sample of converted data
    print(f"\nSample of converted data:")
    print(refsix_df.head(3).to_string())

    return refsix_df


if __name__ == "__main__":
    main()
