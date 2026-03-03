# Example Reftown Download

Place your Reftown download files in the parent directory (`1 Download_from_Reftown/`).

## Expected Format

Your Reftown download should be an Excel file (.xlsx) with the following columns:

- **GameID** - Unique game identifier
- **Date** - Game date (MM/DD/YYYY or Excel serial date)
- **Time** - Game time (HH:MM AM/PM or Excel serial time)
- **Level_1** - Age group (e.g., "U15 Girls", "U12B")
- **Home** - Home team full name
- **Visitor_1** - Away team full name
- **Location** - Venue name
- **League** - Competition/league name
- **Official** - Referee name
- **Official.1** - Assistant Referee 1 name
- **Official.2** - Assistant Referee 2 name
- **Official.3** - 4th Official name
- **Organization** - Organization name (for Tag 2)
- **CrewType** - Crew type (for Tag 4)

## File Naming

The tool automatically finds the most recent file based on modification time.
