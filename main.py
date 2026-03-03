"""
Reftown to Refsix Assignment Converter

Main entry point with command-line interface.
"""

import argparse
from pathlib import Path
from reftown_to_refsix import (
    get_latest_reftown_file,
    load_conversion_tables,
    convert_reftown_to_refsix,
)
import pandas as pd
from datetime import datetime


def parse_arguments():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="Convert Reftown assignments to Refsix format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py                           # Use default paths
  python main.py --name "John Doe"         # Specify official name
  python main.py --input my_data.xlsx      # Use specific input file
  python main.py --output custom.xlsx      # Specify output filename
        """
    )

    parser.add_argument(
        "--name",
        default="",
        help="Your official name (optional, for identifying your role in assignments)"
    )

    parser.add_argument(
        "--input",
        type=Path,
        help="Specific input Reftown file (default: latest in 1 Download_from_Reftown/)"
    )

    parser.add_argument(
        "--output",
        type=str,
        help="Output filename (default: Export_assignments_toRefsix_DDMMMYYYY.xlsx)"
    )

    parser.add_argument(
        "--download-folder",
        type=Path,
        default=Path.cwd() / "1 Download_from_Reftown",
        help="Path to Reftown downloads folder"
    )

    parser.add_argument(
        "--conversion-file",
        type=Path,
        default=Path.cwd() / "2 Conversion_file" / "REFSIXUploadMatchesTemplate.xlsx",
        help="Path to conversion template file"
    )

    parser.add_argument(
        "--output-folder",
        type=Path,
        default=Path.cwd() / "3 Upload_to_Refsix",
        help="Path to output folder"
    )

    return parser.parse_args()


def main():
    """Main conversion function with CLI support."""
    args = parse_arguments()

    print("=" * 60)
    print("Reftown to Refsix Assignment Converter")
    print("=" * 60)
    print()

    # Validate paths
    if not args.download_folder.exists():
        print(f"Error: Download folder not found: {args.download_folder}")
        print("Please create the folder and add Reftown downloads.")
        return

    if not args.conversion_file.exists():
        print(f"Error: Conversion file not found: {args.conversion_file}")
        print("Please ensure the template file exists.")
        return

    # Ensure output folder exists
    args.output_folder.mkdir(parents=True, exist_ok=True)

    # Get input file
    if args.input:
        reftown_file = args.input
        if not reftown_file.exists():
            print(f"Error: Input file not found: {reftown_file}")
            return
    else:
        try:
            reftown_file = get_latest_reftown_file(args.download_folder)
        except FileNotFoundError as e:
            print(f"Error: {e}")
            print("\nPlease place your Reftown downloads in the download folder.")
            return

    print(f"Input file: {reftown_file.name}")
    print(f"Official name: {args.name}")
    print()

    # Load Reftown data
    print("Reading Reftown data...")
    reftown_df = pd.read_excel(reftown_file)
    print(f"Found {len(reftown_df)} assignments")
    print()

    # Load conversion lookup tables
    print("Loading conversion tables...")
    lookup_tables = load_conversion_tables(args.conversion_file)
    print()

    # Convert data
    print("Converting data to Refsix format...")
    refsix_df = convert_reftown_to_refsix(reftown_df, lookup_tables, my_name=args.name)

    # Generate output filename
    if args.output:
        output_filename = args.output
    else:
        today = datetime.now().strftime("%d%b%Y")
        output_filename = f"Export_assignments_toRefsix_{today}.xlsx"

    output_path = args.output_folder / output_filename

    # Write output
    print(f"Writing output to: {output_filename}")
    refsix_df.to_excel(output_path, index=False)

    print()
    print("=" * 60)
    print("Conversion complete!")
    print("=" * 60)
    print(f"  Processed: {len(refsix_df)} assignments")
    print(f"  Output:    {output_path}")
    print()

    # Show sample of converted data
    print("Sample of converted data:")
    print("-" * 60)
    display_cols = [
        "Date (YYYY-MM-DD)",
        "Time (HH:MM)",
        "Home Team Short Name",
        "Away Team Short Name",
        "Official Role",
    ]
    available_cols = [col for col in display_cols if col in refsix_df.columns]
    print(refsix_df[available_cols].head(5).to_string(index=False))
    print("-" * 60)

    return refsix_df


if __name__ == "__main__":
    main()
