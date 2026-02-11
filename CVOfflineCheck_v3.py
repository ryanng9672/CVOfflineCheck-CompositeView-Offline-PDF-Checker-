import os
import pandas as pd
from datetime import datetime, timedelta
import argparse
import sys
def get_weekday_abbreviation(date):
    """Convert date to weekday abbreviation (Mon, Tue, Wed, Thu, Fri)."""
    weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    return weekdays[date.weekday()]
def get_expected_weekdays_and_dates(current_date):
    """
    Get expected weekdays and dates for reports.
    Returns tuple of (weekday_list, date_list) in reverse chronological order.
    """
    expected_weekdays = []
    expected_dates = []
    check_date = current_date
    while len(expected_weekdays) < 5:
        # Only include weekdays (Monday=0 to Friday=4)
        if check_date.weekday() < 5:
            expected_weekdays.append(get_weekday_abbreviation(check_date))
            expected_dates.append(check_date.strftime('%Y-%m-%d'))
        check_date -= timedelta(days=1)
    return expected_weekdays, expected_dates
def validate_report_date(df, expected_dates):
    """
    Validate if the report contains data from expected dates.
    Returns (is_valid, actual_date, error_message).
    """
    if df.empty:
        return False, None, "Report is empty"
    # Try to find date column (could be 'Date', 'date', 'DATE', etc.)
    date_column = None
    for col in df.columns:
        if 'date' in col.lower():
            date_column = col
            break
    if date_column is None:
        return False, None, "No date column found in report"
    # Get the most recent date from the report
    try:
        df[date_column] = pd.to_datetime(df[date_column])
        actual_date = df[date_column].max().strftime('%Y-%m-%d')
        # Check if the actual date matches any expected date
        if actual_date in expected_dates:
            return True, actual_date, None
        else:
            return False, actual_date, f"Report date {actual_date} does not match expected dates"
    except Exception as e:
        return False, None, f"Error parsing dates: {str(e)}"
def find_and_validate_report(diffreport_folder, report_type, expected_weekdays, expected_dates):
    """
    Find and validate a report file.
    Returns tuple of (dataframe, filename, date) or (None, None, None) if not found.
    """
    found_df = None
    found_file = None
    found_date = None
    print(f"  Searching in: {diffreport_folder}")
    try:
        files_in_folder = os.listdir(diffreport_folder)
        matching_files = [f for f in files_in_folder if f.startswith(report_type)]
        if matching_files:
            print(f"  Found {len(matching_files)} {report_type} files: {', '.join(matching_files)}")
        else:
            print(f"  ⚠ No {report_type} files found in folder!")
    except Exception as e:
        print(f"  ✗ Cannot access folder: {e}")
    for weekday in reversed(expected_weekdays):
        csv_file = os.path.join(diffreport_folder, f"{report_type}_Diff_{weekday}.csv")
        print(f"  Looking for: {report_type}_Diff_{weekday}.csv ... ", end="")
        if os.path.exists(csv_file):
            print("FOUND")
            df = pd.read_csv(csv_file)
            is_valid, actual_date, error_msg = validate_report_date(df, expected_dates)
            if is_valid:
                print(f"    ✓ Valid report with date: {actual_date}")
                return df, f"{report_type}_Diff_{weekday}.csv", actual_date
            else:
                print(f"    ✗ Invalid or outdated: {error_msg}")
        else:
            print("NOT FOUND")
    return None, None, None
def find_pdf(pdf_base_path, equipment_name):
    """
    Search for PDF file in nested folder structure.
    Returns the full path to the PDF if found, None otherwise.
    """
    equipment_name_clean = equipment_name.strip()
    # Search in all subdirectories
    for root, dirs, files in os.walk(pdf_base_path):
        for file in files:
            if file.lower().endswith('.pdf'):
                # Check if equipment name matches (case-insensitive)
                if equipment_name_clean.lower() in file.lower() or file.lower().replace('.pdf', '') == equipment_name_clean.lower():
                    return os.path.join(root, file)
    return None
def check_pdf_status(diffreport_folder, pdf_base_path, output_folder, output_filename):
    """
    Main function to check PDF status against DiffReport files.
    """
    print("\n" + "="*80)
    print("PDF STATUS CHECK - START")
    print("="*80)
    # Get current date and expected report information
    current_date = datetime.now()
    print(f"\nCurrent Date: {current_date.strftime('%Y-%m-%d')} ({get_weekday_abbreviation(current_date)})")
    expected_weekdays, expected_dates = get_expected_weekdays_and_dates(current_date)
    print(f"\nExpected Report Weekdays: {', '.join(expected_weekdays)}")
    print(f"Expected Report Dates: {', '.join(expected_dates)}")
    print("-" * 80)
    # Load and validate CompositeView report
    print("\n[1] Loading and validating CompositeView DiffReport...")
    cv_df, cv_file, cv_date = find_and_validate_report(
        diffreport_folder, 
        "CompositeView", 
        expected_weekdays, 
        expected_dates
    )
    # Load and validate Substation report
    print("\n[2] Loading and validating Substation DiffReport...")
    sub_df, sub_file, sub_date = find_and_validate_report(
        diffreport_folder, 
        "Substation", 
        expected_weekdays, 
        expected_dates
    )
    # Validation results
    print("\n" + "="*80)
    print("VALIDATION RESULTS:")
    print("-" * 80)
    validation_failed = False
    error_messages = []
    if cv_df is None:
        validation_failed = True
        error_messages.append(f"CompositeView: No CompositeView report found for weekdays: {', '.join(expected_weekdays)}")
    else:
        print(f"✓ CompositeView: Using {cv_file} (Date: {cv_date})")
    if sub_df is None:
        validation_failed = True
        error_messages.append(f"Substation: No Substation report found for weekdays: {', '.join(expected_weekdays)}")
    else:
        print(f"✓ Substation: Using {sub_file} (Date: {sub_date})")
    if validation_failed:
        print("\n✗ VALIDATION FAILED - Both reports are outdated or missing")
        for msg in error_messages:
            print(f"  ✗ {msg}")
        print("="*80)
        print("\n✗ PROGRAM STOPPED - Both DiffReport validations failed!")
        print("\n✗ Program stopped due to validation errors!")
        return None
    print("="*80)
    # Combine the two dataframes
    print("\n[3] Combining CompositeView and Substation data...")
    combined_df = pd.concat([cv_df, sub_df], ignore_index=True)
    print(f"  Total records: {len(combined_df)}")
    # Check for required columns
    required_columns = ['Equipment Name', 'Equipment Type']
    missing_columns = [col for col in required_columns if col not in combined_df.columns]
    if missing_columns:
        print(f"\n✗ Error: Missing required columns: {', '.join(missing_columns)}")
        print(f"  Available columns: {', '.join(combined_df.columns.tolist())}")
        return None
    # Filter for specific equipment types
    print("\n[4] Filtering for Circuit Breaker and Switch equipment...")
    filtered_df = combined_df[
        combined_df['Equipment Type'].isin(['Circuit Breaker', 'Switch'])
    ].copy()
    print(f"  Filtered records: {len(filtered_df)}")
    if filtered_df.empty:
        print("\n⚠ Warning: No Circuit Breaker or Switch equipment found in the reports")
        return None
    # Check PDF existence
    print("\n[5] Checking PDF file existence...")
    print(f"  Base search path: {pdf_base_path}")
    pdf_status = []
    total_equipment = len(filtered_df)
    for idx, row in filtered_df.iterrows():
        equipment_name = row['Equipment Name']
        equipment_type = row['Equipment Type']
        # Show progress
        current_num = idx + 1
        print(f"  [{current_num}/{total_equipment}] Checking: {equipment_name} ... ", end="")
        pdf_path = find_pdf(pdf_base_path, equipment_name)
        if pdf_path:
            print(f"FOUND")
            pdf_status.append({
                'Equipment Name': equipment_name,
                'Equipment Type': equipment_type,
                'PDF Status': 'Exists',
                'PDF Path': pdf_path
            })
        else:
            print(f"NOT FOUND")
            pdf_status.append({
                'Equipment Name': equipment_name,
                'Equipment Type': equipment_type,
                'PDF Status': 'Missing',
                'PDF Path': ''
            })
    # Create result dataframe
    result_df = pd.DataFrame(pdf_status)
    # Generate summary statistics
    total_count = len(result_df)
    exists_count = len(result_df[result_df['PDF Status'] == 'Exists'])
    missing_count = len(result_df[result_df['PDF Status'] == 'Missing'])
    print("\n" + "="*80)
    print("SUMMARY:")
    print("-" * 80)
    print(f"Total Equipment Checked: {total_count}")
    print(f"  ✓ PDFs Found:    {exists_count} ({exists_count/total_count*100:.1f}%)")
    print(f"  ✗ PDFs Missing:  {missing_count} ({missing_count/total_count*100:.1f}%)")
    print("="*80)
    # Save to CSV
    output_path = os.path.join(output_folder, output_filename)
    result_df.to_csv(output_path, index=False)
    print(f"\n✓ Results saved to: {output_path}")
    # Display missing PDFs if any
    if missing_count > 0:
        print("\n⚠ Missing PDFs:")
        missing_df = result_df[result_df['PDF Status'] == 'Missing']
        for idx, row in missing_df.iterrows():
            print(f"  - {row['Equipment Name']} ({row['Equipment Type']})")
    return result_df
def interactive_path_input():
    """
    Interactive prompt for path configuration when running without arguments.
    Returns tuple: (diffreport_path, pdf_path, output_path) or None to exit.
    """
    print("\n" + "="*80)
    print("COMPOSITE VIEW OFFLINE CHECK - INTERACTIVE MODE")
    print("="*80)
    print("\nPlease choose how to configure paths:")
    print("-" * 80)
    print("1. Type 'Draft' or 'draft'  → Use default paths (admssim01)")
    print("2. Type '-'                 → Show command-line usage guide")
    print("3. Enter custom path        → Manually specify DiffReport folder path")
    print("4. Press ENTER only         → Use default paths")
    print("-" * 80)
    user_input = input("\nYour choice: ").strip()
    # Option 1 & 4: Use default paths
    if user_input.lower() == 'draft' or user_input == '':
        print("\n✓ Using default paths:")
        diffreport = r"\\admssim01\ADMS_DataEngineering\CompositeViewBackup\DiffReport"
        pdf_path = r"\\admssim01\ADMS_DataEngineering\DMS_Picture_offline"
        output = r"C:\ADMS_DataEngineering\DMS_Picture_offline"
        print(f"  DiffReport: {diffreport}")
        print(f"  PDF Base:   {pdf_path}")
        print(f"  Output:     {output}")
        return diffreport, pdf_path, output
    # Option 2: Show usage guide
    elif user_input == '-':
        print("\n" + "="*80)
        print("COMMAND-LINE USAGE GUIDE")
        print("="*80)
        print("\nTo run this program with custom paths from command line:")
        print("-" * 80)
        print("\nBasic usage:")
        print('  CVOfflineCheck_v2.exe --diffreport "YOUR_PATH" --output "YOUR_OUTPUT"\n')
        print("Available parameters:")
        print("  --diffreport PATH")
        print("      Path to DiffReport folder containing CSV files")
        print('      Example: --diffreport "\\\\admssim01\\ADMS_DataEngineering\\CompositeViewBackup\\DiffReport"')
        print()
        print("  --pdf-path PATH")
        print("      Base path to search for PDF files")
        print('      Example: --pdf-path "\\\\admssim01\\ADMS_DataEngineering\\DMS_Picture_offline"')
        print()
        print("  --output PATH")
        print("      Output folder for result CSV")
        print('      Example: --output "C:\\ADMS_DataEngineering\\DMS_Picture_offline"')
        print()
        print("  --output-filename NAME")
        print("      Output CSV filename (default: _CVOfflineCheck.csv)")
        print('      Example: --output-filename "MyReport.csv"')
        print()
        print("Full example:")
        print('  CVOfflineCheck_v2.exe --diffreport "\\\\admssim01\\ADMS_DataEngineering\\CompositeViewBackup\\DiffReport" --pdf-path "\\\\admssim01\\ADMS_DataEngineering\\DMS_Picture_offline" --output "D:\\MyReports"')
        print()
        print("Using default paths:")
        print('  CVOfflineCheck_v2.exe')
        print("="*80)
        input("\nPress ENTER to exit...")
        return None
    # Option 3: Custom path input
    else:
        print("\n" + "="*80)
        print("CUSTOM PATH CONFIGURATION")
        print("="*80)
        # Use the input as DiffReport path
        diffreport = user_input
        # Ask for PDF path
        print(f"\nDiffReport path set to: {diffreport}")
        pdf_path_input = input("\nEnter PDF base path (or press ENTER for default): ").strip()
        if pdf_path_input == '':
            pdf_path = r"\\admssim01\ADMS_DataEngineering\DMS_Picture_offline"
            print(f"Using default: {pdf_path}")
        else:
            pdf_path = pdf_path_input
        # Ask for output path
        output_input = input("\nEnter output folder path (or press ENTER for default): ").strip()
        if output_input == '':
            output = r"C:\ADMS_DataEngineering\DMS_Picture_offline"
            print(f"Using default: {output}")
        else:
            output = output_input
        print("\n" + "-"*80)
        print("Configuration summary:")
        print(f"  DiffReport: {diffreport}")
        print(f"  PDF Base:   {pdf_path}")
        print(f"  Output:     {output}")
        print("-"*80)
        confirm = input("\nProceed with these paths? (Y/n): ").strip().lower()
        if confirm == 'n':
            print("\n✗ Configuration cancelled.")
            input("Press ENTER to exit...")
            return None
        return diffreport, pdf_path, output
def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Check PDF status against DiffReport CSV files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (double-click exe or run without parameters)
  CVOfflineCheck_v2.exe
  # Use default paths
  CVOfflineCheck_v2.exe
  # Custom DiffReport path only
  CVOfflineCheck_v2.exe --diffreport "\\\\server\\path\\DiffReport"
  # Full custom configuration
  CVOfflineCheck_v2.exe --diffreport "\\\\server\\path\\DiffReport" --pdf-path "\\\\server\\path\\PDFs" --output "C:\\MyOutput"
Interactive Mode Options:
  - Type 'Draft' → Use default paths
  - Type '-'     → Show this help message
  - Enter path   → Configure custom paths interactively
        """
    )
    parser.add_argument(
        '--diffreport',
        type=str,
        default=r"\\admssim01\ADMS_DataEngineering\CompositeViewBackup\DiffReport",
        help='Path to DiffReport folder containing CSV files'
    )
    parser.add_argument(
        '--pdf-path',
        type=str,
        default=r"\\admssim01\ADMS_DataEngineering\DMS_Picture_offline",
        help='Base path to search for PDF files'
    )
    parser.add_argument(
        '--output',
        type=str,
        default=r"C:\ADMS_DataEngineering\DMS_Picture_offline",
        help='Output folder for result CSV'
    )
    parser.add_argument(
        '--output-filename',
        type=str,
        default='_CVOfflineCheck.csv',
        help='Output CSV filename (default: _CVOfflineCheck.csv)'
    )
    return parser.parse_args()
if __name__ == "__main__":
    # Check if running with command-line arguments
    if len(sys.argv) > 1:
        # Command-line mode
        args = parse_arguments()
        diffreport_path = args.diffreport
        pdf_base_path = args.pdf_path
        output_folder = args.output
        output_filename = args.output_filename
    else:
        # Interactive mode (double-clicked exe)
        path_config = interactive_path_input()
        if path_config is None:
            # User chose to exit or cancelled
            sys.exit(0)
        diffreport_path, pdf_base_path, output_folder = path_config
        output_filename = '_CVOfflineCheck.csv'
    # Display configuration
    print("\n" + "="*80)
    print("CONFIGURATION:")
    print("-" * 80)
    print(f"DiffReport folder: {diffreport_path}")
    print(f"PDF base path: {pdf_base_path}")
    print(f"Output folder: {output_folder}")
    print(f"Output filename: {output_filename}")
    # Validate paths
    print("\nValidating paths...")
    # Check DiffReport path
    if not os.path.exists(diffreport_path):
        print("="*80)
        print(f"✗ ERROR: DiffReport folder does not exist!")
        print(f"  Specified path: {diffreport_path}")
        print(f"\n  Please check:")
        print(f"    1. Network connection (if using \\\\server\\path)")
        print(f"    2. Path spelling and format")
        print(f"    3. Access permissions")
        print("="*80)
        input("\nPress ENTER to exit...")
        sys.exit(1)
    else:
        print(f"  ✓ DiffReport folder exists")
    # Check PDF path (warning only)
    if not os.path.exists(pdf_base_path):
        print(f"  ⚠ WARNING: PDF base path does not exist: {pdf_base_path}")
        print(f"    Program will continue but may not find PDF files.")
    else:
        print(f"  ✓ PDF base path exists")
    print("="*80 + "\n")
    # Create output folder
    os.makedirs(output_folder, exist_ok=True)
    # Run the main check
    result = check_pdf_status(
        diffreport_folder=diffreport_path,
        pdf_base_path=pdf_base_path,
        output_folder=output_folder,
        output_filename=output_filename
    )
    if result is not None:
        print("\n" + "="*80)
        print("✓ PROGRAM COMPLETED SUCCESSFULLY!")
        print("="*80)
        print(f"\nOutput file: {os.path.join(output_folder, output_filename)}")
    else:
        print("\n" + "="*80)
        print("✗ PROGRAM STOPPED DUE TO VALIDATION ERRORS!")
        print("="*80)
    input("\nPress ENTER to exit...")
 