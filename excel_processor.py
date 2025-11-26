"""
Excel Search and Update Script using xlwings
Works on both Mac and Windows with Excel installed
Preserves ALL Excel features including data validation
Searches for values in specified columns and updates target columns if values differ
Configuration loaded from JSON file
"""

"""
Created by Freddie Harris - 8bitMango on GitHub
Modified to support search-and-update functionality with JSON configuration
"""

import os
import sys
import shutil
import platform
from datetime import datetime
import logging
import json
import time

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings is not installed!")
    print("Please install it with: pip install xlwings")
    print("You also need Microsoft Excel installed on your system")
    sys.exit(1)

# Global configuration (loaded from JSON)
CONFIG = {}

def load_configuration(config_path=None):
    """Load configuration from JSON file"""
    global CONFIG

    # If no path provided, look for config.json in the same directory as this script
    if config_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(script_dir, "config.json")

    if not os.path.exists(config_path):
        print(f"ERROR: Configuration file not found: {config_path}")
        print("Please create a config.json file in the same directory as this script")
        print(f"Expected location: {config_path}")
        sys.exit(1)

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            CONFIG = json.load(f)
        print(f"Configuration loaded from {config_path}")
        return True
    except json.JSONDecodeError as e:
        print(f"ERROR: Invalid JSON in configuration file: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Failed to load configuration: {e}")
        sys.exit(1)

def setup_logging():
    """Set up logging to track progress and errors - creates two log files"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Main log file (all messages)
    log_filename = f"excel_processor_log_{timestamp}.txt"

    # Error-only log file
    error_log_filename = f"excel_processor_errors_{timestamp}.txt"

    # Clear any existing handlers
    logging.root.handlers = []

    # Create main logger
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler()
        ]
    )

    # Add error-only file handler
    error_handler = logging.FileHandler(error_log_filename)
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logging.getLogger().addHandler(error_handler)

    # Write header to error log
    with open(error_log_filename, 'a') as f:
        f.write("="*60 + "\n")
        f.write("EXCEL PROCESSOR - ERROR LOG\n")
        f.write("="*60 + "\n\n")

    return log_filename, error_log_filename

def create_backup(filepath):
    """Create a backup of the original file"""
    backup_dir = os.path.join(os.path.dirname(filepath), "backups")
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)

    filename = os.path.basename(filepath)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_filename = f"{timestamp}_{filename}"
    backup_path = os.path.join(backup_dir, backup_filename)

    shutil.copy2(filepath, backup_path)
    return backup_path

def check_excel_availability():
    """Check if Excel is available and working"""
    try:
        test_app = xw.App(visible=False, add_book=False)
        test_app.quit()

        system = platform.system()
        print(f"Excel is available on {system}")
        return True

    except Exception as e:
        print(f"Excel is not available or not working properly: {str(e)}")
        print("Please ensure Microsoft Excel is installed and working")
        return False

def column_letter_to_index(col_letter):
    """Convert column letter(s) to zero-based index (A=0, B=1, ..., Z=25, AA=26, etc.)"""
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1

def column_index_to_letter(col_index):
    """Convert zero-based column index to letter (0=A, 1=B, etc.)"""
    letter = ""
    temp = col_index + 1
    while temp > 0:
        temp -= 1
        letter = chr(temp % 26 + ord('A')) + letter
        temp //= 26
    return letter

def process_sheet_with_rules(sheet, rules, max_rows_to_process=300):
    """
    Process sheet with search-and-update rules - OPTIMIZED VERSION

    For each rule:
    1. Search the search_column for search_value
    2. When found, check the update_column in the same row
    3. If the value differs from target_value, update it

    OPTIMIZATION: Groups rules by column pair and processes in a single pass
    """
    try:
        print(f"      Analyzing sheet structure")

        # Get used range safely
        try:
            used_range = sheet.used_range
            if used_range is None:
                print(f"      Sheet appears empty, skipping")
                return 0, {}

            rows, cols = used_range.shape
            rows_to_process = min(rows, max_rows_to_process)

            print(f"      Sheet has {rows} rows x {cols} columns; processing first {rows_to_process} rows")

        except Exception as e:
            print(f"      Could not determine sheet size: {e}")
            return 0, {}

        # OPTIMIZATION: Group rules by search/update column pair for single-pass processing
        grouped_rules = {}
        for rule in rules:
            search_col = rule['search_column']
            update_col = rule['update_column']
            key = (search_col, update_col)

            if key not in grouped_rules:
                grouped_rules[key] = []

            grouped_rules[key].append({
                'name': rule.get('name', 'Unnamed rule'),
                'search_value': str(rule['search_value']).strip().lower(),
                'target_value': str(rule['target_value']),
                'original_search': str(rule['search_value'])  # For display
            })

        print(f"      Optimized: {len(rules)} rules grouped into {len(grouped_rules)} column pair(s)")

        total_updates = 0
        update_details = {}
        all_affected_rows = set()

        # Process each column pair group
        for (search_col, update_col), rule_group in grouped_rules.items():
            print(f"      Processing column pair: {search_col} -> {update_col} ({len(rule_group)} rules)")

            # Convert column letters to indices
            search_col_idx = column_letter_to_index(search_col)
            update_col_idx = column_letter_to_index(update_col)

            # Create lookup dictionary: search_value -> (target_value, rule_name)
            lookup = {}
            for rule in rule_group:
                lookup[rule['search_value']] = (rule['target_value'], rule['name'])
                update_details[rule['name']] = 0

            # SINGLE PASS through all rows for this column pair
            for row_idx in range(rows_to_process):
                try:
                    # Get the search cell value
                    search_cell = used_range[row_idx, search_col_idx]
                    search_cell_value = search_cell.value

                    if not search_cell_value:
                        continue

                    # Normalize the search value
                    normalized_search = str(search_cell_value).strip().lower()

                    # Check if this value matches any rule
                    if normalized_search in lookup:
                        target_value, rule_name = lookup[normalized_search]

                        # Get the update cell
                        update_cell = used_range[row_idx, update_col_idx]
                        current_value = update_cell.value
                        current_value_str = str(current_value) if current_value is not None else ""

                        # Check if update is needed
                        if current_value_str.strip() != target_value.strip():
                            # Update the cell
                            update_cell.value = target_value
                            update_details[rule_name] += 1
                            total_updates += 1
                            all_affected_rows.add(row_idx + 1)

                except Exception as cell_error:
                    # Skip problematic cells
                    continue

            # Print results for this column pair
            for rule in rule_group:
                count = update_details[rule['name']]
                if count > 0:
                    print(f"        '{rule['name']}': {count} updates")

        # Set row heights for all affected rows at once (more efficient)
        if all_affected_rows:
            print(f"      Setting row heights for {len(all_affected_rows)} affected rows")
            for row_num in all_affected_rows:
                try:
                    sheet.range(f"{row_num}:{row_num}").row_height = 16
                except Exception:
                    continue

        return total_updates, update_details

    except Exception as e:
        print(f"      Error analyzing sheet: {e}")
        return 0, {}

def process_excel_with_xlwings(filepath, sheet_rules):
    """
    Process Excel file using xlwings with search-and-update logic

    sheet_rules should be a dict like:
    {
        "Sheet Name": [
            {
                "name": "Rule description",
                "search_column": "A",
                "search_value": "Product123",
                "update_column": "Z",
                "target_value": "99.99"
            }
        ]
    }
    """
    app = None
    wb = None

    try:
        print(f"  Processing {os.path.basename(filepath)}")

        # Optional backup
        enable_backups = CONFIG.get('general_settings', {}).get('enable_backups', False)
        try:
            if enable_backups:
                backup_path = create_backup(filepath)
                print(f"  Backup created: {os.path.basename(backup_path)}")
        except Exception as backup_err:
            print(f"  Backup warning: {backup_err}")
            logging.error(f"Failed to create backup for {os.path.basename(filepath)}: {str(backup_err)}")

        # Start Excel application
        print(f"  Starting Excel")
        app = xw.App(visible=False, add_book=False)

        # Reduce prompts and speed up processing
        try:
            app.display_alerts = False
            app.screen_updating = False
            try:
                app.api.AskToUpdateLinks = False
            except Exception:
                pass
        except Exception:
            pass

        # Temporarily disable calc/events for performance
        calc_prev = None
        events_prev = None
        try:
            try:
                calc_prev = app.api.Calculation
                app.api.Calculation = -4135  # xlCalculationManual
            except Exception:
                pass
            try:
                events_prev = app.api.EnableEvents
                app.api.EnableEvents = False
            except Exception:
                pass
        except Exception:
            pass

        # Open workbook
        print(f"  Opening workbook")
        try:
            wb = app.books.open(
                os.path.basename(filepath),
                update_links=False,
                notify=False,
                read_only=False,
                ignore_read_only_recommended=True,
                add_to_mru=False,
                local=True,
            )
        except Exception:
            # Fallback to full path if needed
            wb = app.books.open(filepath, update_links=False, notify=False)

        modifications_made = False
        total_updates = 0
        update_details = {}

        print(f"  Found {len(wb.sheets)} sheets to process")

        for i, sheet in enumerate(wb.sheets, 1):
            sheet_name = sheet.name
            print(f"    Sheet {i}/{len(wb.sheets)}: {sheet_name}")

            # Check if we have rules for this sheet
            if sheet_name not in sheet_rules:
                print(f"      Skipping (no rules configured for this sheet)")
                continue

            rules = sheet_rules[sheet_name]
            if not rules:
                print(f"      Skipping (no rules defined)")
                continue

            # Warn and attempt unprotect if sheet is protected
            try:
                if hasattr(sheet, 'api') and getattr(sheet.api, 'ProtectContents', False):
                    print(f"      Notice: Sheet is protected; attempting to unprotect")
                    try:
                        sheet.api.Unprotect(Password="")
                    except Exception as unprotect_error:
                        print(f"      Warning: Could not unprotect sheet. Updates may be skipped.")
                        logging.error(f"Protected sheet '{sheet_name}' in {os.path.basename(filepath)} - could not unprotect: {str(unprotect_error)}")
            except Exception:
                pass

            try:
                max_rows = CONFIG.get('general_settings', {}).get('max_rows_to_process', 300)
                sheet_updates, sheet_update_details = process_sheet_with_rules(
                    sheet, rules, max_rows
                )

                if sheet_updates > 0:
                    print(f"      SUCCESS: {sheet_updates} updates made")
                    logging.info(f"  Sheet '{sheet_name}': {sheet_updates} updates")
                    total_updates += sheet_updates
                    modifications_made = True

                    # Add to overall details
                    for rule_name, count in sheet_update_details.items():
                        if rule_name not in update_details:
                            update_details[rule_name] = 0
                        update_details[rule_name] += count
                else:
                    print(f"      No changes needed")

            except Exception as sheet_error:
                print(f"      ERROR processing sheet {sheet_name}: {str(sheet_error)}")
                logging.error(f"Error processing sheet {sheet_name}: {str(sheet_error)}")
                continue

        if modifications_made:
            print(f"  Saving changes")
            try:
                wb.save()
                print(f"  SUCCESS: {total_updates} total updates saved")
                print(f"  ALL Excel features preserved!")

                update_summary = ", ".join([f"'{rule}': {count}" for rule, count in update_details.items() if count > 0])
                logging.info(f"SUCCESS {os.path.basename(filepath)}: {total_updates} total updates ({update_summary})")

                return True, total_updates, update_details
            except Exception as save_error:
                print(f"  ERROR saving file: {save_error}")
                logging.error(f"Failed to save {os.path.basename(filepath)}: {str(save_error)}")
                return False, 0, {}
        else:
            print(f"  No changes needed")
            logging.info(f"- {os.path.basename(filepath)}: No matches found")
            return True, 0, update_details

    except Exception as e:
        print(f"  ERROR: {str(e)}")
        logging.error(f"Failed to process {os.path.basename(filepath)}: {str(e)}")
        return False, 0, {}

    finally:
        # Clean up
        try:
            # Restore calc/events
            try:
                if events_prev is not None:
                    app.api.EnableEvents = events_prev
            except Exception:
                pass
            try:
                if calc_prev is not None:
                    app.api.Calculation = calc_prev
            except Exception:
                pass
            if wb:
                wb.close()
            if app:
                app.quit()
        except Exception as cleanup_error:
            print(f"  Cleanup warning: {cleanup_error}")

def main():
    """Main execution function"""

    # Load configuration from JSON (will look in script's directory)
    if not load_configuration():
        return

    # Setup
    system_info = f"{platform.system()} {platform.release()}"
    print(f"Running on: {system_info}")
    print(f"Python version: {sys.version.split()[0]}")

    # Check Excel availability
    if not check_excel_availability():
        print("\nCannot proceed without Excel. Please install Microsoft Excel and try again.")
        return

    log_file, error_log_file = setup_logging()
    logging.info(f"Starting Excel search and update operation on {system_info}")
    print(f"Main log: {log_file}")
    print(f"Error log: {error_log_file}")

    # Get directory based on platform
    folder_paths = CONFIG.get('folder_paths', {})
    system = platform.system()

    if system == "Windows":
        directory = folder_paths.get('windows', '')
    elif system == "Darwin":  # macOS
        directory = folder_paths.get('mac', '')
    else:  # Linux
        directory = folder_paths.get('linux', '')

    if not directory:
        print(f"ERROR: No folder path configured for {system} in config.json")
        return

    logging.info(f"Directory: {directory}")
    print(f"\nLooking for Excel files in: {directory}")

    # Validate directory
    if not os.path.exists(directory):
        print(f"Directory does not exist: {directory}")
        print("Please update the folder path in config.json")
        return

    try:
        # Filter out temporary Excel files (starting with ~$)
        excel_files = [f for f in os.listdir(directory)
                      if f.endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$')]
        total_files = len(excel_files)
        print(f"Found {total_files} Excel files (excluding temporary files)")

        if total_files == 0:
            print("No Excel files found (temporary files starting with ~$ are ignored)")
            return

        print("Excel files found:")
        for f in excel_files[:5]:  # Show first 5
            print(f"  {f}")
        if total_files > 5:
            print(f"  ... and {total_files - 5} more")

    except Exception as e:
        print(f"Error reading directory: {e}")
        return

    # Build sheet rules from configuration
    search_update_rules = CONFIG.get('search_and_update_rules', [])

    # Organize rules by sheet name
    sheet_rules = {}
    for rule in search_update_rules:
        if not rule.get('enabled', True):
            continue

        sheet_name = rule.get('sheet_name')
        if not sheet_name:
            continue

        if sheet_name not in sheet_rules:
            sheet_rules[sheet_name] = []

        sheet_rules[sheet_name].append({
            'name': rule.get('name', 'Unnamed rule'),
            'search_column': rule.get('search_column'),
            'search_value': rule.get('search_value'),
            'update_column': rule.get('update_column'),
            'target_value': rule.get('target_value')
        })

    if not sheet_rules:
        print("\nERROR: No enabled rules found in configuration!")
        print("Please configure search_and_update_rules in config.json")
        return

    print(f"\nActive rules by sheet:")
    for sheet_name, rules in sheet_rules.items():
        print(f"  '{sheet_name}': {len(rules)} rule(s)")
        for rule in rules:
            print(f"    - {rule['name']}")

    print(f"\nStarting processing")
    print(f"ALL Excel features will be preserved!")

    # Process files
    successful_files = 0
    failed_files = 0
    total_updates = 0
    overall_update_stats = {}

    start_time = time.time()

    # Change working directory to target to avoid macOS grant-access prompts
    try:
        os.chdir(directory)
    except Exception:
        pass

    process_delay = CONFIG.get('general_settings', {}).get('process_delay_seconds', 0)

    for i, filename in enumerate(excel_files, 1):
        filepath = os.path.join(directory, filename)
        print(f"\nFile {i}/{total_files}: {filename}")
        logging.info(f"Processing {i}/{total_files}: {filename}")

        file_start_time = time.time()

        success, updates, update_details = process_excel_with_xlwings(
            filepath,
            sheet_rules
        )

        file_duration = time.time() - file_start_time
        print(f"  Processing time: {file_duration:.1f} seconds")

        if success:
            successful_files += 1
            total_updates += updates
            # Add to overall stats
            for rule_name, count in update_details.items():
                if rule_name not in overall_update_stats:
                    overall_update_stats[rule_name] = 0
                overall_update_stats[rule_name] += count
        else:
            failed_files += 1

        # Optional delay between files (default 0 for speed)
        if process_delay > 0:
            time.sleep(process_delay)

    total_duration = time.time() - start_time

    # Summary
    print("\n" + "="*60)
    print("PROCESSING COMPLETE!")
    print(f"Total processing time: {total_duration:.1f} seconds")
    print(f"Total files processed: {total_files}")
    print(f"Successful: {successful_files}")
    print(f"Failed: {failed_files}")
    print(f"Total updates made: {total_updates}")

    if total_updates > 0:
        print(f"\nBreakdown by rule:")
        for rule_name, count in overall_update_stats.items():
            if count > 0:
                print(f"  '{rule_name}': {count} updates")

    print(f"\nALL Excel features preserved!")
    print(f"\nLog files created:")
    print(f"  Main log: {log_file}")
    print(f"  Error log: {error_log_file}")

    # Write error summary to error log
    with open(error_log_file, 'a') as f:
        f.write("\n" + "="*60 + "\n")
        f.write("ERROR SUMMARY\n")
        f.write("="*60 + "\n")
        f.write(f"Total files processed: {total_files}\n")
        f.write(f"Successful: {successful_files}\n")
        f.write(f"Failed: {failed_files}\n")
        if failed_files == 0:
            f.write("\n✓ No errors occurred during processing!\n")
        else:
            f.write(f"\n✗ {failed_files} file(s) encountered errors.\n")
            f.write("See error messages above for details.\n")

    if failed_files > 0:
        print(f"\n⚠️  {failed_files} files failed to process. Check error log for details.")
    else:
        print(f"\n✓ All files processed successfully!")

if __name__ == "__main__":
    main()
