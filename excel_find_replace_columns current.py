"""
Excel Find/Replace Script using xlwings - Column-Specific Version
Works on both Mac and Windows with Excel installed
Preserves ALL Excel features including data validation
Uses reliable cell-by-cell processing for guaranteed results
Now supports different target columns for different sheets
"""

"""
Created by Freddie Harris - 8bitMango on GitHub
"""

""" 
Enable backups on line 39
Define directory on line 518
Define find and replace pairs on line 528
Define target columns on line 605
"""

import os
import sys
import shutil
import platform
from datetime import datetime
import logging
import re
import unicodedata
import time

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings is not installed!")
    print("Please install it with: pip install xlwings")
    print("You also need Microsoft Excel installed on your system")
    sys.exit(1)

# Backup configuration (disabled by default)
ENABLE_BACKUPS = False

def setup_logging():
    """Set up logging to track progress and errors"""
    log_filename = f"find_replace_xlwings_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler()
        ]
    )
    return log_filename

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

def parse_column_range(column_spec):
    """
    Parse column specification into list of column indices
    Supports:
    - Single column: 'A', 'B', 'Z'
    - Range: 'A:C', 'B:F' 
    - List: 'A,C,E', 'B,D:F,H'
    - Mix: 'A,C:E,G'
    Returns list of zero-based column indices
    """
    if not column_spec:
        return None
    
    columns = set()
    parts = [part.strip() for part in column_spec.split(',')]
    
    for part in parts:
        if ':' in part:
            # Range like 'A:C' or 'B:F'
            start_col, end_col = part.split(':')
            start_idx = column_letter_to_index(start_col.strip())
            end_idx = column_letter_to_index(end_col.strip())
            for idx in range(start_idx, end_idx + 1):
                columns.add(idx)
        else:
            # Single column like 'A' or 'B'
            columns.add(column_letter_to_index(part))
    
    return sorted(list(columns))

def process_sheet_intelligently(sheet, replacement_pairs, target_columns=None, max_cells_to_check=50000):
    """
    Process sheet with intelligent cell checking - stops early if no text found
    Now supports targeting specific columns
    """
    try:
        print(f"      Analyzing sheet structure")
        
        # Parse target columns
        column_indices = None
        if target_columns:
            column_indices = parse_column_range(target_columns)
            if column_indices:
                column_letters = []
                for idx in column_indices:
                    # Convert back to letter for display
                    letter = ""
                    temp = idx + 1
                    while temp > 0:
                        temp -= 1
                        letter = chr(temp % 26 + ord('A')) + letter
                        temp //= 26
                    column_letters.append(letter)
                print(f"      Targeting columns: {', '.join(column_letters)} (indices: {column_indices})")
        
        # Get used range safely
        try:
            used_range = sheet.used_range
            if used_range is None:
                print(f"      Sheet appears empty, skipping")
                return 0, {}
            
            rows, cols = used_range.shape
            # Limit processing to first 300 rows
            rows_to_process = min(rows, 300)
            
            # If we have target columns, limit cols to only those we need
            if column_indices:
                max_target_col = max(column_indices) + 1
                cols = min(cols, max_target_col)
            
            total_cells = rows_to_process * (len(column_indices) if column_indices else cols)
            
            print(f"      Sheet has {rows} rows x {cols} columns; processing first {rows_to_process} rows ({total_cells:,} target cells)")
            
        except Exception as e:
            print(f"      Could not determine sheet size: {e}")
            return 0, {}
        
        # For very large sheets, proceed with controlled scan
        if total_cells > max_cells_to_check:
            print(f"      Large sheet detected - proceeding with controlled scan")
        
        # Process the sheet in manageable chunks
        return process_sheet_in_chunks(sheet, used_range, replacement_pairs, rows_to_process, cols, column_indices)
        
    except Exception as e:
        print(f"      Error analyzing sheet: {e}")
        return 0, {}

def process_sheet_in_chunks(sheet, used_range, replacement_pairs, rows, cols, target_column_indices):
    """
    Process sheet in small chunks to avoid memory/timeout issues
    Now supports targeting specific columns
    """
    total_replacements = 0
    replacement_details = {}
    affected_rows_overall = set()
    
    # Initialize tracking
    for find_text, replace_text in replacement_pairs:
        replacement_details[find_text] = 0

    # Precompile regex patterns once per sheet for performance
    compiled_patterns = []  # list of tuples: (find_text, compiled_regex, replace_text)
    for find_text, replace_text in replacement_pairs:
        compiled_patterns.append(
            (find_text, re.compile(re.escape(find_text), flags=re.IGNORECASE), replace_text)
        )
    
    # Determine optimal chunk size based on sheet size
    if rows > 100000:
        chunk_size = 500
    elif rows > 10000:
        chunk_size = 1000
    else:
        chunk_size = 5000
    
    print(f"      Processing {rows:,} rows in chunks of {chunk_size}")
    
    # Process in row chunks
    chunks_processed = 0
    chunks_with_changes = 0
    
    for start_row in range(0, rows, chunk_size):
        end_row = min(start_row + chunk_size, rows)
        chunks_processed += 1
        
        if chunks_processed % 10 == 0:  # Progress update every 10 chunks
            print(f"        Progress: {chunks_processed * chunk_size:,}/{rows:,} rows processed")
        
        try:
            # Process this chunk
            chunk_replacements, affected_rows = process_chunk(
                used_range, start_row, end_row, cols, compiled_patterns, target_column_indices
            )
            
            if chunk_replacements:
                total_replacements += sum(chunk_replacements.values())
                chunks_with_changes += 1
                affected_rows_overall.update(affected_rows)
                
                # Add to overall details
                for find_text, count in chunk_replacements.items():
                    replacement_details[find_text] += count
                    
        except Exception as chunk_error:
            print(f"        Warning: Error processing rows {start_row+1}-{end_row}: {chunk_error}")
            continue
    
    if chunks_with_changes > 0:
        # Set row height to 16 for all rows that had changes (cross-platform)
        try:
            for sheet_row in sorted(affected_rows_overall):
                try:
                    sheet.range(f"{sheet_row}:{sheet_row}").row_height = 16
                except Exception:
                    continue
        except Exception:
            pass
        print(f"      Completed: {total_replacements} replacements in {chunks_with_changes}/{chunks_processed} chunks")
    else:
        print(f"      Completed: No text requiring replacement found")
    
    return total_replacements, replacement_details

def process_chunk(used_range, start_row, end_row, cols, compiled_patterns, target_column_indices):
    """
    Process a specific chunk of rows
    Now supports targeting specific columns
    """
    chunk_replacements = {}
    affected_rows = set()
    
    # Initialize tracking for this chunk
    for find_text, _, _ in compiled_patterns:
        chunk_replacements[find_text] = 0
    
    # Check each row in the chunk
    for row_idx in range(start_row, end_row):
        try:
            # Determine which columns to check
            if target_column_indices:
                columns_to_check = [col_idx for col_idx in target_column_indices if col_idx < cols]
            else:
                columns_to_check = range(min(cols, 108))  # Default: limit to columns A:DD
            
            # Check each target column in this row
            for col_idx in columns_to_check:
                try:
                    cell = used_range[row_idx, col_idx]
                    cell_value = cell.value
                    
                    # Only process if it's a string with content
                    if cell_value and isinstance(cell_value, str) and len(cell_value.strip()) > 0:
                        
                        original_value = cell_value
                        modified = False
                        
                        # Apply all replacement patterns to this cell (single pass per pattern)
                        for find_text, compiled_re, replace_text in compiled_patterns:
                            new_value, num_subs = compiled_re.subn(replace_text, cell_value)
                            if num_subs > 0:
                                cell_value = new_value
                                chunk_replacements[find_text] += num_subs
                                modified = True
                        
                        # Update the cell if we made changes
                        if modified:
                            cell.value = cell_value
                            try:
                                affected_rows.add(cell.row)
                            except Exception:
                                pass
                            
                except Exception as cell_error:
                    # Skip problematic cells
                    continue
                    
        except Exception as row_error:
            # Skip problematic rows
            continue
    
    return chunk_replacements, affected_rows

def _normalize_sheet_name(name):
    """Normalize sheet names: unicode fold, trim, collapse inner spaces, lowercase."""
    if not isinstance(name, str):
        return ""
    # Unicode normalize and casefold
    normalized = unicodedata.normalize('NFKC', name).strip()
    # Collapse multiple whitespace to single space
    normalized = re.sub(r"\s+", " ", normalized)
    return normalized.lower()


def process_excel_with_xlwings(filepath, replacement_pairs, sheet_configurations=None):
    """
    Process Excel file using xlwings - simplified reliable approach
    Now supports sheet-specific column configurations
    
    sheet_configurations should be a dict like:
    {
        "Sheet Name": {
            "columns": "A,C,E",  # or "A:C" or "B,D:F,H" etc.
            "enabled": True
        }
    }
    """
    app = None
    wb = None
    
    try:
        print(f"  Processing {os.path.basename(filepath)}")

        # Optional backup
        try:
            if ENABLE_BACKUPS:
                backup_path = create_backup(filepath)
                print(f"  Backup created: {os.path.basename(backup_path)}")
        except Exception as backup_err:
            print(f"  Backup warning: {backup_err}")
        
        # Start Excel application
        print(f"  Starting Excel")
        app = xw.App(visible=False, add_book=False)
        # Reduce prompts and speed up processing
        try:
            app.display_alerts = False
            app.screen_updating = False
            # Further reduce prompts (best-effort; may not be available on all platforms)
            try:
                app.api.AskToUpdateLinks = False
            except Exception:
                pass
        except Exception:
            pass
        
        # Temporarily disable calc/events for performance; restore in finally
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

        # Open workbook with timeout protection
        print(f"  Opening workbook")
        # On macOS, opening via basename after chdir avoids repeated "Grant access" prompts
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
        total_replacements = 0
        replacement_details = {}
        
        # Initialize tracking
        for find_text, replace_text in replacement_pairs:
            replacement_details[find_text] = 0
        
        print(f"  Found {len(wb.sheets)} sheets to process")

        for i, sheet in enumerate(wb.sheets, 1):
            sheet_name = sheet.name
            print(f"    Sheet {i}/{len(wb.sheets)}: {sheet_name}")

            # Check sheet configuration
            if sheet_configurations:
                # Find matching configuration (case-insensitive)
                sheet_config = None
                normalized_sheet_name = _normalize_sheet_name(sheet_name)
                
                for config_name, config in sheet_configurations.items():
                    if _normalize_sheet_name(config_name) == normalized_sheet_name:
                        sheet_config = config
                        break
                
                if not sheet_config:
                    print(f"      Skipping (not in configuration)")
                    continue
                    
                if not sheet_config.get('enabled', True):
                    print(f"      Skipping (disabled in configuration)")
                    continue
                
                target_columns = sheet_config.get('columns', None)
            else:
                target_columns = None

            # Warn and attempt unprotect if sheet is protected (prevents writing changes)
            try:
                if hasattr(sheet, 'api') and getattr(sheet.api, 'ProtectContents', False):
                    print(f"      Notice: Sheet is protected; attempting to unprotect")
                    try:
                        sheet.api.Unprotect(Password="")
                    except Exception:
                        print(f"      Warning: Could not unprotect sheet. Replacements may be skipped.")
            except Exception:
                pass
            
            try:
                sheet_replacements, sheet_replacement_details = process_sheet_intelligently(
                    sheet, replacement_pairs, target_columns
                )
                
                if sheet_replacements > 0:
                    print(f"      SUCCESS: {sheet_replacements} replacements made")
                    logging.info(f"  Sheet '{sheet_name}': {sheet_replacements} replacements")
                    total_replacements += sheet_replacements
                    modifications_made = True
                    
                    # Add to overall details
                    for find_text, count in sheet_replacement_details.items():
                        replacement_details[find_text] += count
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
                print(f"  SUCCESS: {total_replacements} total replacements saved")
                print(f"  ALL Excel features preserved!")
                
                replacement_summary = ", ".join([f"'{find}': {count}" for find, count in replacement_details.items() if count > 0])
                logging.info(f"SUCCESS {os.path.basename(filepath)}: {total_replacements} total replacements ({replacement_summary})")
                
                return True, total_replacements, replacement_details
            except Exception as save_error:
                print(f"  ERROR saving file: {save_error}")
                return False, 0, {}
        else:
            print(f"  No changes needed")
            logging.info(f"- {os.path.basename(filepath)}: No matches found")
            return True, 0, replacement_details
            
    except Exception as e:
        print(f"  ERROR: {str(e)}")
        logging.error(f"Failed to process {os.path.basename(filepath)}: {str(e)}")
        return False, 0, {}
        
    finally:
        # Clean up
        try:
            # Restore calc/events
            try:
                if 'events_prev' in locals() and events_prev is not None:
                    app.api.EnableEvents = events_prev
            except Exception:
                pass
            try:
                if 'calc_prev' in locals() and calc_prev is not None:
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
    # Configuration
    if platform.system() == "Windows":
        directory = r"C:\Users\Freddieharris\Desktop\Bulk Replace Test"  # Windows path
    else:  # macOS/Linux
        directory = "/Users/freddieharris/Downloads/untitled folder"  # Mac path
    
    # Define your find/replace pairs here
    replacement_pairs = [
        ("Original Series MK2 Test Steel", "OS3 T Line Steel"),
        ("Original Series MK2 range", "OS3 T Line range"),
        ("Elite model", "E Line model"),
        ("Original Series MK2 Test Titanium", "OS3 T Line Titanium"),
    ]
    
    # Setup
    system_info = f"{platform.system()} {platform.release()}"
    print(f"Running on: {system_info}")
    print(f"Python version: {sys.version.split()[0]}")
    
    # Check Excel availability
    if not check_excel_availability():
        print("\nCannot proceed without Excel. Please install Microsoft Excel and try again.")
        return
    
    log_file = setup_logging()
    logging.info(f"Starting xlwings Excel find and replace operation on {system_info}")
    logging.info(f"Directory: {directory}")
    
    print(f"\nLooking for Excel files in: {directory}")
    
    # Validate directory
    if not os.path.exists(directory):
        print(f"Directory does not exist: {directory}")
        print("Please update the 'directory' variable in the script with the correct path")
        return
    
    try:
        # Filter out temporary Excel files (starting with ~$)
        excel_files = [f for f in os.listdir(directory) 
                      if f.endswith(('.xlsx', '.xlsm', '.xls')) and not f.startswith('~$')]
        total_files = len(excel_files)
        print(f"Found {total_files} Excel files (excluding temporary files)")
        
        if total_files == 0:
            print("No Excel files found (temporary files starting with ~$ are ignored)")
            print("Files in directory:")
            for f in os.listdir(directory)[:10]:  # Show first 10 files
                print(f"  {f}")
            return
        
        print("Excel files found:")
        for f in excel_files[:5]:  # Show first 5
            print(f"  {f}")
        if total_files > 5:
            print(f" and {total_files - 5} more")
            
    except Exception as e:
        print(f"Error reading directory: {e}")
        return
    
    print(f"\nStarting processing")
    print(f"Replacement pairs:")
    for i, (find_text, replace_text) in enumerate(replacement_pairs, 1):
        print(f"  {i}. '{find_text}' -> '{replace_text}'")
    print(f"ALL Excel features will be preserved!")
    
    # Process files
    successful_files = 0
    failed_files = 0
    total_replacements = 0
    overall_replacement_stats = {}
    
    # Initialize overall stats tracking
    for find_text, replace_text in replacement_pairs:
        overall_replacement_stats[find_text] = 0
    
    start_time = time.time()

    # Change working directory to target to avoid macOS grant-access prompts
    try:
        os.chdir(directory)
    except Exception:
        pass

    """
    Configure which sheets to process and which columns to target
    - Single column: 'A', 'B', 'Z'
    - Range: 'A:C', 'B:F' 
    - List: 'A,C,E', 'B,D:F,H'
    - Mix: 'A,C:E,G'
    All excepted
    """
    sheet_configurations = {
        "Std Cricket Products Upload": {
            "columns": "Z",  # Only process columns A, C, E, G
            "enabled": True
        },
        "Product URL formula sheet": {
            "columns": "R",  # Process columns B through F
            "enabled": True
        },
        "Sub Categories": {
            "columns": "S",  # Process column A, columns D through F, and column H
            "enabled": True
        }
    }
    
    print(f"\nSheet configurations:")
    for sheet_name, config in sheet_configurations.items():
        status = "ENABLED" if config.get('enabled', True) else "DISABLED"
        columns = config.get('columns', 'ALL')
        print(f"  '{sheet_name}': {status}, Columns: {columns}")
    
    for i, filename in enumerate(excel_files, 1):
        filepath = os.path.join(directory, filename)
        print(f"\nFile {i}/{total_files}: {filename}")
        logging.info(f"Processing {i}/{total_files}: {filename}")
        
        file_start_time = time.time()
        
        success, replacements, replacement_details = process_excel_with_xlwings(
            filepath,
            replacement_pairs,
            sheet_configurations=sheet_configurations,
        )
        
        file_duration = time.time() - file_start_time
        print(f"  Processing time: {file_duration:.1f} seconds")
        
        if success:
            successful_files += 1
            total_replacements += replacements
            # Add to overall stats
            for find_text, count in replacement_details.items():
                overall_replacement_stats[find_text] += count
        else:
            failed_files += 1
        
        # Small delay between files
        time.sleep(1)
    
    total_duration = time.time() - start_time
    
    # Summary
    print("\n" + "="*60)
    print("PROCESSING COMPLETE!")
    print(f"Total processing time: {total_duration:.1f} seconds")
    print(f"Total files processed: {total_files}")
    print(f"Successful: {successful_files}")
    print(f"Failed: {failed_files}")
    print(f"Total replacements made: {total_replacements}")
    
    if total_replacements > 0:
        print(f"\nBreakdown by replacement type:")
        for find_text, count in overall_replacement_stats.items():
            if count > 0:
                print(f"  '{find_text}': {count} replacements")
    
    print(f"\nALL Excel features preserved!")
    print(f"Log saved: {log_file}")
    
    if failed_files > 0:
        print(f"\n{failed_files} files failed to process. Check the log for details.")

if __name__ == "__main__":
    main()