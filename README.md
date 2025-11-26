# Excel Search and Update Processor

This script processes Excel spreadsheets to search for specific values in columns and update related cells based on configurable rules.

## What Changed from the Original Script

### Original Behavior
- Found and replaced text strings throughout Excel files
- Used hardcoded find/replace pairs in the script
- Targeted specific columns for replacements

### New Behavior
- Searches for a specific value in a designated column
- When found, checks another column in the same row
- Updates that cell only if its value differs from the target value
- All configuration is stored in `config.json` (no hardcoding!)

## Features

- **JSON Configuration**: All settings stored in `config.json`
- **Search and Update**: Find rows by value and conditionally update cells
- **Multi-Rule Support**: Apply multiple rules across different sheets
- **Preserves Excel Features**: Maintains formatting, formulas, data validation, etc.
- **Cross-Platform**: Works on Windows, macOS, and Linux
- **Efficient Processing**: Handles hundreds of Excel files
- **Detailed Logging**: Tracks all changes made

## Requirements

- Python 3.6 or higher
- Microsoft Excel installed on your system
- xlwings library (`pip install xlwings`)

## Installation

1. Install Python dependencies:
   ```bash
   pip install xlwings
   ```

2. Ensure Microsoft Excel is installed on your computer

## Configuration

Edit `config.json` to configure the script behavior:

### Main Configuration Sections

#### 1. General Settings
```json
"general_settings": {
  "enable_backups": false,
  "max_rows_to_process": 300,
  "process_delay_seconds": 1
}
```

#### 2. Folder Paths
Specify where your Excel files are located for each platform:
```json
"folder_paths": {
  "windows": "C:\\Users\\YourName\\Desktop\\Excel Files",
  "mac": "/Users/yourname/Documents/Excel Files",
  "linux": "/home/yourname/excel_files"
}
```

#### 3. Search and Update Rules
Define rules to search for values and update cells:
```json
"search_and_update_rules": [
  {
    "name": "Update product prices",
    "sheet_name": "Products",
    "search_column": "A",
    "search_value": "Product123",
    "update_column": "E",
    "target_value": "99.99",
    "enabled": true
  }
]
```

**Rule Parameters:**
- `name`: Description of what this rule does
- `sheet_name`: The exact name of the sheet to process
- `search_column`: Column to search (e.g., "A", "B", "Z", "AA")
- `search_value`: Value to find in the search column
- `update_column`: Column to check/update in the same row
- `target_value`: Desired value (updates only if different)
- `enabled`: Set to `false` to skip this rule

## How It Works

For each Excel file in the configured folder:

1. **Opens the file** using xlwings
2. **For each enabled rule:**
   - Finds the specified sheet
   - Searches the `search_column` for `search_value`
   - When a match is found:
     - Checks the `update_column` in that same row
     - If the current value differs from `target_value`:
       - Updates it to `target_value`
       - Logs the change
3. **Saves the file** if any changes were made
4. **Moves to the next file**

## Usage

### Basic Usage

1. Edit `config.json` with your settings
2. Run the script:
   ```bash
   python excel_processor.py
   ```

### Example Scenario

You have 800 Excel files with product data. You want to:
- Find rows where column A contains "Product123"
- Check column E (price) in those rows
- Update column E to "99.99" if it's different

**Configuration:**
```json
{
  "folder_paths": {
    "windows": "C:\\Users\\YourName\\Products"
  },
  "search_and_update_rules": [
    {
      "name": "Update Product123 price",
      "sheet_name": "Product List",
      "search_column": "A",
      "search_value": "Product123",
      "update_column": "E",
      "target_value": "99.99",
      "enabled": true
    }
  ]
}
```

### Multiple Rules Example

You can define multiple rules for different sheets and columns:

```json
"search_and_update_rules": [
  {
    "name": "Update product prices",
    "sheet_name": "Products",
    "search_column": "A",
    "search_value": "Widget-Pro",
    "update_column": "F",
    "target_value": "149.99",
    "enabled": true
  },
  {
    "name": "Update product URLs",
    "sheet_name": "Products",
    "search_column": "B",
    "search_value": "SKU-12345",
    "update_column": "H",
    "target_value": "https://example.com/product",
    "enabled": true
  },
  {
    "name": "Update category names",
    "sheet_name": "Categories",
    "search_column": "C",
    "search_value": "Old Category",
    "update_column": "D",
    "target_value": "New Category",
    "enabled": true
  }
]
```

## Output

The script provides:

1. **Console Output**: Real-time progress and results
2. **Log File**: Detailed log saved as `excel_processor_log_YYYYMMDD_HHMMSS.txt`
3. **Summary Report**: Shows total files processed, updates made, and breakdown by rule

Example output:
```
Processing File 1/800: products_2024.xlsx
  Sheet 1/3: Product List
    Applying rule: Update product prices
      Found 5 matches, made 3 updates
  SUCCESS: 3 updates made

PROCESSING COMPLETE!
Total files processed: 800
Successful: 795
Failed: 5
Total updates made: 2,341

Breakdown by rule:
  'Update product prices': 2,341 updates
```

## Backup Option

To enable automatic backups before processing:

```json
"general_settings": {
  "enable_backups": true
}
```

Backups are saved in a `backups` subfolder with timestamps.

## Troubleshooting

### Excel Not Found
- Ensure Microsoft Excel is installed
- On macOS: Grant necessary permissions when prompted

### No Files Found
- Check the folder path in `config.json`
- Ensure the path matches your operating system
- Verify Excel files exist in the folder

### No Updates Made
- Verify `search_value` matches exactly (case-insensitive)
- Check that `sheet_name` matches the actual sheet name
- Ensure the rule is `enabled: true`
- Review the log file for details

### Permission Errors
- Close any open Excel files before running
- On protected sheets, the script will attempt to unprotect

## Tips

1. **Test First**: Start with a copy of your files to test configuration
2. **Use Descriptive Names**: Give rules clear, meaningful names
3. **Enable Logging**: Check log files to verify changes
4. **Disable Rules**: Set `enabled: false` to temporarily skip rules
5. **Column Letters**: Use Excel column letters (A, B, C... AA, AB, etc.)

## Differences from Original Script

| Feature | Original | New |
|---------|----------|-----|
| Configuration | Hardcoded in script | JSON file |
| Operation | Find and replace text | Search and conditional update |
| Target | All occurrences | Specific row/column pairs |
| Flexibility | Requires code editing | Edit config file |
| Multi-rule | Limited | Unlimited rules |

## Support

For issues or questions, refer to the log files or check:
- Column letters are correct
- Sheet names match exactly
- Search values are accurate
- File paths are valid for your OS

## License

Created by Freddie Harris - 8bitMango on GitHub
