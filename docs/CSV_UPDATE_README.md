# CSV Operations Update for Excel MCP Server

This update adds comprehensive CSV import and export capabilities to the Excel MCP Server, including both single file and bulk operations.

## Features Added

- **Single CSV Import/Export**: Import a CSV file into a worksheet or export a worksheet as CSV
- **Bulk CSV Import**: Import multiple CSV files into different worksheets in one operation
- **Bulk CSV Export**: Export multiple worksheets as separate CSV files in one operation
- **Web UI Integration**: Added UI components for all CSV operations
- **REST API Endpoints**: Created API endpoints for programmatic access
- **MCP Tools**: Added MCP tools for use by Large Language Models

## How to Use

### Via Web UI

1. Navigate to the Excel MCP Server web interface (default: http://localhost:8080)
2. Click on the "CSV" dropdown button next to a file
3. Choose the desired CSV operation:
   - Import CSV: Import a single CSV file into a worksheet
   - Export to CSV: Export a worksheet to CSV
   - Bulk CSV Import: Import multiple CSV files at once
   - Bulk CSV Export: Export multiple worksheets at once

### Via REST API

See detailed API documentation in [CSV_OPERATIONS.md](CSV_OPERATIONS.md).

### Via MCP Tools (for LLMs)

The following MCP tools are available for CSV operations:

```python
# Single operations
import_csv(excel_filename, csv_content_base64, sheet_name, delimiter, create_if_missing, merge_mode)
export_worksheet_to_csv(excel_filename, sheet_name, delimiter, include_header_row)

# Bulk operations
bulk_import_csv(excel_filename, csv_data_list, create_if_missing)
bulk_export_worksheets_to_csv(excel_filename, worksheet_list)
```

## Testing

Two test scripts are provided to validate the bulk CSV operations:

- `test_bulk_csv_operations.py`: Python script for testing bulk CSV operations
- `test_bulk_csv_operations.ps1`: PowerShell script for testing bulk CSV operations on Windows

To run the tests:

```bash
# Python test
python test_bulk_csv_operations.py

# PowerShell test (Windows)
./test_bulk_csv_operations.ps1
```
