# CSV Operations in Excel MCP Server

This document describes the CSV import and export functionality in Excel MCP Server, including the new bulk operations capabilities.

## CSV Operations Overview

The Excel MCP Server now provides comprehensive functionality for working with CSV files:

- **Single CSV Import**: Import a CSV file into a specific worksheet
- **Single CSV Export**: Export a worksheet as CSV
- **Bulk CSV Import**: Import multiple CSV files into different worksheets in one operation
- **Bulk CSV Export**: Export multiple worksheets as separate CSV files in one operation

## Using the API

### Single CSV Operations

#### Import CSV

```http
POST /api/csv/import/{excel_filename}
```

Request body:
```json
{
  "csv_content_base64": "base64-encoded-csv-content",
  "sheet_name": "Sheet1",  // optional, uses active sheet if not provided
  "delimiter": ",",        // optional, defaults to comma
  "merge_mode": "replace"  // "replace", "append", or "new_sheet"
}
```

Response:
```json
{
  "success": true,
  "message": "Imported CSV data to Sheet1 in example.xlsx",
  "rows_imported": 100,
  "columns_imported": 5,
  "sheet_name": "Sheet1",
  "excel_path": "path/to/excel/file.xlsx"
}
```

#### Export CSV

```http
POST /api/csv/export/{excel_filename}
```

Request body:
```json
{
  "sheet_name": "Sheet1",       // optional, uses active sheet if not provided
  "delimiter": ",",             // optional, defaults to comma
  "include_header_row": true    // optional, defaults to true
}
```

Response:
```json
{
  "success": true,
  "message": "Exported worksheet Sheet1 from example.xlsx",
  "csv_content_base64": "base64-encoded-csv-content",
  "rows_exported": 100,
  "sheet_name": "Sheet1"
}
```

### Bulk CSV Operations

#### Bulk CSV Import

```http
POST /api/csv/bulk-import/{excel_filename}
```

Request body:
```json
{
  "csv_data_list": [
    {
      "csv_content_base64": "base64-encoded-csv-content-1",
      "sheet_name": "Sheet1",
      "delimiter": ",",
      "merge_mode": "replace"
    },
    {
      "csv_content_base64": "base64-encoded-csv-content-2",
      "sheet_name": "Sheet2",
      "delimiter": ",",
      "merge_mode": "new_sheet"
    }
  ],
  "create_if_missing": true
}
```

Response:
```json
{
  "success": true,
  "message": "Processed 2 CSV imports (2 successful, 0 failed)",
  "results": [
    {
      "success": true,
      "sheet_name": "Sheet1",
      "rows_imported": 100,
      "columns_imported": 5,
      "message": "Imported CSV data to Sheet1 in example.xlsx"
    },
    {
      "success": true,
      "sheet_name": "Sheet2",
      "rows_imported": 50,
      "columns_imported": 3,
      "message": "Imported CSV data to Sheet2 in example.xlsx"
    }
  ],
  "excel_path": "path/to/excel/file.xlsx"
}
```

#### Bulk CSV Export

```http
POST /api/csv/bulk-export/{excel_filename}
```

Request body:
```json
{
  "worksheet_list": [
    {
      "sheet_name": "Sheet1",
      "delimiter": ",",
      "include_header_row": true
    },
    {
      "sheet_name": "Sheet2",
      "delimiter": ";",
      "include_header_row": false
    }
  ]
}
```

Response:
```json
{
  "success": true,
  "message": "Processed 2 CSV exports (2 successful, 0 failed)",
  "results": [
    {
      "success": true,
      "sheet_name": "Sheet1",
      "rows_exported": 100,
      "columns_exported": 5,
      "csv_content_base64": "base64-encoded-csv-content-1",
      "message": "Exported worksheet Sheet1 from example.xlsx"
    },
    {
      "success": true,
      "sheet_name": "Sheet2",
      "rows_exported": 50,
      "columns_exported": 3,
      "csv_content_base64": "base64-encoded-csv-content-2",
      "message": "Exported worksheet Sheet2 from example.xlsx"
    }
  ]
}
```

## Using the MCP Tools

The Excel MCP Server provides several tools for working with CSV files through the Model Context Protocol.

### Single CSV Operations

#### import_csv

Imports a CSV file into an Excel workbook.

```python
import_csv(
    excel_filename: str,
    csv_content_base64: str,
    sheet_name: str = None,
    delimiter: str = ',',
    create_if_missing: bool = True,
    merge_mode: str = 'replace'
)
```

#### export_worksheet_to_csv

Exports an Excel worksheet to CSV format.

```python
export_worksheet_to_csv(
    excel_filename: str,
    sheet_name: str = None,
    delimiter: str = ',',
    include_header_row: bool = True
)
```

### Bulk CSV Operations

#### bulk_import_csv

Imports multiple CSV files into an Excel workbook.

```python
bulk_import_csv(
    excel_filename: str,
    csv_data_list: List[Dict[str, Any]],
    create_if_missing: bool = True
)
```

#### bulk_export_worksheets_to_csv

Exports multiple Excel worksheets as CSV files.

```python
bulk_export_worksheets_to_csv(
    excel_filename: str,
    worksheet_list: List[Dict[str, Any]]
)
```

## Using the Web UI

The Excel MCP Server web interface now provides a user-friendly way to work with CSV files:

1. Access the file manager at `/files`
2. For each Excel file, you'll see a CSV dropdown menu with the following options:
   - **Import CSV**: Import a single CSV file into a worksheet
   - **Export to CSV**: Export a worksheet to CSV format
   - **Bulk CSV Import**: Import multiple CSV files into different worksheets
   - **Bulk CSV Export**: Export multiple worksheets as CSV files

### Import CSV

1. Click "Import CSV" from the CSV dropdown menu
2. Select a CSV file to import
3. Specify the sheet name (optional), delimiter, and merge mode
4. Click "Import"

### Export CSV

1. Click "Export to CSV" from the CSV dropdown menu
2. Specify the sheet name (optional), delimiter, and whether to include headers
3. Click "Export"
4. Download the exported CSV file

### Bulk CSV Import

1. Click "Bulk CSV Import" from the CSV dropdown menu
2. Add one or more CSV files, specifying sheet name, delimiter, and merge mode for each
3. Click "Import All"

### Bulk CSV Export

1. Click "Bulk CSV Export" from the CSV dropdown menu
2. Click "Get Worksheets" to retrieve the available worksheets
3. Select the worksheets to export and configure delimiter and header options
4. Click "Export Selected"
5. Download the exported CSV files
