# File Operations for LLMs

The Excel MCP Server now provides direct file operations through its Server-Sent Events (SSE) interface, allowing Large Language Models (LLMs) to manage files without requiring the web interface. This enables a more streamlined workflow when working with AI assistants.

## New Capabilities

- **List Files**: Get information about available Excel files on the server
- **Upload Files**: Directly upload Excel files using base64 encoding
- **Download Files**: Retrieve Excel files as base64 encoded content
- **Delete Files**: Remove Excel files from the server
- **CSV Import**: Import CSV data directly into Excel workbooks
- **CSV Export**: Extract Excel worksheets as CSV files

## Benefits

1. **Streamlined Workflow**: LLMs can now handle the entire workflow from file creation, manipulation, to deletion all within the MCP interface
2. **Reduced Context Switching**: No need to switch between the MCP interface and web interface for file management
3. **Programmatic Access**: File operations can be fully automated and integrated into AI-driven workflows
4. **Base64 Support**: Files can be easily transferred as text using base64 encoding

## Available Tools

The file operation tools are fully documented in [TOOLS.md](TOOLS.md) and include:

- `list_excel_files()`: List all available Excel files with metadata
- `upload_file(filename, file_content_base64)`: Upload files using base64 encoding
- `download_file(filename)`: Download files as base64 encoded content
- `delete_excel_file(filename)`: Delete files from the server
- `import_csv(excel_filename, csv_content_base64, ...)`: Import CSV data into Excel files
- `export_worksheet_to_csv(excel_filename, sheet_name, ...)`: Export Excel worksheets as CSV

## Example Usage

LLMs can now perform operations like:

1. Creating a new Excel workbook
2. Adding sheets and data
3. Downloading the file (as base64)
4. Modifying the file locally
5. Uploading the modified file back with a new name
6. Listing all files to verify the changes
7. Importing CSV data into Excel workbooks 
8. Exporting worksheet data as CSV files
9. Converting between CSV and Excel formats

All of these operations happen directly through the MCP SSE interface without requiring any manual intervention through the web interface.

### CSV Operations Example

Here's an example workflow for working with CSV data:

1. Import CSV data into a new Excel file:
   ```python
   import_csv(
       excel_filename="sales_data.xlsx",
       csv_content_base64="<base64-encoded-csv-content>",
       sheet_name="Q1_Sales",
       create_if_missing=True
   )
   ```

2. Export a worksheet as CSV:
   ```python
   export_result = export_worksheet_to_csv(
       excel_filename="sales_data.xlsx",
       sheet_name="Q1_Sales"
   )
   csv_content = export_result["csv_content_base64"]  # Base64-encoded CSV
   ```

3. Merge multiple CSV files into one Excel workbook:
   ```python
   # First CSV file creates the workbook
   import_csv(
       excel_filename="annual_report.xlsx",
       csv_content_base64="<q1-sales-data>",
       sheet_name="Q1"
   )
   
   # Second CSV goes to a new sheet
   import_csv(
       excel_filename="annual_report.xlsx", 
       csv_content_base64="<q2-sales-data>",
       sheet_name="Q2"
   )
   ```

## Testing

Test scripts are provided to verify the file operations functionality:

```bash
python test_file_operations.py  # Test basic file operations
python test_csv_operations.py   # Test CSV import/export operations
```

For PowerShell users:

```powershell
.\test_csv_operations.ps1  # Test CSV operations in PowerShell
```

These scripts demonstrate the complete workflow of creating, downloading, deleting, and re-uploading Excel files, as well as importing and exporting CSV data through the MCP interface.
