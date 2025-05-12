import logging
import sys
import os
import time
import base64
from datetime import datetime
from pathlib import Path
from typing import Any, List, Dict, Optional

from mcp.server.fastmcp import FastMCP

# Import exceptions
from excel_mcp.exceptions import (
    ValidationError,
    WorkbookError,
    SheetError,
    DataError,
    FormattingError,
    CalculationError,
    PivotError,
    ChartError,
    FileOperationError
)

# Import from excel_mcp package with consistent _impl suffixes
from excel_mcp.validation import (
    validate_formula_in_cell_operation as validate_formula_impl,
    validate_range_in_sheet_operation as validate_range_impl
)
from excel_mcp.chart import create_chart_in_sheet as create_chart_impl
from excel_mcp.workbook import get_workbook_info
from excel_mcp.data import write_data
from excel_mcp.pivot import create_pivot_table as create_pivot_table_impl
from excel_mcp.sheet import (
    copy_sheet,
    delete_sheet,
    rename_sheet,
    merge_range,
    unmerge_range,
)
from excel_mcp.csv_operations import (
    import_csv_file_base64,
    export_worksheet_to_csv_base64,
    bulk_import_csv_file_base64,
    bulk_export_worksheets_to_csv_base64
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("excel-mcp.log")
    ],
    force=True
)

logger = logging.getLogger("excel-mcp")

# Get Excel files path from environment or use default
EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")

# Create the directory if it doesn't exist
os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

# Initialize FastMCP server
mcp = FastMCP(
    "excel-mcp",
    version="0.1.1",
    description="Excel MCP Server for manipulating Excel files",
    dependencies=["openpyxl>=3.1.2"],
    env_vars={
        "EXCEL_FILES_PATH": {
            "description": "Path to Excel files directory",
            "required": False,
            "default": EXCEL_FILES_PATH
        }
    }
)

def get_excel_path(filename: str) -> str:
    """Get full path to Excel file.
    
    Args:
        filename: Name of Excel file
        
    Returns:
        Full path to Excel file
    """
    # If filename is already an absolute path, return it
    if os.path.isabs(filename):
        return filename
        
    # Use the configured Excel files path
    return os.path.join(EXCEL_FILES_PATH, filename)

@mcp.tool()
def apply_formula(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> str:
    """
    Apply Excel formula to cell.
    Excel formula will write to cell with verification.
    """
    try:
        full_path = get_excel_path(filepath)
        # First validate the formula
        validation = validate_formula_impl(full_path, sheet_name, cell, formula)
        if isinstance(validation, dict) and "error" in validation:
            return f"Error: {validation['error']}"
            
        # If valid, apply the formula
        from excel_mcp.calculations import apply_formula as apply_formula_impl
        result = apply_formula_impl(full_path, sheet_name, cell, formula)
        return result["message"]
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error applying formula: {e}")
        raise

@mcp.tool()
def validate_formula_syntax(
    filepath: str,
    sheet_name: str,
    cell: str,
    formula: str,
) -> str:
    """Validate Excel formula syntax without applying it."""
    try:
        full_path = get_excel_path(filepath)
        result = validate_formula_impl(full_path, sheet_name, cell, formula)
        return result["message"]
    except (ValidationError, CalculationError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating formula: {e}")
        raise

@mcp.tool()
def format_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_size: int = None,
    font_color: str = None,
    bg_color: str = None,
    border_style: str = None,
    border_color: str = None,
    number_format: str = None,
    alignment: str = None,
    wrap_text: bool = False,
    merge_cells: bool = False,
    protection: Dict[str, Any] = None,
    conditional_format: Dict[str, Any] = None
) -> str:
    """Apply formatting to a range of cells."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.formatting import format_range as format_range_func
        
        result = format_range_func(
            filepath=full_path,
            sheet_name=sheet_name,
            start_cell=start_cell,
            end_cell=end_cell,
            bold=bold,
            italic=italic,
            underline=underline,
            font_size=font_size,
            font_color=font_color,
            bg_color=bg_color,
            border_style=border_style,
            border_color=border_color,
            number_format=number_format,
            alignment=alignment,
            wrap_text=wrap_text,
            merge_cells=merge_cells,
            protection=protection,
            conditional_format=conditional_format
        )
        return "Range formatted successfully"
    except (ValidationError, FormattingError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error formatting range: {e}")
        raise

@mcp.tool()
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    preview_only: bool = False
) -> str:
    """
    Read data from Excel worksheet.
    
    Returns:  
    Data from Excel worksheet as json string. list of lists or empty list if no data found. sublists are assumed to be rows.
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.data import read_excel_range
        result = read_excel_range(full_path, sheet_name, start_cell, end_cell, preview_only)
        if not result:
            return "No data found in specified range"
        # Convert the list of dicts to a formatted string
        data_str = "\n".join([str(row) for row in result])
        return data_str
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise

@mcp.tool()
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1",
) -> str:
    """
    Write data to Excel worksheet.
    Excel formula will write to cell without any verification.

    PARAMETERS:  
    filepath: Path to Excel file
    sheet_name: Name of worksheet to write to
    data: List of lists containing data to write to the worksheet, sublists are assumed to be rows
    start_cell: Cell to start writing to, default is "A1"
  
    """
    try:
        full_path = get_excel_path(filepath)
        result = write_data(full_path, sheet_name, data, start_cell)
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise

@mcp.tool()
def create_workbook(filepath: str) -> str:
    """Create new Excel workbook."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_workbook as create_workbook_impl
        result = create_workbook_impl(full_path)
        return f"Created workbook at {full_path}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise

@mcp.tool()
def create_worksheet(filepath: str, sheet_name: str) -> str:
    """Create new worksheet in workbook."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_sheet as create_worksheet_impl
        result = create_worksheet_impl(full_path, sheet_name)
        return result["message"]
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating worksheet: {e}")
        raise

@mcp.tool()
def create_chart(
    filepath: str,
    sheet_name: str,
    data_range: str,
    chart_type: str,
    target_cell: str,
    title: str = "",
    x_axis: str = "",
    y_axis: str = ""
) -> str:
    """Create chart in worksheet."""
    try:
        full_path = get_excel_path(filepath)
        result = create_chart_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            data_range=data_range,
            chart_type=chart_type,
            target_cell=target_cell,
            title=title,
            x_axis=x_axis,
            y_axis=y_axis
        )
        return result["message"]
    except (ValidationError, ChartError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating chart: {e}")
        raise

@mcp.tool()
def create_pivot_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    rows: List[str],
    values: List[str],
    columns: List[str] = None,
    agg_func: str = "mean"
) -> str:
    """Create pivot table in worksheet."""
    try:
        full_path = get_excel_path(filepath)
        result = create_pivot_table_impl(
            filepath=full_path,
            sheet_name=sheet_name,
            data_range=data_range,
            rows=rows,
            values=values,
            columns=columns or [],
            agg_func=agg_func
        )
        return result["message"]
    except (ValidationError, PivotError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating pivot table: {e}")
        raise

@mcp.tool()
def copy_worksheet(
    filepath: str,
    source_sheet: str,
    target_sheet: str
) -> str:
    """Copy worksheet within workbook."""
    try:
        full_path = get_excel_path(filepath)
        result = copy_sheet(full_path, source_sheet, target_sheet)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying worksheet: {e}")
        raise

@mcp.tool()
def delete_worksheet(
    filepath: str,
    sheet_name: str
) -> str:
    """Delete worksheet from workbook."""
    try:
        full_path = get_excel_path(filepath)
        result = delete_sheet(full_path, sheet_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting worksheet: {e}")
        raise

@mcp.tool()
def rename_worksheet(
    filepath: str,
    old_name: str,
    new_name: str
) -> str:
    """Rename worksheet in workbook."""
    try:
        full_path = get_excel_path(filepath)
        result = rename_sheet(full_path, old_name, new_name)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error renaming worksheet: {e}")
        raise

@mcp.tool()
def get_workbook_metadata(
    filepath: str,
    include_ranges: bool = False
) -> str:
    """Get metadata about workbook including sheets, ranges, etc."""
    try:
        full_path = get_excel_path(filepath)
        result = get_workbook_info(full_path, include_ranges=include_ranges)
        return str(result)
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise

@mcp.tool()
def merge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Merge a range of cells."""
    try:
        full_path = get_excel_path(filepath)
        result = merge_range(full_path, sheet_name, start_cell, end_cell)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error merging cells: {e}")
        raise

@mcp.tool()
def unmerge_cells(filepath: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """Unmerge a range of cells."""
    try:
        full_path = get_excel_path(filepath)
        result = unmerge_range(full_path, sheet_name, start_cell, end_cell)
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error unmerging cells: {e}")
        raise

@mcp.tool()
def copy_range(
    filepath: str,
    sheet_name: str,
    source_start: str,
    source_end: str,
    target_start: str,
    target_sheet: str = None
) -> str:
    """Copy a range of cells to another location."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.sheet import copy_range_operation
        result = copy_range_operation(
            full_path,
            sheet_name,
            source_start,
            source_end,
            target_start,
            target_sheet
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error copying range: {e}")
        raise

@mcp.tool()
def delete_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str,
    shift_direction: str = "up"
) -> str:
    """Delete a range of cells and shift remaining cells."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.sheet import delete_range_operation
        result = delete_range_operation(
            full_path,
            sheet_name,
            start_cell,
            end_cell,
            shift_direction
        )
        return result["message"]
    except (ValidationError, SheetError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error deleting range: {e}")
        raise

@mcp.tool()
def validate_excel_range(
    filepath: str,
    sheet_name: str,
    start_cell: str,
    end_cell: str = None
) -> str:
    """Validate if a range exists and is properly formatted."""
    try:
        full_path = get_excel_path(filepath)
        range_str = start_cell if not end_cell else f"{start_cell}:{end_cell}"
        result = validate_range_impl(full_path, sheet_name, range_str)
        return result["message"]
    except ValidationError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error validating range: {e}")
        raise

@mcp.tool()
def health_check():
    """Health check endpoint for the MCP server."""
    return {"status": "healthy", "timestamp": time.time(), "service": "mcp-server"}

# Helper function for formatting file sizes
def format_size(size_bytes):
    """Format file size in a human-readable format."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

@mcp.tool()
def list_excel_files() -> Dict[str, Any]:
    """
    List all Excel files available on the server.
    Returns information about files including name, size, and modification date.
    """
    try:
        files_path = Path(EXCEL_FILES_PATH)
        if not files_path.exists():
            files_path.mkdir(parents=True, exist_ok=True)
            
        files_info = []
        for file_path in files_path.glob("*.xls*"):  # Match both .xls and .xlsx
            if file_path.is_file():
                stats = file_path.stat()
                files_info.append({
                    "name": file_path.name,
                    "path": str(file_path),
                    "size": stats.st_size,
                    "size_formatted": format_size(stats.st_size),
                    "modified": datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                    "created": datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d %H:%M:%S")
                })
        
        # Sort files alphabetically by name
        return {
            "success": True,
            "files": sorted(files_info, key=lambda x: x["name"]),
            "message": f"Found {len(files_info)} Excel files"
        }
    except Exception as e:
        logger.error(f"Error listing Excel files: {e}")
        return {
            "success": False,
            "files": [],
            "message": f"Failed to list files: {str(e)}"
        }

@mcp.tool()
def upload_file(filename: str, file_content_base64: str) -> Dict[str, Any]:
    """
    Upload an Excel file to the server using base64 encoding.
    
    Args:
        filename: Name of the file to create
        file_content_base64: Base64 encoded file content
        
    Returns:
        Dictionary with upload status information
    """
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return {
            "success": False,
            "message": "Invalid file extension. Only .xlsx and .xls files are supported."
        }
    
    try:
        # Decode base64 content
        file_content = base64.b64decode(file_content_base64)
        
        # Create target directory if it doesn't exist
        os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
        
        # Save file
        file_path = Path(EXCEL_FILES_PATH) / filename
        with open(file_path, "wb") as f:
            f.write(file_content)
            
        file_size = file_path.stat().st_size
            
        return {
            "success": True,
            "message": f"File {filename} uploaded successfully ({format_size(file_size)})",
            "file_path": str(file_path),
            "size": file_size,
            "size_formatted": format_size(file_size)
        }
    except Exception as e:
        logger.error(f"Error uploading file: {e}")
        return {
            "success": False,
            "message": f"Failed to upload file: {str(e)}"
        }

@mcp.tool()
def download_file(filename: str) -> Dict[str, Any]:
    """
    Download an Excel file from the server as base64 encoded content.
    
    Args:
        filename: Name of the file to download
        
    Returns:
        Dictionary with file content (base64 encoded) and metadata
    """
    try:
        file_path = Path(EXCEL_FILES_PATH) / filename
        
        if not file_path.exists():
            return {
                "success": False,
                "message": f"File {filename} not found"
            }
        
        # Read file content
        with open(file_path, "rb") as f:
            file_content = f.read()
        
        # Get file stats
        stats = file_path.stat()
        
        return {
            "success": True,
            "filename": filename,
            "content_base64": base64.b64encode(file_content).decode('utf-8'),
            "size": stats.st_size,
            "size_formatted": format_size(stats.st_size),
            "modified": datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "created": datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
            "message": f"Downloaded {filename} ({format_size(stats.st_size)})"
        }
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        return {
            "success": False,
            "message": f"Failed to download file: {str(e)}"
        }

@mcp.tool()
def delete_excel_file(filename: str) -> Dict[str, Any]:
    """
    Delete an Excel file from the server.
    
    Args:
        filename: Name of the file to delete
        
    Returns:
        Dictionary with deletion status
    """
    try:
        file_path = Path(EXCEL_FILES_PATH) / filename
        
        if not file_path.exists():
            return {
                "success": False,
                "message": f"File {filename} not found"
            }
        
        # Delete the file
        os.remove(file_path)
        
        return {
            "success": True,
            "message": f"File {filename} deleted successfully"
        }
    except Exception as e:
        logger.error(f"Error deleting file: {e}")
        return {
            "success": False,
            "message": f"Failed to delete file: {str(e)}"
        }
        
@mcp.tool()
def import_csv(
    excel_filename: str,
    csv_content_base64: str,
    sheet_name: str = None,
    delimiter: str = ',',
    create_if_missing: bool = True,
    merge_mode: str = 'replace'
) -> Dict[str, Any]:
    """
    Import a CSV file into an Excel workbook.
    
    Args:
        excel_filename: Name of the Excel file to create or modify
        csv_content_base64: Base64 encoded CSV content
        sheet_name: Name of the worksheet to import into (created if doesn't exist)
        delimiter: CSV delimiter character (default: ',')
        create_if_missing: Whether to create the Excel file if it doesn't exist
        merge_mode: How to handle existing data ('replace', 'append', 'new_sheet')
    
    Returns:
        Dictionary with information about the import operation
    """
    try:
        excel_path = get_excel_path(excel_filename)
        
        # Import the CSV
        result = import_csv_file_base64(
            excel_path,
            csv_content_base64,
            sheet_name,
            delimiter,
            create_if_missing,
            merge_mode
        )
        
        return {
            "success": True,
            "message": result["message"],
            "rows_imported": result["rows_imported"],
            "columns_imported": result.get("columns_imported", 0),
            "sheet_name": result["sheet_name"],
            "excel_path": result["filepath"]
        }
    except FileOperationError as e:
        logger.error(f"Error importing CSV: {e}")
        return {
            "success": False,
            "message": f"CSV import failed: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Unexpected error importing CSV: {e}")
        return {
            "success": False,
            "message": f"Unexpected error during CSV import: {str(e)}"
        }

@mcp.tool()
def export_worksheet_to_csv(
    excel_filename: str,
    sheet_name: str = None,
    delimiter: str = ',',
    include_header_row: bool = True
) -> Dict[str, Any]:
    """
    Export an Excel worksheet as a CSV file.
    
    Args:
        excel_filename: Name of the Excel file
        sheet_name: Name of the worksheet to export (uses active sheet if None)
        delimiter: CSV delimiter character (default: ',')
        include_header_row: Whether to treat the first row as headers
    
    Returns:
        Dictionary with CSV content and export information
    """
    try:
        excel_path = get_excel_path(excel_filename)
        
        # Export the worksheet to CSV
        result = export_worksheet_to_csv_base64(
            excel_path,
            sheet_name,
            delimiter,
            include_header_row
        )
        
        return {
            "success": True,
            "message": result["message"],
            "csv_content_base64": result["csv_content_base64"],
            "rows_exported": result["rows_exported"],
            "sheet_name": result["sheet_name"]
        }
    except FileOperationError as e:
        logger.error(f"Error exporting to CSV: {e}")
        return {
            "success": False,
            "message": f"CSV export failed: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Unexpected error exporting to CSV: {e}")
        return {
            "success": False,
            "message": f"Unexpected error during CSV export: {str(e)}"
        }

@mcp.tool()
def bulk_import_csv(
    excel_filename: str,
    csv_data_list: List[Dict[str, Any]],
    create_if_missing: bool = True
) -> Dict[str, Any]:
    """
    Import multiple CSV files into an Excel workbook.
    
    Args:
        excel_filename: Name of the Excel file to create or modify
        csv_data_list: List of dictionaries containing CSV information:
            - csv_content_base64: Base64 encoded content of the CSV file
            - sheet_name: Name of worksheet to import into (created if doesn't exist)
            - delimiter: CSV delimiter character (default: ',')
            - merge_mode: How to handle existing data ('replace', 'append', 'new_sheet')
        create_if_missing: Whether to create the Excel file if it doesn't exist
    
    Returns:
        Dictionary with information about the bulk import operation
    """
    try:
        excel_path = get_excel_path(excel_filename)
        
        # Import the CSV files
        results = bulk_import_csv_file_base64(
            excel_path,
            csv_data_list,
            create_if_missing
        )
        
        # Count successful imports
        successful = sum(1 for result in results if result.get("success", False))
        total = len(results)
        
        return {
            "success": True,
            "message": f"Processed {total} CSV imports ({successful} successful, {total - successful} failed)",
            "results": results,
            "excel_path": str(excel_path)
        }
    except FileOperationError as e:
        logger.error(f"Error in bulk CSV import: {e}")
        return {
            "success": False,
            "message": f"Bulk CSV import failed: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Unexpected error in bulk CSV import: {e}")
        return {
            "success": False,
            "message": f"Unexpected error during bulk CSV import: {str(e)}"
        }

@mcp.tool()
def bulk_export_worksheets_to_csv(
    excel_filename: str,
    worksheet_list: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """
    Export multiple Excel worksheets as CSV files.
    
    Args:
        excel_filename: Name of the Excel file
        worksheet_list: List of dictionaries containing worksheet information:
            - sheet_name: Name of worksheet to export (uses active sheet if None)
            - delimiter: CSV delimiter character (default: ',')
            - include_header_row: Whether to treat the first row as headers
    
    Returns:
        Dictionary with CSV content and export information for each worksheet
    """
    try:
        excel_path = get_excel_path(excel_filename)
        
        # Export the worksheets to CSV
        results = bulk_export_worksheets_to_csv_base64(
            excel_path,
            worksheet_list
        )
        
        # Count successful exports
        successful = sum(1 for result in results if result.get("success", False))
        total = len(results)
        
        return {
            "success": True,
            "message": f"Processed {total} CSV exports ({successful} successful, {total - successful} failed)",
            "results": results
        }
    except FileOperationError as e:
        logger.error(f"Error in bulk CSV export: {e}")
        return {
            "success": False,
            "message": f"Bulk CSV export failed: {str(e)}"
        }
    except Exception as e:
        logger.error(f"Unexpected error in bulk CSV export: {e}")
        return {
            "success": False,
            "message": f"Unexpected error during bulk CSV export: {str(e)}"
        }

async def run_server():
    """Run the Excel MCP server."""
    try:
        # Log server startup info
        logger.info(f"Starting Excel MCP server on port {os.environ.get('FASTMCP_PORT', '8000')}")
        logger.info(f"Excel files directory: {EXCEL_FILES_PATH}")

        # Make sure Excel files directory exists with proper permissions
        os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
        logger.info(f"Ensuring Excel files directory exists at {EXCEL_FILES_PATH}")
        
        # Start the MCP server
        await mcp.run_sse_async()
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
        await mcp.shutdown()
    except Exception as e:
        logger.error(f"Server failed: {e}")
        logger.exception("Exception details:")
        raise
    finally:
        logger.info("Server shutdown complete")