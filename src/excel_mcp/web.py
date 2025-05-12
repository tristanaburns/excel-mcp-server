"""Web interface and API for Excel MCP Server."""

import os
import time
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, Request, Form, Depends
from fastapi.responses import HTMLResponse, RedirectResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

from excel_mcp.server import get_excel_path, EXCEL_FILES_PATH
from excel_mcp.workbook import get_workbook_info as get_info_impl

# Create the FastAPI app - module-level variable
app = FastAPI(
    title="Excel MCP Server",
    description="Web UI and API for Excel MCP Server",
    version="0.1.0",
)

# Export app for external imports
__all__ = ["app"]

# Set up templates and static files
current_dir = Path(__file__).parent
templates = Jinja2Templates(directory=str(current_dir / "templates"))
app.mount("/static", StaticFiles(directory=str(current_dir / "static")), name="static")

class FileInfo(BaseModel):
    """Model for Excel file information."""
    name: str
    path: str
    size: int
    size_formatted: str
    modified: str
    created: str
    
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

def get_excel_files() -> List[FileInfo]:
    """Get a list of all Excel files in the configured directory."""
    files_path = Path(EXCEL_FILES_PATH)
    if not files_path.exists():
        files_path.mkdir(parents=True, exist_ok=True)
        
    files_info = []
    for file_path in files_path.glob("*.xls*"):  # Match both .xls and .xlsx
        if file_path.is_file():
            stats = file_path.stat()
            files_info.append(FileInfo(
                name=file_path.name,
                path=str(file_path),
                size=stats.st_size,
                size_formatted=format_size(stats.st_size),
                modified=datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                created=datetime.fromtimestamp(stats.st_ctime).strftime("%Y-%m-%d %H:%M:%S")
            ))
    
    # Sort by modification time (newest first)
    return sorted(files_info, key=lambda x: x.path, reverse=True)

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Render the home page."""
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/files", response_class=HTMLResponse)
async def file_manager(request: Request, message: Optional[str] = None, message_type: Optional[str] = None):
    """Render the file manager page."""
    files = get_excel_files()
    return templates.TemplateResponse(
        "files.html", 
        {
            "request": request, 
            "files": files, 
            "message": message, 
            "message_type": message_type or "info"
        }
    )

@app.post("/upload")
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel file."""
    # Validate file extension
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
    
    # Save the file
    file_path = Path(EXCEL_FILES_PATH) / file.filename
    
    try:
        # Create the directory if it doesn't exist
        os.makedirs(EXCEL_FILES_PATH, exist_ok=True)
        
        # Write the file
        with open(file_path, "wb") as buffer:
            contents = await file.read()
            buffer.write(contents)
            
        return {"filename": file.filename, "size": len(contents)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to upload file: {str(e)}")

@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download an Excel file."""
    file_path = Path(EXCEL_FILES_PATH) / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/delete/{filename}")
async def delete_file(filename: str):
    """Delete an Excel file."""
    file_path = Path(EXCEL_FILES_PATH) / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    try:
        os.remove(file_path)
        return RedirectResponse(url="/files?message=File deleted successfully&message_type=success")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to delete file: {str(e)}")

@app.get("/api/files")
async def list_files():
    """API endpoint to list all Excel files."""
    files = get_excel_files()
    return {"files": [file.dict() for file in files]}

@app.get("/api/info/{filename}")
async def get_workbook_info(filename: str):
    """API endpoint to get Excel workbook information."""
    file_path = get_excel_path(filename)
    path = Path(file_path)
    
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    try:
        info = get_info_impl(str(file_path))
        return info
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to get workbook info: {str(e)}")

@app.get("/health")
async def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "timestamp": time.time()}

# CSV Operations Models
class CsvImportRequest(BaseModel):
    """Model for CSV import request."""
    csv_content_base64: str
    sheet_name: Optional[str] = None
    delimiter: str = ','
    merge_mode: str = 'replace'
    
class BulkCsvImportRequest(BaseModel):
    """Model for bulk CSV import request."""
    csv_data_list: List[CsvImportRequest]
    create_if_missing: bool = True
    
class CsvExportRequest(BaseModel):
    """Model for CSV export request."""
    sheet_name: Optional[str] = None
    delimiter: str = ','
    include_header_row: bool = True
    
class BulkCsvExportRequest(BaseModel):
    """Model for bulk CSV export request."""
    worksheet_list: List[CsvExportRequest]

# CSV Operations API Endpoints
@app.post("/api/csv/import/{filename}")
async def api_import_csv(filename: str, request: CsvImportRequest):
    """API endpoint to import CSV data into an Excel file."""
    from excel_mcp.server import import_csv
    
    try:
        result = import_csv(
            excel_filename=filename,
            csv_content_base64=request.csv_content_base64,
            sheet_name=request.sheet_name,
            delimiter=request.delimiter,
            create_if_missing=True,
            merge_mode=request.merge_mode
        )
        
        if not result.get("success", False):
            raise HTTPException(status_code=400, detail=result.get("message", "CSV import failed"))
            
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error importing CSV: {str(e)}")

@app.post("/api/csv/export/{filename}")
async def api_export_csv(filename: str, request: CsvExportRequest):
    """API endpoint to export an Excel worksheet to CSV."""
    from excel_mcp.server import export_worksheet_to_csv
    
    try:
        result = export_worksheet_to_csv(
            excel_filename=filename,
            sheet_name=request.sheet_name,
            delimiter=request.delimiter,
            include_header_row=request.include_header_row
        )
        
        if not result.get("success", False):
            raise HTTPException(status_code=400, detail=result.get("message", "CSV export failed"))
            
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error exporting to CSV: {str(e)}")

@app.post("/api/csv/bulk-import/{filename}")
async def api_bulk_import_csv(filename: str, request: BulkCsvImportRequest):
    """API endpoint to import multiple CSV data into an Excel file."""
    from excel_mcp.server import bulk_import_csv
    
    try:
        # Convert the Pydantic model to the format expected by the function
        csv_data_list = []
        for item in request.csv_data_list:
            csv_data_list.append({
                "csv_content_base64": item.csv_content_base64,
                "sheet_name": item.sheet_name,
                "delimiter": item.delimiter,
                "merge_mode": item.merge_mode
            })
            
        result = bulk_import_csv(
            excel_filename=filename,
            csv_data_list=csv_data_list,
            create_if_missing=request.create_if_missing
        )
        
        if not result.get("success", False):
            raise HTTPException(status_code=400, detail=result.get("message", "Bulk CSV import failed"))
            
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error in bulk CSV import: {str(e)}")

@app.post("/api/csv/bulk-export/{filename}")
async def api_bulk_export_csv(filename: str, request: BulkCsvExportRequest):
    """API endpoint to export multiple Excel worksheets to CSV."""
    from excel_mcp.server import bulk_export_worksheets_to_csv
    
    try:
        # Convert the Pydantic model to the format expected by the function
        worksheet_list = []
        for item in request.worksheet_list:
            worksheet_list.append({
                "sheet_name": item.sheet_name,
                "delimiter": item.delimiter,
                "include_header_row": item.include_header_row
            })
            
        result = bulk_export_worksheets_to_csv(
            excel_filename=filename,
            worksheet_list=worksheet_list
        )
        
        if not result.get("success", False):
            raise HTTPException(status_code=400, detail=result.get("message", "Bulk CSV export failed"))
            
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error in bulk CSV export: {str(e)}")
