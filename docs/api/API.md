# Excel MCP Server API Documentation

This document provides comprehensive documentation for both the web API interface and the Model Context Protocol (MCP) Server-Sent Events (SSE) interface of the Excel MCP Server.

## Web API Interface

The web API is accessible at port 8080 by default and provides RESTful endpoints for file management and Excel information retrieval.

### Base URL

```
http://localhost:8080
```

### Endpoints

#### Health Check

```
GET /health
```

Returns the health status of the server.

**Response:**
```json
{
  "status": "healthy",
  "timestamp": 1747034092.5410728
}
```

#### List Files

```
GET /api/files
```

Returns a list of all Excel files available on the server.

**Response:**
```json
{
  "files": [
    {
      "name": "example.xlsx",
      "path": "/app/excel_files/example.xlsx",
      "size": 12345,
      "size_formatted": "12.06 KB",
      "modified": "2025-05-12 07:19:20",
      "created": "2025-05-12 07:19:20"
    }
  ]
}
```

#### Get Workbook Info

```
GET /api/info/{filename}
```

Returns metadata information about an Excel workbook.

**Parameters:**
- `filename`: Name of the Excel file

**Response:**
```json
{
  "sheets": ["Sheet1", "Sheet2"],
  "activeSheet": "Sheet1",
  "created": "2025-05-12 07:19:20",
  "modified": "2025-05-12 07:19:20"
}
```

#### Download File

```
GET /download/{filename}
```

Downloads an Excel file.

**Parameters:**
- `filename`: Name of the Excel file

**Response:**
- File download (Excel file)

#### Upload File

```
POST /upload
```

Uploads an Excel file.

**Form Data:**
- `file`: Excel file to upload (.xlsx or .xls)

**Response:**
```json
{
  "filename": "example.xlsx",
  "size": 12345
}
```

#### Delete File

```
GET /delete/{filename}
```

Deletes an Excel file.

**Parameters:**
- `filename`: Name of the Excel file to delete

**Response:**
- Redirect to `/files` page with success message

## MCP SSE Interface

The MCP interface is accessible at port 8000 by default and uses Server-Sent Events (SSE) for communication. This allows AI assistants to directly interact with Excel files and perform operations.

### Base URL

```
http://localhost:8000/sse
```

### Tools

#### Excel File Management Tools

##### list_excel_files

Lists all Excel files available on the server with metadata.

**Request:**
```json
{
  "id": "request_id",
  "type": "function",
  "function": {
    "name": "list_excel_files",
    "arguments": "{}"
  }
}
```

**Response:**
```json
{
  "success": true,
  "files": [
    {
      "name": "example.xlsx",
      "path": "/app/excel_files/example.xlsx",
      "size": 12345,
      "size_formatted": "12.06 KB",
      "modified": "2025-05-12 07:19:20",
      "created": "2025-05-12 07:19:20"
    }
  ],
  "message": "Found 1 Excel files"
}
```

##### upload_file

Uploads an Excel file to the server using base64 encoding.

**Request:**
```json
{
  "id": "request_id",
  "type": "function",
  "function": {
    "name": "upload_file",
    "arguments": "{\"filename\": \"example.xlsx\", \"file_content_base64\": \"UEsDBBQABg...\"}"
  }
}
```

**Parameters:**
- `filename`: Name of the file to create
- `file_content_base64`: Base64 encoded file content

**Response:**
```json
{
  "success": true,
  "message": "File example.xlsx uploaded successfully (12.06 KB)",
  "file_path": "/app/excel_files/example.xlsx",
  "size": 12345,
  "size_formatted": "12.06 KB"
}
```

##### download_file

Downloads an Excel file from the server as base64 encoded content.

**Request:**
```json
{
  "id": "request_id",
  "type": "function",
  "function": {
    "name": "download_file",
    "arguments": "{\"filename\": \"example.xlsx\"}"
  }
}
```

**Parameters:**
- `filename`: Name of the file to download

**Response:**
```json
{
  "success": true,
  "filename": "example.xlsx",
  "content_base64": "UEsDBBQABg...",
  "size": 12345,
  "size_formatted": "12.06 KB",
  "modified": "2025-05-12 07:19:20",
  "created": "2025-05-12 07:19:20",
  "message": "Downloaded example.xlsx (12.06 KB)"
}
```

##### delete_excel_file

Deletes an Excel file from the server.

**Request:**
```json
{
  "id": "request_id",
  "type": "function",
  "function": {
    "name": "delete_excel_file",
    "arguments": "{\"filename\": \"example.xlsx\"}"
  }
}
```

**Parameters:**
- `filename`: Name of the file to delete

**Response:**
```json
{
  "success": true,
  "message": "File example.xlsx deleted successfully"
}
```

#### Excel Manipulation Tools

For a complete list of Excel manipulation tools available through the MCP SSE interface, refer to the [TOOLS.md](./TOOLS.md) documentation.

## Error Handling

Both interfaces return appropriate error responses when operations fail:

### Web API Errors

Web API errors are returned as HTTP status codes with error details:

```json
{
  "detail": "Error message describing the issue"
}
```

### MCP SSE Errors

MCP SSE errors are returned in the response object with success set to false:

```json
{
  "success": false,
  "message": "Error description"
}
```

## Authentication

Currently, both interfaces are unauthenticated. Future versions may implement authentication mechanisms for improved security.

## Rate Limiting

No rate limiting is currently implemented. However, excessive requests may impact server performance.

## Conclusion

The Excel MCP Server provides two complementary interfaces:
1. A web API for human users to manage Excel files through a browser
2. An MCP SSE interface for AI assistants to directly interact with Excel files

Both interfaces share the same underlying functionality but present it in different ways suitable for their respective consumers.
