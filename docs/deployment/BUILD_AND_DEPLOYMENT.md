# Excel MCP Server - Build, Test and Deployment

This document outlines the build, test, and deployment processes for the Excel MCP Server.

## Development and Testing Tools

The Excel MCP Server project includes several scripts to help with development, testing, and deployment:

### Linting

Run code quality checks using Flake8 and MyPy:

```powershell
# Windows
.\build\lint.ps1

# Linux/macOS
./build/lint.sh
```

### Building Docker Image

Build the Docker image for the Excel MCP Server:

```powershell
# Windows
.\build\build.ps1

# Linux/macOS
./build/build.sh
```

### Running Tests

Execute the test suite to verify server functionality:

```powershell
# Windows
.\build\test.ps1

# Linux/macOS
./build/test.sh
```

The test suite includes:
- Server functionality tests (`test_server.py`)
- File operations tests (`test_file_operations.py`)
- MCP file operations tests (`test_mcp_file_operations.py`)
- CSV operations tests (`test_csv_operations.py`)
- Bulk CSV operations tests (`test_bulk_csv_operations.py`)

### Complete Deployment

For a complete deployment process (lint, build, test, and deploy):

```powershell
# Windows
.\deploy.ps1

# Linux/macOS
./deploy.sh
```

To skip tests during deployment:

```powershell
# Windows
.\deploy.ps1 --skip-tests

# Linux/macOS
./deploy.sh --skip-tests
```

## Docker Compose Integration

The Excel MCP Server is integrated into the main `docker-compose.yml` file with the following configuration:

```yaml
# Excel MCP Server
excel-mcp-server:
  build:
    context: ./excel-mcp-server
    dockerfile: Dockerfile
  container_name: excel-mcp-server
  environment:
    - FASTMCP_PORT=8000
    - FASTAPI_PORT=8080
    - EXCEL_FILES_PATH=/app/excel_files
    - PYTHONUNBUFFERED=1
  ports:
    - "8000:8000"
    - "8080:8080"  volumes:
    - ${EXCEL_FILES_PATH:-./excel-mcp-server/excel_files}:/app/excel_files
    - ./excel-mcp-server/health/health.py:/app/health.py
  healthcheck:
    test: ["CMD", "python", "/app/health.py"]
    interval: 30s
    timeout: 10s
    retries: 3
    start_period: 15s
  networks:
    - mcp-network
  restart: unless-stopped
```

## Health Check System

The Excel MCP Server features a comprehensive health check system (`health.py`) that validates:

1. Socket-level connectivity for both the MCP server (port 8000) and web interface (port 8080)
2. HTTP endpoint verification for the web server's `/health` endpoint
3. API functionality check for the `/api/files` endpoint

Docker's built-in healthcheck uses this script to monitor the service health with the following parameters:
- 30-second check interval
- 10-second timeout
- 3 retries for transient issues
- 15-second startup grace period

## Running Individual Services

To run just the Excel MCP Server using Docker Compose:

```powershell
# From project root
docker-compose up -d excel-mcp-server
```

To stop the service:

```powershell
docker-compose stop excel-mcp-server
```

## Accessing the Services

After deployment, access the services at:
- MCP Server: http://localhost:8000
- Web Interface: http://localhost:8080
