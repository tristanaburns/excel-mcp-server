# Excel MCP Server

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 Create and modify Excel workbooks
- 📝 Read and write data
- 🎨 Apply formatting and styles
- 📈 Create charts and visualizations
- 📊 Generate pivot tables
- 🔄 Manage worksheets and ranges
- 📄 Import and export CSV files
- 📂 File upload/download via SSE interface (for LLMs)
- 🌐 Web interface for file management
- 🔌 REST API for programmatic access

## Quick Start

### Prerequisites

- Python 3.10 or higher

### Installation

1. Clone the repository:
```bash
git clone https://github.com/haris-musa/excel-mcp-server.git
cd excel-mcp-server
```

2. Install using uv:
```bash
uv pip install -e .
```

### Running the Server

Start the server (default ports: 8000 for MCP, 8080 for web interface):
```bash
uv run excel-mcp-server
```

Custom ports:

```bash
# Bash/Linux/macOS
export FASTMCP_PORT=8000 FASTAPI_PORT=8080 && uv run excel-mcp-server

# Windows PowerShell
$env:FASTMCP_PORT = "8000"; $env:FASTAPI_PORT = "8080"; uv run excel-mcp-server
```

### Web Interface

The Excel MCP Server includes a web interface for managing Excel files:

- **URL**: http://localhost:8080 (default port)
- **Features**:
  - Upload Excel files
  - Download existing files
  - Delete files
  - View file metadata
  - File browser interface

![Web Interface](https://i.imgur.com/example.png)

#### Web Interface Endpoints

- `GET /` - Home page
- `GET /files` - File manager page
- `POST /upload` - Upload Excel files
- `GET /download/{filename}` - Download file
- `GET /delete/{filename}` - Delete file

## Using with AI Tools

### Cursor IDE

1. Add this configuration to Cursor:
```json
{
  "mcpServers": {
    "excel": {
      "url": "http://localhost:8000/sse",
      "env": {
        "EXCEL_FILES_PATH": "/path/to/excel/files"
      }
    }
  }
}
```

## Docker Deployment

### Prerequisites

- Docker and Docker Compose installed

### Build and Deployment Scripts

For convenient development and deployment, several scripts are provided:

#### Linting

```powershell
# Windows
.\build\lint.ps1

# Linux/macOS
./build/lint.sh
```

#### Building Docker Image

```powershell
# Windows
.\build\build.ps1

# Linux/macOS
./build/build.sh
```

#### Running Tests

```powershell
# Windows
.\build\test.ps1

# Linux/macOS
./build/test.sh
```

#### Complete Deployment

```powershell
# Windows
.\build\deploy.ps1

# Linux/macOS
./build/deploy.sh
```

For more detailed information, see [docs/deployment/BUILD_AND_DEPLOYMENT.md](docs/deployment/BUILD_AND_DEPLOYMENT.md).

### Using Docker

1. Build and start the container:
```bash
# From the mcp-servers root directory
docker-compose up -d excel-mcp-server

# To build or rebuild the image first
docker-compose build excel-mcp-server
docker-compose up -d excel-mcp-server
```

2. Stop the container:
```bash
npm run docker:excel-mcp-server:down
```

3. View logs:
```bash
npm run docker:excel-mcp-server:logs
```

### Manual Docker Commands

```bash
# Build the image
docker build -t excel-mcp-server ./excel-mcp-server

# Run the container with a custom local directory for Excel files
# Replace C:/path/to/excel/files with your actual local directory path
docker run -p 8000:8000 -p 8080:8080 -v C:/path/to/excel/files:/app/excel_files -e PYTHONUNBUFFERED=1 excel-mcp-server
```

This will start both:
- MCP Server on port 8000 for AI tools
- Web interface on port 8080 for file management

The Excel tools will be available through your AI assistant while the web interface will be accessible at http://localhost:8080 through your browser.

### Container Health Checks

The Docker container includes a comprehensive health check system that validates both services:

- Validates port availability for both the MCP server (8000) and Web interface (8080)
- Verifies the Web interface `/health` endpoint is responding
- Checks API functionality through the `/api/files` endpoint
- Automatically restarts the container if health checks fail

Health check configuration:
```yaml
healthcheck:
  test: ["CMD", "python", "/app/health.py"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 15s
```

This ensures:
- Reliable operation in production environments
- Automatic recovery from temporary failures
- Proper validation of both service components

For detailed information on the health check system, see [HEALTH_CHECKS.md](docs/deployment/HEALTH_CHECKS.md).

### Using with AI Assistants

The Excel MCP Server provides powerful capabilities when used with AI assistants like GitHub Copilot, Claude, or GPT models.

For detailed examples and integration guides, see [USING_WITH_AI.md](docs/USING_WITH_AI.md).

## Remote Hosting & Transport Protocols

This server uses Server-Sent Events (SSE) transport protocol. For different use cases:

1. **Using with Claude Desktop (requires stdio):**
   - Use [Supergateway](https://github.com/supercorp-ai/supergateway) to convert SSE to stdio

2. **Hosting Your MCP Server:**
   - [Remote MCP Server Guide](https://developers.cloudflare.com/agents/guides/remote-mcp-server/)

## Environment Variables

- `FASTMCP_PORT`: Server port for MCP server (default: 8000)
- `FASTAPI_PORT`: Port for web interface and REST API (default: 8080)
- `EXCEL_FILES_PATH`: Directory for Excel files (default: `./excel_files`)

## Web Interface and API

The Excel MCP Server includes a web interface and REST API for managing Excel files.

### Web Interface

Access the web interface at http://localhost:8080 when running the server.

Features:
- View all available Excel files
- Upload new Excel files
- Download existing Excel files
- Delete Excel files
- View Excel file information

### REST API Endpoints

- `GET /api/files` - List all Excel files
- `GET /api/info/{filename}` - Get information about an Excel workbook
- `GET /download/{filename}` - Download an Excel file
- `POST /upload` - Upload an Excel file
- `GET /delete/{filename}` - Delete an Excel file
- `GET /health` - Health check endpoint

### Accessing Local Files with Docker

To access files from your local filesystem when using Docker:

1. Create a directory on your local machine to store Excel files (e.g., `C:/Users/YourName/Documents/excel-files`)

2. When running with docker-compose:
   ```bash
   # Create a .env file in the mcp-servers root directory
   echo "EXCEL_FILES_PATH=C:/Users/YourName/Documents/excel-files" >> .env
   
   # Then start the container
   npm run docker:excel-mcp-server:up
   ```

3. When running with direct Docker commands:
   ```bash
   docker run -p 8000:8000 -v C:/Users/YourName/Documents/excel-files:/app/excel_files excel-mcp-server
   ```

4. Update your MCP client configuration to point to the same directory:
   ```json
   {
     "mcpServers": {
       "excel": {
         "url": "http://localhost:8000/sse",
         "env": {
           "EXCEL_FILES_PATH": "C:/Users/YourName/Documents/excel-files"
         }
       }
     }
   }
   ```

**Note**: On Windows, use forward slashes (`/`) or escaped backslashes (`\\`) in paths for Docker volume mappings.

## Documentation

The Excel MCP Server comes with comprehensive documentation:

- [TOOLS.md](docs/api/TOOLS.md) - Complete list of available Excel manipulation tools
- [API.md](docs/api/API.md) - Complete API reference for both interfaces
- [HEALTH_CHECKS.md](docs/deployment/HEALTH_CHECKS.md) - Details about the container health check system
- [BUILD_AND_DEPLOYMENT.md](docs/deployment/BUILD_AND_DEPLOYMENT.md) - Guide for building and deploying
- [CSV_OPERATIONS.md](docs/features/CSV_OPERATIONS.md) - Guide to CSV import/export functionality
- [FILE_OPERATIONS.md](docs/features/FILE_OPERATIONS.md) - Details on file management capabilities
- [USING_WITH_AI.md](docs/USING_WITH_AI.md) - Examples of using the server with AI assistants

See the [documentation index](docs/README.md) for more information.

## Testing

The repository includes tools to help test your server installation:

### Creating Test Data

Create a sample Excel file for testing:

```bash
python create_test_file.py
```

### Testing Server Functionality

Verify that both the MCP server and web interface are working correctly:

```bash
python test_server.py --test-file test_data.xlsx
```

This script validates:
- Port availability for both services
- API endpoints functionality
- File upload capabilities
- SSE endpoint connectivity

## File Operations for LLMs

The Excel MCP Server provides direct file operations through its SSE interface, allowing LLMs to manage files without requiring the web interface. For detailed documentation on these capabilities, see [FILE_OPERATIONS.md](docs/features/FILE_OPERATIONS.md).

## License

MIT License - see [LICENSE](LICENSE) for details.
