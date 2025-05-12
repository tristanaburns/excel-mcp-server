# Excel MCP Server Health Checks

This document provides details about the health check system implemented for the Excel MCP Server Docker container.

## Overview

The Excel MCP Server runs two distinct services:
1. **MCP Server** on port 8000 - Server-sent events (SSE) endpoint for AI assistants
2. **Web Interface** on port 8080 - FastAPI application for file management

The health check system ensures both services are running properly and can recover automatically from failures.

## Health Check Implementation

The health check script (`health.py`) in the container performs several validations:

### 1. Port Availability Checks

Verifies that both services are listening on their respective ports:
- MCP Server on port 8000
- Web Interface on port 8080

```python
def check_port(host, port, max_retries=3):
    """Check if a port is open and listening."""
    for attempt in range(max_retries):
        try:
            # Try to create a socket connection
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            result = sock.connect_ex((host, port))
            sock.close()
            
            # If result is 0, the port is open
            if result == 0:
                return True, f"{host}:{port} is open and listening"
            else:
                if attempt == max_retries - 1:
                    return False, f"{host}:{port} is not open or not responding"
        except Exception as e:
            if attempt == max_retries - 1:
                return False, f"Failed to connect to {host}:{port}: {str(e)}"
        
        # Wait before retrying
        time.sleep(1)
```

### 2. Web Interface Health Check

Verifies the `/health` endpoint is responding correctly:

```python
def check_endpoint(host, port, path, method="GET", max_retries=3):
    """Check if an endpoint is responding correctly."""
    for attempt in range(max_retries):
        try:
            # Connect to the server
            conn = http.client.HTTPConnection(host, port, timeout=5)
            
            # Send a request to the health endpoint
            conn.request(method, path)
            
            # Get the response
            response = conn.getresponse()
            
            # Check if the server is responding
            if 200 <= response.status < 300:
                return True, f"{host}:{port}{path} is healthy ({response.status})"
            else:
                if attempt == max_retries - 1:
                    return False, f"{host}:{port}{path} returned status {response.status}"
        except Exception as e:
            if attempt == max_retries - 1:
                return False, f"Failed to connect to {host}:{port}{path}: {str(e)}"
        finally:
            conn.close() if 'conn' in locals() else None
            
        # Wait before retrying
        time.sleep(1)
```

### 3. API Functionality Verification

Checks if the `/api/files` endpoint is working properly, which validates that the file system is accessible.

```python
# Check all endpoints
results = check_all_endpoints()

# Check if /api/files endpoint works on the web server
api_healthy, api_message = check_endpoint("localhost", 8080, "/api/files")
results.append(("Web Server /api/files", api_healthy, api_message))
```

## Docker Health Check Configuration

The Docker health check is configured in the `docker-compose.yml` file:

```yaml
healthcheck:
  test: ["CMD", "python", "/app/health.py"]
  interval: 30s
  timeout: 10s
  retries: 3
  start_period: 15s
```

Parameters explained:
- **test**: Command to run the health check script
- **interval**: How often to check (30 seconds)
- **timeout**: Maximum time a health check can take before considered failed (10 seconds)
- **retries**: Number of consecutive failures required to mark as unhealthy (3)
- **start_period**: Initial grace period before starting health checks (15 seconds)

## Benefits

1. **Reliability**: Automatic recovery from failures through container restart
2. **Transparency**: Clear health status visible through Docker commands
3. **Comprehensive**: Checks both service components
4. **Resilience**: Graceful handling of temporary network issues with retry mechanism

## Troubleshooting

If the container is marked as unhealthy:

1. Check container logs:
   ```bash
   docker logs excel-mcp-server
   ```

2. Run the health check manually:
   ```bash
   docker exec excel-mcp-server python /app/health.py
   ```

3. Common issues:
   - File permission problems in the Excel files directory
   - Network port conflicts
   - Resource constraints (CPU/memory)

## Example Health Check Output

Successful check:
```
Health check results:
✅ PASS - MCP Server Port: localhost:8000 is open and listening
✅ PASS - Web Server Port: localhost:8080 is open and listening
✅ PASS - Web Server /health: localhost:8080/health is healthy (200)
✅ PASS - Web Server /api/files: localhost:8080/api/files is healthy (200)

All health checks passed: Both MCP and Web servers are running correctly
```

Failed check:
```
Health check results:
✅ PASS - MCP Server Port: localhost:8000 is open and listening
❌ FAIL - Web Server Port: localhost:8080 is not open or not responding
❌ FAIL - Web Server /health: Failed to connect to localhost:8080/health: Connection refused
❌ FAIL - Web Server /api/files: Failed to connect to localhost:8080/api/files: Connection refused

At least one health check failed
```
