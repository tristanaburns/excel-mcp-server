#!/usr/bin/env python3

import http.client
import sys
import time
import json
import socket

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
            
            # Read the response body
            body = response.read()
            
            # Check if the server is responding
            if 200 <= response.status < 300:
                return True, f"{host}:{port}{path} is healthy ({response.status})"
            else:
                if attempt == max_retries - 1:
                    return False, f"{host}:{port}{path} returned status {response.status}, body: {body}"
        except Exception as e:
            if attempt == max_retries - 1:
                return False, f"Failed to connect to {host}:{port}{path}: {str(e)}"
        finally:
            conn.close() if 'conn' in locals() else None
            
        # Wait before retrying
        time.sleep(1)

def check_all_endpoints():
    results = []
    
    # Check if MCP server port is open and listening
    mcp_port_healthy, mcp_port_message = check_port("localhost", 8000)
    results.append(("MCP Server Port", mcp_port_healthy, mcp_port_message))
    
    # Check if Web server port is open and listening
    web_port_healthy, web_port_message = check_port("localhost", 8080)
    results.append(("Web Server Port", web_port_healthy, web_port_message))
    
    # Check Web server health endpoint
    web_healthy, web_message = check_endpoint("localhost", 8080, "/health")
    results.append(("Web Server /health", web_healthy, web_message))
    
    # Check if /api/files endpoint works on the web server
    api_healthy, api_message = check_endpoint("localhost", 8080, "/api/files")
    results.append(("Web Server /api/files", api_healthy, api_message))
    
    return results

try:
    # Check all endpoints
    results = check_all_endpoints()
    
    # Print results
    print("Health check results:")
    all_healthy = True
    
    for name, is_healthy, message in results:
        status = "✅ PASS" if is_healthy else "❌ FAIL"
        print(f"{status} - {name}: {message}")
        if not is_healthy:
            all_healthy = False
    
    if all_healthy:
        print("\nAll health checks passed: Both MCP and Web servers are running correctly")
        sys.exit(0)
    else:
        print("\nAt least one health check failed")
        sys.exit(1)
except Exception as e:
    print(f"Health check script failed: {str(e)}")
    sys.exit(1)