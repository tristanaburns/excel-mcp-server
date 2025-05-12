import asyncio
import os
import threading
import signal
import sys
import time
import uvicorn
from .server import run_server
from .web import app

def run_fastapi():
    """Run the FastAPI web server."""
    port = int(os.environ.get("FASTAPI_PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")

def signal_handler(sig, frame):
    """Handle interrupt signals properly."""
    print("\nShutting down servers...")
    sys.exit(0)

def main():
    """Start the Excel MCP server and FastAPI web server."""
    # Register signal handler
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    try:
        print("Excel MCP Server")
        print("---------------")
        print("Starting servers...")
        
        # Ensure Excel files directory exists
        excel_files_path = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
        os.makedirs(excel_files_path, exist_ok=True)
        print(f"Excel files directory: {excel_files_path}")
        
        # Start FastAPI in a separate thread
        fastapi_port = os.environ.get("FASTAPI_PORT", "8080")
        print(f"Starting web interface on port {fastapi_port}...")
        fastapi_thread = threading.Thread(target=run_fastapi, daemon=True)
        fastapi_thread.start()
        
        # Give FastAPI time to start
        time.sleep(1)
        print(f"✓ Web interface available at http://0.0.0.0:{fastapi_port}")
        
        # Run MCP server in the main thread
        mcp_port = os.environ.get('FASTMCP_PORT', '8000')
        print(f"Starting MCP server on port {mcp_port}...")
        print(f"✓ MCP server available at http://0.0.0.0:{mcp_port}/sse for AI tools")
        asyncio.run(run_server())
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 