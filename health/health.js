const http = require('http');

/**
 * Check the health of a server endpoint
 * @param {string} hostname - The hostname to check
 * @param {number} port - The port to check
 * @param {string} path - The path to check
 * @returns {Promise<{isHealthy: boolean, message: string}>} - Health status and message
 */
function checkEndpoint(hostname, port, path) {
  return new Promise((resolve) => {
    const req = http.request({
      hostname,
      port,
      path,
      method: 'HEAD',
      timeout: 5000 // 5 seconds timeout
    }, (res) => {
      // Check if the response status code is OK (200-299)
      if (res.statusCode >= 200 && res.statusCode < 300) {
        resolve({ isHealthy: true, message: `${hostname}:${port} is healthy` });
      } else {
        resolve({ 
          isHealthy: false, 
          message: `${hostname}:${port} returned status ${res.statusCode}` 
        });
      }
    });

    // Handle request errors
    req.on('error', (err) => {
      resolve({ 
        isHealthy: false, 
        message: `Failed to connect to ${hostname}:${port}: ${err.message}` 
      });
    });

    // Handle timeout
    req.on('timeout', () => {
      req.abort();
      resolve({ 
        isHealthy: false, 
        message: `Request to ${hostname}:${port} timed out` 
      });
    });

    // End the request
    req.end();
  });
}

// Check both services
async function runHealthCheck() {
  try {
    // Check MCP server health
    const mcpHealth = await checkEndpoint('localhost', 8000, '/health');
    
    // Check Web server health
    const webHealth = await checkEndpoint('localhost', 8080, '/health');
    
    if (mcpHealth.isHealthy && webHealth.isHealthy) {
      console.log('Health check passed: Both MCP and Web servers are running');
      process.exit(0); // Exit with success
    } else {
      console.error('Health check failed:');
      console.error(`MCP Server: ${mcpHealth.message}`);
      console.error(`Web Server: ${webHealth.message}`);
      process.exit(1); // Exit with error
    }
  } catch (err) {
    console.error(`Health check failed: ${err.message}`);
    process.exit(1); // Exit with error
  }
}

// Run the health check
runHealthCheck();