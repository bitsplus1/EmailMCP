"""HTTP server implementation for MCP server to enable remote access."""

import asyncio
import json
import logging
from typing import Dict, Any, Optional
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
from urllib.parse import urlparse, parse_qs

from .server import OutlookMCPServer
from .logging.logger import get_logger


class MCPHTTPRequestHandler(BaseHTTPRequestHandler):
    """HTTP request handler for MCP server."""
    
    def __init__(self, mcp_server: OutlookMCPServer, *args, **kwargs):
        self.mcp_server = mcp_server
        self.logger = get_logger(__name__)
        super().__init__(*args, **kwargs)
    
    def do_POST(self):
        """Handle POST requests."""
        try:
            # Check if this is an MCP endpoint
            if self.path != '/mcp':
                self.send_error(404, "Not Found")
                return
            
            # Get content length
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length == 0:
                self.send_error(400, "Empty request body")
                return
            
            # Read request body
            post_data = self.rfile.read(content_length)
            
            try:
                # Parse JSON request
                request_data = json.loads(post_data.decode('utf-8'))
                self.logger.debug(f"Received HTTP MCP request: {request_data}")
                
                # Process the request asynchronously
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                
                try:
                    response = loop.run_until_complete(
                        self.mcp_server.handle_request(request_data)
                    )
                finally:
                    loop.close()
                
                # Send response
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
                self.send_header('Access-Control-Allow-Headers', 'Content-Type')
                self.end_headers()
                
                response_json = json.dumps(response, ensure_ascii=False)
                self.wfile.write(response_json.encode('utf-8'))
                
                self.logger.debug(f"Sent HTTP MCP response: {response_json}")
                
            except json.JSONDecodeError as e:
                self.logger.error(f"Invalid JSON in request: {e}")
                self.send_error(400, f"Invalid JSON: {str(e)}")
            except Exception as e:
                self.logger.error(f"Error processing request: {e}", exc_info=True)
                self.send_error(500, f"Internal server error: {str(e)}")
                
        except Exception as e:
            self.logger.error(f"Error handling POST request: {e}", exc_info=True)
            self.send_error(500, "Internal server error")
    
    def do_OPTIONS(self):
        """Handle OPTIONS requests for CORS."""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
    
    def do_GET(self):
        """Handle GET requests."""
        if self.path == '/health':
            # Health check endpoint
            try:
                health_status = self.mcp_server.get_health_status()
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                
                response = {
                    "status": "healthy" if health_status.get("healthy", False) else "unhealthy",
                    "timestamp": health_status.get("timestamp"),
                    "server_info": self.mcp_server.get_server_info()
                }
                
                self.wfile.write(json.dumps(response).encode('utf-8'))
                
            except Exception as e:
                self.logger.error(f"Health check failed: {e}")
                self.send_error(500, "Health check failed")
        else:
            self.send_error(404, "Not Found")
    
    def log_message(self, format, *args):
        """Override to use our logger instead of stderr."""
        self.logger.info(f"{self.address_string()} - {format % args}")


class MCPHTTPServer:
    """HTTP server wrapper for MCP server."""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """Initialize the HTTP MCP server."""
        self.config = config or {}
        self.logger = get_logger(__name__)
        self.mcp_server: Optional[OutlookMCPServer] = None
        self.http_server: Optional[HTTPServer] = None
        self.server_thread: Optional[threading.Thread] = None
        self._running = False
        
        # Get server configuration
        self.host = self.config.get("server_host", "127.0.0.1")
        self.port = self.config.get("server_port", 8080)
    
    async def start(self) -> None:
        """Start the HTTP MCP server."""
        try:
            self.logger.info(f"Starting MCP HTTP server on {self.host}:{self.port}")
            
            # Create and start the core MCP server
            self.mcp_server = OutlookMCPServer(self.config)
            await self.mcp_server.start()
            
            # Create HTTP server
            def handler_factory(*args, **kwargs):
                return MCPHTTPRequestHandler(self.mcp_server, *args, **kwargs)
            
            self.http_server = HTTPServer((self.host, self.port), handler_factory)
            
            # Start HTTP server in a separate thread
            self.server_thread = threading.Thread(
                target=self._run_http_server,
                daemon=True
            )
            self.server_thread.start()
            
            self._running = True
            self.logger.info(f"MCP HTTP server started successfully on http://{self.host}:{self.port}")
            self.logger.info("Available endpoints:")
            self.logger.info(f"  - POST http://{self.host}:{self.port}/mcp (MCP requests)")
            self.logger.info(f"  - GET  http://{self.host}:{self.port}/health (Health check)")
            
        except Exception as e:
            self.logger.error(f"Failed to start MCP HTTP server: {str(e)}", exc_info=True)
            raise
    
    async def stop(self) -> None:
        """Stop the HTTP MCP server."""
        if not self._running:
            return
        
        self.logger.info("Stopping MCP HTTP server")
        self._running = False
        
        # Stop HTTP server
        if self.http_server:
            self.http_server.shutdown()
            self.http_server.server_close()
            self.http_server = None
        
        # Wait for server thread to finish
        if self.server_thread and self.server_thread.is_alive():
            self.server_thread.join(timeout=5.0)
            self.server_thread = None
        
        # Stop MCP server
        if self.mcp_server:
            await self.mcp_server.stop()
            self.mcp_server = None
        
        self.logger.info("MCP HTTP server stopped")
    
    def _run_http_server(self) -> None:
        """Run the HTTP server in a separate thread."""
        try:
            self.logger.debug("HTTP server thread started")
            self.http_server.serve_forever()
        except Exception as e:
            if self._running:  # Only log if we're supposed to be running
                self.logger.error(f"HTTP server thread error: {e}", exc_info=True)
        finally:
            self.logger.debug("HTTP server thread finished")
    
    def is_running(self) -> bool:
        """Check if the server is running."""
        return self._running and self.mcp_server is not None
    
    def get_server_info(self) -> Dict[str, Any]:
        """Get server information."""
        info = {
            "mode": "http",
            "host": self.host,
            "port": self.port,
            "running": self.is_running(),
            "endpoints": {
                "mcp": f"http://{self.host}:{self.port}/mcp",
                "health": f"http://{self.host}:{self.port}/health"
            }
        }
        
        if self.mcp_server:
            info.update(self.mcp_server.get_server_info())
        
        return info


async def run_http_server(config: Optional[Dict[str, Any]] = None) -> None:
    """Run the MCP HTTP server."""
    server = MCPHTTPServer(config)
    
    try:
        await server.start()
        
        # Keep the server running
        while server.is_running():
            await asyncio.sleep(1)
            
    except KeyboardInterrupt:
        pass
    finally:
        await server.stop()


if __name__ == "__main__":
    # Run as HTTP server
    asyncio.run(run_http_server())