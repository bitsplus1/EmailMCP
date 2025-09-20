"""MCP Server with stdio transport for standard MCP client communication."""

import asyncio
import json
import sys
from typing import Dict, Any, Optional
import logging

from .server import OutlookMCPServer
from .logging.logger import get_logger


class MCPStdioServer:
    """MCP Server that communicates via stdin/stdout for standard MCP protocol."""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """Initialize the stdio MCP server."""
        self.config = config or {}
        self.logger = get_logger(__name__)
        self.server: Optional[OutlookMCPServer] = None
        self._running = False
        
        # Disable console output for stdio mode to avoid interfering with MCP communication
        if "enable_console_output" not in self.config:
            self.config["enable_console_output"] = False
    
    async def start(self) -> None:
        """Start the MCP stdio server."""
        try:
            self.logger.info("Starting MCP stdio server")
            
            # Create and start the core server
            self.server = OutlookMCPServer(self.config)
            await self.server.start()
            
            self._running = True
            self.logger.info("MCP stdio server started successfully")
            
            # Start the stdio communication loop
            await self._run_stdio_loop()
            
        except Exception as e:
            self.logger.error(f"Failed to start MCP stdio server: {str(e)}", exc_info=True)
            raise
    
    async def stop(self) -> None:
        """Stop the MCP stdio server."""
        if not self._running:
            return
        
        self.logger.info("Stopping MCP stdio server")
        self._running = False
        
        if self.server:
            await self.server.stop()
            self.server = None
        
        self.logger.info("MCP stdio server stopped")
    
    async def _run_stdio_loop(self) -> None:
        """Run the main stdio communication loop."""
        self.logger.debug("Starting stdio communication loop")
        
        try:
            # Read from stdin and write to stdout
            reader = asyncio.StreamReader()
            protocol = asyncio.StreamReaderProtocol(reader)
            
            # Connect to stdin
            transport, _ = await asyncio.get_event_loop().connect_read_pipe(
                lambda: protocol, sys.stdin
            )
            
            while self._running:
                try:
                    # Read a line from stdin
                    line = await reader.readline()
                    
                    if not line:
                        # EOF reached
                        self.logger.debug("EOF reached on stdin")
                        break
                    
                    # Decode and parse JSON
                    line_str = line.decode('utf-8').strip()
                    if not line_str:
                        continue
                    
                    self.logger.debug(f"Received MCP request: {line_str}")
                    
                    try:
                        request_data = json.loads(line_str)
                    except json.JSONDecodeError as e:
                        self.logger.error(f"Invalid JSON received: {e}")
                        # Send error response
                        error_response = {
                            "jsonrpc": "2.0",
                            "id": None,
                            "error": {
                                "code": -32700,
                                "message": "Parse error",
                                "data": {"details": str(e)}
                            }
                        }
                        await self._send_response(error_response)
                        continue
                    
                    # Handle the request
                    response = await self._handle_request(request_data)
                    
                    # Send response
                    await self._send_response(response)
                    
                except Exception as e:
                    self.logger.error(f"Error in stdio loop: {str(e)}", exc_info=True)
                    # Try to send error response
                    try:
                        error_response = {
                            "jsonrpc": "2.0",
                            "id": None,
                            "error": {
                                "code": -32603,
                                "message": "Internal error",
                                "data": {"details": str(e)}
                            }
                        }
                        await self._send_response(error_response)
                    except:
                        pass  # If we can't send error response, just continue
            
            transport.close()
            
        except Exception as e:
            self.logger.error(f"Fatal error in stdio loop: {str(e)}", exc_info=True)
            raise
    
    async def _handle_request(self, request_data: Dict[str, Any]) -> Dict[str, Any]:
        """Handle an MCP request and return response."""
        try:
            # Check for special MCP protocol methods
            method = request_data.get("method")
            request_id = request_data.get("id")
            
            if method == "initialize":
                # Handle MCP initialization
                return await self._handle_initialize(request_data)
            elif method == "notifications/initialized":
                # Handle initialization notification
                return await self._handle_initialized(request_data)
            elif method == "ping":
                # Handle ping
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": {"status": "pong"}
                }
            else:
                # Handle regular MCP tool calls
                if not self.server:
                    return {
                        "jsonrpc": "2.0",
                        "id": request_id,
                        "error": {
                            "code": -32002,
                            "message": "Server not initialized"
                        }
                    }
                
                return await self.server.handle_request(request_data)
                
        except Exception as e:
            self.logger.error(f"Error handling request: {str(e)}", exc_info=True)
            return {
                "jsonrpc": "2.0",
                "id": request_data.get("id"),
                "error": {
                    "code": -32603,
                    "message": "Internal error",
                    "data": {"details": str(e)}
                }
            }
    
    async def _handle_initialize(self, request_data: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP initialize request."""
        request_id = request_data.get("id")
        
        try:
            # Get client info from params
            params = request_data.get("params", {})
            client_info = params.get("clientInfo", {})
            protocol_version = params.get("protocolVersion", "2024-11-05")
            
            self.logger.info(f"MCP initialize request from client: {client_info}")
            
            # Perform handshake with protocol handler
            if self.server and self.server.protocol_handler:
                handshake_info = {
                    "protocolVersion": protocol_version,
                    "clientInfo": client_info
                }
                
                handshake_response = self.server.protocol_handler.handle_handshake(handshake_info)
                
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": handshake_response
                }
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32002,
                        "message": "Server not ready"
                    }
                }
                
        except Exception as e:
            self.logger.error(f"Error in initialize: {str(e)}", exc_info=True)
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "error": {
                    "code": -32603,
                    "message": "Initialize failed",
                    "data": {"details": str(e)}
                }
            }
    
    async def _handle_initialized(self, request_data: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP initialized notification."""
        self.logger.info("MCP client initialization completed")
        
        # This is a notification, so no response is expected
        # But we'll return None to indicate no response should be sent
        return None
    
    async def _send_response(self, response: Optional[Dict[str, Any]]) -> None:
        """Send response to stdout."""
        if response is None:
            return
        
        try:
            response_json = json.dumps(response, ensure_ascii=False)
            self.logger.debug(f"Sending MCP response: {response_json}")
            
            # Write to stdout with newline
            sys.stdout.write(response_json + "\n")
            sys.stdout.flush()
            
        except Exception as e:
            self.logger.error(f"Error sending response: {str(e)}", exc_info=True)


async def run_stdio_server(config: Optional[Dict[str, Any]] = None) -> None:
    """Run the MCP stdio server."""
    server = MCPStdioServer(config)
    
    try:
        await server.start()
    except KeyboardInterrupt:
        pass
    finally:
        await server.stop()


if __name__ == "__main__":
    # Run as stdio server
    asyncio.run(run_stdio_server())