"""Main MCP server application that coordinates all components."""

import asyncio
import json
import signal
import sys
import time
from typing import Dict, Any, Optional, Callable
from concurrent.futures import ThreadPoolExecutor
import threading

from .adapters.outlook_adapter import OutlookAdapter
from .services.folder_service import FolderService
from .services.email_service import EmailService
from .protocol.mcp_protocol_handler import MCPProtocolHandler
from .routing.request_router import RequestRouter
from .error_handler import ErrorHandler, ErrorContext
from .logging.logger import get_logger, configure_logging
from .models.mcp_models import MCPRequest, MCPResponse
from .models.exceptions import (
    OutlookMCPError,
    OutlookConnectionError,
    ValidationError
)


class OutlookMCPServer:
    """Main MCP server class that coordinates all components."""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize the Outlook MCP server.
        
        Args:
            config: Optional configuration dictionary
        """
        self.config = config or self._get_default_config()
        self.logger = get_logger(__name__)
        
        # Initialize components
        self.outlook_adapter: Optional[OutlookAdapter] = None
        self.email_service: Optional[EmailService] = None
        self.folder_service: Optional[FolderService] = None
        self.protocol_handler: Optional[MCPProtocolHandler] = None
        self.request_router: Optional[RequestRouter] = None
        self.error_handler: Optional[ErrorHandler] = None
        
        # Server state
        self._running = False
        self._shutdown_event = threading.Event()
        self._executor: Optional[ThreadPoolExecutor] = None
        
        # Statistics
        self._stats = {
            "requests_processed": 0,
            "requests_successful": 0,
            "requests_failed": 0,
            "start_time": None,
            "connections_established": 0
        }
        
        self.logger.info("Outlook MCP Server initialized", config=self.config)
    
    def _get_default_config(self) -> Dict[str, Any]:
        """Get default server configuration."""
        return {
            "log_level": "INFO",
            "log_dir": "logs",
            "max_concurrent_requests": 10,
            "request_timeout": 30,
            "outlook_connection_timeout": 10,
            "enable_performance_logging": True,
            "enable_console_output": True
        }
    
    async def start(self) -> None:
        """
        Start the MCP server and initialize all components.
        
        Raises:
            OutlookConnectionError: If Outlook connection cannot be established
            RuntimeError: If server is already running
        """
        if self._running:
            raise RuntimeError("Server is already running")
        
        try:
            self.logger.info("Starting Outlook MCP Server")
            
            # Configure logging
            self._configure_logging()
            
            # Initialize components
            await self._initialize_components()
            
            # Setup signal handlers for graceful shutdown
            self._setup_signal_handlers()
            
            # Start the server
            self._running = True
            self._stats["start_time"] = time.time()
            
            # Create thread pool for concurrent request processing
            self._executor = ThreadPoolExecutor(
                max_workers=self.config["max_concurrent_requests"],
                thread_name_prefix="mcp-request"
            )
            
            self.logger.info("Outlook MCP Server started successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to start server: {str(e)}", exc_info=True)
            await self._cleanup()
            raise
    
    async def stop(self) -> None:
        """Stop the MCP server and cleanup resources."""
        if not self._running:
            return
        
        self.logger.info("Stopping Outlook MCP Server")
        
        try:
            # Signal shutdown
            self._running = False
            self._shutdown_event.set()
            
            # Shutdown thread pool
            if self._executor:
                self.logger.debug("Shutting down thread pool")
                try:
                    # Try with timeout parameter (Python 3.9+)
                    self._executor.shutdown(wait=True, timeout=5.0)
                except TypeError:
                    # Fallback for older Python versions
                    self._executor.shutdown(wait=True)
                self._executor = None
            
            # Cleanup components
            await self._cleanup()
            
            # Log final statistics
            self._log_final_stats()
            
            self.logger.info("Outlook MCP Server stopped successfully")
            
        except Exception as e:
            self.logger.error(f"Error during server shutdown: {str(e)}", exc_info=True)
    
    def _configure_logging(self) -> None:
        """Configure the logging system based on server configuration."""
        configure_logging(
            log_level=self.config["log_level"],
            log_dir=self.config["log_dir"],
            console_output=self.config["enable_console_output"]
        )
        
        self.logger.info("Logging configured", 
                        log_level=self.config["log_level"],
                        log_dir=self.config["log_dir"])
    
    async def _initialize_components(self) -> None:
        """Initialize all server components."""
        self.logger.debug("Initializing server components")
        
        # Initialize error handler
        self.error_handler = ErrorHandler(self.logger._logger)
        
        # Initialize Outlook adapter
        self.outlook_adapter = OutlookAdapter()
        
        # Connect to Outlook
        await self._connect_to_outlook()
        
        # Initialize services
        self.email_service = EmailService(self.outlook_adapter)
        self.folder_service = FolderService(self.outlook_adapter)
        
        # Initialize protocol handler
        self.protocol_handler = MCPProtocolHandler()
        
        # Initialize request router and register handlers
        self.request_router = RequestRouter()
        self._register_request_handlers()
        
        self.logger.info("All components initialized successfully")
    
    async def _connect_to_outlook(self) -> None:
        """Connect to Outlook with timeout and retry logic."""
        self.logger.debug("Connecting to Outlook")
        
        timeout = self.config["outlook_connection_timeout"]
        
        try:
            # Use asyncio to run the blocking connect operation with timeout
            loop = asyncio.get_event_loop()
            await asyncio.wait_for(
                loop.run_in_executor(None, self.outlook_adapter.connect),
                timeout=timeout
            )
            
            self._stats["connections_established"] += 1
            self.logger.log_connection_status(True, "Successfully connected to Outlook")
            
        except asyncio.TimeoutError:
            error_msg = f"Outlook connection timed out after {timeout} seconds"
            self.logger.log_connection_status(False, error_msg)
            raise OutlookConnectionError(error_msg)
        except Exception as e:
            self.logger.log_connection_status(False, str(e))
            raise
    
    def _register_request_handlers(self) -> None:
        """Register MCP method handlers with the request router."""
        self.logger.debug("Registering request handlers")
        
        # Register email operations
        self.request_router.register_handler("list_emails", self._handle_list_emails)
        self.request_router.register_handler("get_email", self._handle_get_email)
        self.request_router.register_handler("search_emails", self._handle_search_emails)
        
        # Register folder operations
        self.request_router.register_handler("get_folders", self._handle_get_folders)
        
        self.logger.debug("Request handlers registered successfully")
    
    def _setup_signal_handlers(self) -> None:
        """Setup signal handlers for graceful shutdown."""
        def signal_handler(signum, frame):
            self.logger.info(f"Received signal {signum}, initiating graceful shutdown")
            asyncio.create_task(self.stop())
        
        # Register signal handlers for graceful shutdown
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
        
        if hasattr(signal, 'SIGBREAK'):  # Windows
            signal.signal(signal.SIGBREAK, signal_handler)
    
    async def handle_request(self, request_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Handle an incoming MCP request.
        
        Args:
            request_data: Raw request data dictionary
            
        Returns:
            Dict containing the MCP response
        """
        request_id = request_data.get("id", "unknown")
        start_time = time.time()
        
        try:
            # Parse and validate MCP request
            request = MCPRequest.from_dict(request_data)
            
            self.logger.log_mcp_request(request.id, request.method, request.params)
            
            # Create error context for potential error handling
            error_context = self.error_handler.create_context(
                request_id=request.id,
                method=request.method,
                parameters=request.params
            )
            
            # Process request with protocol handler
            protocol_response = self.protocol_handler.process_request(request)
            if protocol_response and protocol_response.error:
                # Protocol validation failed
                self._update_stats(False)
                return protocol_response.to_dict()
            
            # Route request to appropriate handler
            result = await self._route_request_async(request)
            
            # Format successful response
            response_dict = self.protocol_handler.format_response(result, request.id)
            
            # Update statistics and log success
            duration = time.time() - start_time
            self._update_stats(True)
            self.logger.log_mcp_response(request.id, request.method, True, duration)
            
            return response_dict
            
        except Exception as e:
            # Handle error with comprehensive error handling
            duration = time.time() - start_time
            self._update_stats(False)
            
            # Create error context if not already created
            if 'error_context' not in locals():
                error_context = self.error_handler.create_context(
                    request_id=request_id,
                    method=request_data.get("method", "unknown"),
                    parameters=request_data.get("params", {})
                )
            
            # Use error handler to process the error
            error_response = self.error_handler.handle_error(e, error_context)
            
            self.logger.log_mcp_response(request_id, request_data.get("method", "unknown"), False, duration)
            
            return error_response
    
    async def _route_request_async(self, request: MCPRequest) -> Any:
        """Route request asynchronously using thread pool."""
        if not self._executor:
            raise RuntimeError("Server not properly initialized")
        
        # Run the synchronous request routing in thread pool
        loop = asyncio.get_event_loop()
        return await loop.run_in_executor(
            self._executor,
            self.request_router.route_request,
            request
        )
    
    def _handle_list_emails(self, folder: str = None, unread_only: bool = False, limit: int = 50) -> Dict[str, Any]:
        """Handle list_emails MCP method."""
        with self.logger.time_operation("list_emails"):
            return self.email_service.list_emails(folder, unread_only, limit)
    
    def _handle_get_email(self, email_id: str) -> Dict[str, Any]:
        """Handle get_email MCP method."""
        with self.logger.time_operation("get_email"):
            return self.email_service.get_email(email_id)
    
    def _handle_search_emails(self, query: str, folder: str = None, limit: int = 50) -> Dict[str, Any]:
        """Handle search_emails MCP method."""
        with self.logger.time_operation("search_emails"):
            return self.email_service.search_emails(query, folder, limit)
    
    def _handle_get_folders(self) -> Dict[str, Any]:
        """Handle get_folders MCP method."""
        with self.logger.time_operation("get_folders"):
            return {"folders": self.folder_service.get_folders()}
    
    def _update_stats(self, success: bool) -> None:
        """Update server statistics."""
        self._stats["requests_processed"] += 1
        if success:
            self._stats["requests_successful"] += 1
        else:
            self._stats["requests_failed"] += 1
    
    def get_server_info(self) -> Dict[str, Any]:
        """Get server information and capabilities."""
        if self.protocol_handler:
            return self.protocol_handler.get_server_info()
        else:
            return {
                "name": "outlook-mcp-server",
                "version": "1.0.0",
                "status": "initializing"
            }
    
    def get_server_stats(self) -> Dict[str, Any]:
        """Get server statistics."""
        stats = self._stats.copy()
        
        if stats["start_time"]:
            stats["uptime_seconds"] = time.time() - stats["start_time"]
        
        stats["is_running"] = self._running
        stats["outlook_connected"] = self.outlook_adapter.is_connected() if self.outlook_adapter else False
        
        return stats
    
    def is_running(self) -> bool:
        """Check if server is running."""
        return self._running
    
    def is_healthy(self) -> bool:
        """Check if server is healthy (running and connected to Outlook)."""
        return (
            self._running and 
            self.outlook_adapter is not None and 
            self.outlook_adapter.is_connected()
        )
    
    def get_health_status(self) -> Dict[str, Any]:
        """Get detailed health status information."""
        from datetime import datetime
        
        return {
            "status": "healthy" if self.is_healthy() else "unhealthy",
            "healthy": self.is_healthy(),
            "running": self._running,
            "outlook_connected": self.outlook_adapter.is_connected() if self.outlook_adapter else False,
            "timestamp": datetime.utcnow().isoformat(),
            "server_info": self.get_server_info()
        }
    
    async def _cleanup(self) -> None:
        """Cleanup server resources."""
        self.logger.debug("Cleaning up server resources")
        
        try:
            # Disconnect from Outlook
            if self.outlook_adapter:
                loop = asyncio.get_event_loop()
                await loop.run_in_executor(None, self.outlook_adapter.disconnect)
                self.outlook_adapter = None
            
            # Clear other components
            self.email_service = None
            self.folder_service = None
            self.protocol_handler = None
            self.request_router = None
            self.error_handler = None
            
            self.logger.debug("Server resources cleaned up")
            
        except Exception as e:
            self.logger.error(f"Error during cleanup: {str(e)}", exc_info=True)
    
    def _log_final_stats(self) -> None:
        """Log final server statistics."""
        stats = self.get_server_stats()
        self.logger.info("Final server statistics", stats=stats)


# Server configuration and initialization utilities
def create_server_config(
    log_level: str = "INFO",
    log_dir: str = "logs",
    max_concurrent_requests: int = 10,
    request_timeout: int = 30,
    outlook_connection_timeout: int = 10,
    enable_performance_logging: bool = True,
    enable_console_output: bool = True
) -> Dict[str, Any]:
    """
    Create server configuration dictionary.
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_dir: Directory for log files
        max_concurrent_requests: Maximum number of concurrent requests
        request_timeout: Request timeout in seconds
        outlook_connection_timeout: Outlook connection timeout in seconds
        enable_performance_logging: Whether to enable performance logging
        enable_console_output: Whether to output logs to console
        
    Returns:
        Configuration dictionary
    """
    return {
        "log_level": log_level,
        "log_dir": log_dir,
        "max_concurrent_requests": max_concurrent_requests,
        "request_timeout": request_timeout,
        "outlook_connection_timeout": outlook_connection_timeout,
        "enable_performance_logging": enable_performance_logging,
        "enable_console_output": enable_console_output
    }


async def create_and_start_server(config: Optional[Dict[str, Any]] = None) -> OutlookMCPServer:
    """
    Create and start an Outlook MCP server.
    
    Args:
        config: Optional server configuration
        
    Returns:
        Started server instance
        
    Raises:
        OutlookConnectionError: If Outlook connection fails
        RuntimeError: If server startup fails
    """
    server = OutlookMCPServer(config)
    await server.start()
    return server