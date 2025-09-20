"""Outlook MCP Server package."""

from .server import OutlookMCPServer, create_server_config
from .mcp_stdio_server import MCPStdioServer, run_stdio_server
from .error_handler import ErrorHandler, ErrorContext, ErrorSeverity
from .logging import Logger, get_logger, configure_logging

__version__ = "1.0.0"
__all__ = [
    "OutlookMCPServer",
    "MCPStdioServer", 
    "create_server_config",
    "run_stdio_server",
    "ErrorHandler", 
    "ErrorContext", 
    "ErrorSeverity",
    "Logger", 
    "get_logger", 
    "configure_logging"
]