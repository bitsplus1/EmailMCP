"""
EmailMCP - A Model Context Protocol server for reading Outlook emails.

This package provides an MCP server that can connect to and read emails from 
Outlook/Exchange accounts, making email data accessible to MCP clients.
"""

__version__ = "0.1.0"
__author__ = "EmailMCP Team"
__email__ = "support@email-mcp.com"

from .server import EmailMCPServer

__all__ = ["EmailMCPServer"]