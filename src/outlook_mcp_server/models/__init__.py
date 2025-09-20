"""Data models for Outlook MCP Server."""

from .email_data import EmailData
from .folder_data import FolderData
from .mcp_models import MCPRequest, MCPResponse
from .exceptions import (
    OutlookMCPError,
    ValidationError,
    OutlookConnectionError,
    EmailNotFoundError,
    FolderNotFoundError,
    InvalidParameterError
)

__all__ = [
    "EmailData",
    "FolderData", 
    "MCPRequest",
    "MCPResponse",
    "OutlookMCPError",
    "ValidationError",
    "OutlookConnectionError",
    "EmailNotFoundError",
    "FolderNotFoundError",
    "InvalidParameterError"
]