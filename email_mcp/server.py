"""
Main MCP server implementation for EmailMCP.

This module contains the core MCP server that handles email-related operations.
"""

import asyncio
import json
import logging
from typing import Any, Dict, List, Optional, Union

from mcp.server import Server
from mcp.server.models import InitializationOptions
from mcp.server.stdio import stdio_server
from mcp.types import (
    CallToolRequest,
    CallToolResult,
    ListToolsRequest,
    ListToolsResult,
    TextContent,
    Tool,
)
from pydantic import BaseModel

from .config import config
from .clients import EmailClient, MockEmailClient, OutlookClient, ExchangeClient

# Configure logging
logging.basicConfig(level=getattr(logging, config.server.log_level))
logger = logging.getLogger(__name__)


class EmailMCPServer:
    """Main EmailMCP server class that implements MCP protocol for email operations."""

    def __init__(self) -> None:
        """Initialize the EmailMCP server."""
        self.server = Server("email-mcp")
        self.email_client: Optional[EmailClient] = None
        self._setup_handlers()
        self._setup_email_client()
    
    def _setup_email_client(self) -> None:
        """Set up the email client based on configuration."""
        try:
            # Try to set up Outlook client if credentials are provided
            if all([config.outlook.client_id, config.outlook.client_secret, config.outlook.tenant_id]):
                logger.info("Setting up Outlook client")
                self.email_client = OutlookClient(
                    config.outlook.client_id,
                    config.outlook.client_secret,
                    config.outlook.tenant_id
                )
            # Try to set up Exchange client if credentials are provided
            elif all([config.exchange.server, config.exchange.username, config.exchange.password]):
                logger.info("Setting up Exchange client")
                self.email_client = ExchangeClient(
                    config.exchange.server,
                    config.exchange.username,
                    config.exchange.password,
                    config.exchange.domain
                )
            else:
                logger.info("No email credentials configured, using mock client")
                self.email_client = MockEmailClient()
        except Exception as e:
            logger.warning(f"Failed to setup email client: {e}, falling back to mock client")
            self.email_client = MockEmailClient()

    def _setup_handlers(self) -> None:
        """Set up MCP protocol handlers."""
        
        @self.server.list_tools()
        async def list_tools() -> List[Tool]:
            """List available email tools."""
            return [
                Tool(
                    name="list_emails",
                    description="List emails from the inbox with optional filtering",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "folder": {
                                "type": "string",
                                "description": "Email folder to search (default: inbox)",
                                "default": "inbox"
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of emails to return (default: 10)",
                                "default": 10,
                                "minimum": 1,
                                "maximum": 100
                            },
                            "unread_only": {
                                "type": "boolean",
                                "description": "Only return unread emails",
                                "default": False
                            }
                        }
                    }
                ),
                Tool(
                    name="get_email",
                    description="Get detailed information about a specific email",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "email_id": {
                                "type": "string",
                                "description": "Unique identifier of the email"
                            }
                        },
                        "required": ["email_id"]
                    }
                ),
                Tool(
                    name="search_emails",
                    description="Search emails by subject, sender, or content",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "query": {
                                "type": "string",
                                "description": "Search query"
                            },
                            "folder": {
                                "type": "string",
                                "description": "Email folder to search (default: inbox)",
                                "default": "inbox"
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of results (default: 10)",
                                "default": 10,
                                "minimum": 1,
                                "maximum": 50
                            }
                        },
                        "required": ["query"]
                    }
                ),
                Tool(
                    name="get_folders",
                    description="List all available email folders",
                    inputSchema={
                        "type": "object",
                        "properties": {}
                    }
                )
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: Dict[str, Any]) -> CallToolResult:
            """Handle tool calls."""
            try:
                if name == "list_emails":
                    return await self._list_emails(arguments)
                elif name == "get_email":
                    return await self._get_email(arguments)
                elif name == "search_emails":
                    return await self._search_emails(arguments)
                elif name == "get_folders":
                    return await self._get_folders(arguments)
                else:
                    return CallToolResult(
                        content=[
                            TextContent(
                                type="text",
                                text=f"Unknown tool: {name}"
                            )
                        ],
                        isError=True
                    )
            except Exception as e:
                logger.error(f"Error calling tool {name}: {e}")
                return CallToolResult(
                    content=[
                        TextContent(
                            type="text",
                            text=f"Error: {str(e)}"
                        )
                    ],
                    isError=True
                )

    async def _list_emails(self, arguments: Dict[str, Any]) -> CallToolResult:
        """List emails from the specified folder."""
        folder = arguments.get("folder", "inbox")
        limit = min(arguments.get("limit", 10), config.server.max_emails_per_request)
        unread_only = arguments.get("unread_only", False)
        
        try:
            if not self.email_client:
                raise RuntimeError("Email client not configured")
            
            # Connect if not already connected
            if not hasattr(self.email_client, 'connected') or not self.email_client.connected:
                await self.email_client.connect()
            
            emails = await self.email_client.list_emails(folder, limit, unread_only)
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(emails, indent=2)
                    )
                ]
            )
        except Exception as e:
            logger.error(f"Error listing emails: {e}")
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Error listing emails: {str(e)}"
                    )
                ],
                isError=True
            )

    async def _get_email(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Get detailed information about a specific email."""
        email_id = arguments.get("email_id")
        
        if not email_id:
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text="Error: email_id is required"
                    )
                ],
                isError=True
            )
        
        try:
            if not self.email_client:
                raise RuntimeError("Email client not configured")
            
            # Connect if not already connected
            if not hasattr(self.email_client, 'connected') or not self.email_client.connected:
                await self.email_client.connect()
            
            email = await self.email_client.get_email(email_id)
            
            if email is None:
                return CallToolResult(
                    content=[
                        TextContent(
                            type="text",
                            text=f"Email with ID {email_id} not found"
                        )
                    ],
                    isError=True
                )
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(email, indent=2)
                    )
                ]
            )
        except Exception as e:
            logger.error(f"Error getting email {email_id}: {e}")
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Error getting email: {str(e)}"
                    )
                ],
                isError=True
            )

    async def _search_emails(self, arguments: Dict[str, Any]) -> CallToolResult:
        """Search emails by query."""
        query = arguments.get("query")
        folder = arguments.get("folder", "inbox")
        limit = min(arguments.get("limit", 10), config.server.max_emails_per_request)
        
        if not query:
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text="Error: query is required"
                    )
                ],
                isError=True
            )
        
        try:
            if not self.email_client:
                raise RuntimeError("Email client not configured")
            
            # Connect if not already connected
            if not hasattr(self.email_client, 'connected') or not self.email_client.connected:
                await self.email_client.connect()
            
            results = await self.email_client.search_emails(query, folder, limit)
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps({
                            "query": query,
                            "folder": folder,
                            "results": results,
                            "total_found": len(results)
                        }, indent=2)
                    )
                ]
            )
        except Exception as e:
            logger.error(f"Error searching emails: {e}")
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Error searching emails: {str(e)}"
                    )
                ],
                isError=True
            )

    async def _get_folders(self, arguments: Dict[str, Any]) -> CallToolResult:
        """List all available email folders."""
        try:
            if not self.email_client:
                raise RuntimeError("Email client not configured")
            
            # Connect if not already connected
            if not hasattr(self.email_client, 'connected') or not self.email_client.connected:
                await self.email_client.connect()
            
            folders = await self.email_client.list_folders()
            
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=json.dumps(folders, indent=2)
                    )
                ]
            )
        except Exception as e:
            logger.error(f"Error getting folders: {e}")
            return CallToolResult(
                content=[
                    TextContent(
                        type="text",
                        text=f"Error getting folders: {str(e)}"
                    )
                ],
                isError=True
            )

    async def run(self) -> None:
        """Run the MCP server."""
        logger.info("Starting EmailMCP server...")
        try:
            async with stdio_server() as (read_stream, write_stream):
                await self.server.run(
                    read_stream,
                    write_stream,
                    InitializationOptions(
                        server_name="email-mcp",
                        server_version="0.1.0",
                        capabilities=self.server.get_capabilities()
                    )
                )
        finally:
            await self._cleanup()
    
    async def _cleanup(self) -> None:
        """Clean up resources."""
        if self.email_client and hasattr(self.email_client, 'disconnect'):
            try:
                await self.email_client.disconnect()
                logger.info("Disconnected from email client")
            except Exception as e:
                logger.warning(f"Error disconnecting email client: {e}")


async def main() -> None:
    """Main entry point for the EmailMCP server."""
    server = EmailMCPServer()
    await server.run()


if __name__ == "__main__":
    asyncio.run(main())