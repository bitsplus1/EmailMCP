"""MCP Protocol Handler for managing MCP protocol compliance and communication."""

import json
import logging
from typing import Dict, Any, Optional, List
from ..models.mcp_models import MCPRequest, MCPResponse
from ..models.exceptions import ValidationError


logger = logging.getLogger(__name__)


class MCPProtocolHandler:
    """
    Handles MCP protocol compliance, handshake, and message formatting.
    
    This class manages the Model Context Protocol (MCP) communication layer,
    including client handshake, capability negotiation, request parsing,
    and response formatting according to MCP specifications.
    """
    
    # MCP Protocol version
    PROTOCOL_VERSION = "2024-11-05"
    
    # Server capabilities
    SERVER_CAPABILITIES = {
        "tools": [
            {
                "name": "list_emails",
                "description": "List emails from specified folders with filtering options",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "folder": {
                            "type": "string",
                            "description": "Folder name to list emails from"
                        },
                        "unread_only": {
                            "type": "boolean",
                            "description": "Filter to show only unread emails",
                            "default": False
                        },
                        "limit": {
                            "type": "integer",
                            "description": "Maximum number of emails to return",
                            "default": 50,
                            "minimum": 1,
                            "maximum": 1000
                        }
                    }
                }
            },
            {
                "name": "get_email",
                "description": "Retrieve detailed information for a specific email by ID",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "email_id": {
                            "type": "string",
                            "description": "Unique identifier of the email to retrieve"
                        }
                    },
                    "required": ["email_id"]
                }
            },
            {
                "name": "search_emails",
                "description": "Search emails based on user-defined queries",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "Search query to find emails"
                        },
                        "folder": {
                            "type": "string",
                            "description": "Folder to limit search to (optional)"
                        },
                        "limit": {
                            "type": "integer",
                            "description": "Maximum number of results to return",
                            "default": 50,
                            "minimum": 1,
                            "maximum": 1000
                        }
                    },
                    "required": ["query"]
                }
            },
            {
                "name": "get_folders",
                "description": "List all available email folders in Outlook",
                "inputSchema": {
                    "type": "object",
                    "properties": {}
                }
            }
        ]
    }
    
    # Standard MCP error codes
    ERROR_CODES = {
        "PARSE_ERROR": -32700,
        "INVALID_REQUEST": -32600,
        "METHOD_NOT_FOUND": -32601,
        "INVALID_PARAMS": -32602,
        "INTERNAL_ERROR": -32603,
        "SERVER_ERROR": -32000,
        "OUTLOOK_CONNECTION_ERROR": -32001,
        "OUTLOOK_ACCESS_ERROR": -32002,
        "VALIDATION_ERROR": -32003
    }
    
    def __init__(self):
        """Initialize the MCP protocol handler."""
        self.client_info: Optional[Dict[str, Any]] = None
        self.session_active = False
        logger.info("MCP Protocol Handler initialized")
    
    def handle_handshake(self, client_info: Dict[str, Any]) -> Dict[str, Any]:
        """
        Handle MCP handshake and capability negotiation.
        
        Args:
            client_info: Client information including protocol version and capabilities
            
        Returns:
            Server handshake response with capabilities
            
        Raises:
            ValidationError: If client info is invalid or protocol version unsupported
        """
        logger.info(f"Handling handshake from client: {client_info}")
        
        try:
            # Validate client info structure
            self._validate_client_info(client_info)
            
            # Check protocol version compatibility
            client_version = client_info.get("protocolVersion")
            if not self._is_protocol_version_supported(client_version):
                raise ValidationError(f"Unsupported protocol version: {client_version}")
            
            # Store client information
            self.client_info = client_info
            self.session_active = True
            
            # Prepare server response
            handshake_response = {
                "protocolVersion": self.PROTOCOL_VERSION,
                "capabilities": self.SERVER_CAPABILITIES,
                "serverInfo": {
                    "name": "outlook-mcp-server",
                    "version": "1.0.0",
                    "description": "MCP server for Microsoft Outlook email operations"
                }
            }
            
            logger.info("Handshake completed successfully")
            return handshake_response
            
        except Exception as e:
            logger.error(f"Handshake failed: {str(e)}")
            self.session_active = False
            raise
    
    def process_request(self, request: MCPRequest) -> MCPResponse:
        """
        Process an MCP request and return appropriate response.
        
        Args:
            request: Validated MCP request object
            
        Returns:
            MCP response object with result or error
        """
        logger.debug(f"Processing request: {request.method} (ID: {request.id})")
        
        try:
            # Validate session is active
            if not self.session_active:
                return MCPResponse.create_error(
                    request.id,
                    self.ERROR_CODES["SERVER_ERROR"],
                    "No active session. Handshake required."
                )
            
            # Validate request format
            request.validate()
            
            # Check if method is supported
            if not self._is_method_supported(request.method):
                return MCPResponse.create_error(
                    request.id,
                    self.ERROR_CODES["METHOD_NOT_FOUND"],
                    f"Method '{request.method}' not found"
                )
            
            # Validate parameters for the specific method
            validation_error = self._validate_method_params(request.method, request.params)
            if validation_error:
                return MCPResponse.create_error(
                    request.id,
                    self.ERROR_CODES["INVALID_PARAMS"],
                    validation_error
                )
            
            logger.debug(f"Request validation passed for method: {request.method}")
            return None  # This will be handled by the request router
            
        except ValidationError as e:
            logger.warning(f"Request validation failed: {str(e)}")
            return MCPResponse.create_error(
                request.id,
                self.ERROR_CODES["VALIDATION_ERROR"],
                str(e)
            )
        except Exception as e:
            logger.error(f"Unexpected error processing request: {str(e)}")
            return MCPResponse.create_error(
                request.id,
                self.ERROR_CODES["INTERNAL_ERROR"],
                "Internal server error"
            )
    
    def format_response(self, data: Any, request_id: str) -> Dict[str, Any]:
        """
        Format successful response data according to MCP protocol.
        
        Args:
            data: Response data to format
            request_id: Original request ID
            
        Returns:
            Formatted MCP response dictionary
        """
        try:
            response = MCPResponse.create_success(request_id, data)
            formatted = response.to_dict()
            
            logger.debug(f"Formatted successful response for request {request_id}")
            return formatted
            
        except Exception as e:
            logger.error(f"Error formatting response: {str(e)}")
            # Return error response if formatting fails
            error_response = MCPResponse.create_error(
                request_id,
                self.ERROR_CODES["INTERNAL_ERROR"],
                "Error formatting response"
            )
            return error_response.to_dict()
    
    def format_error(self, error: Exception, request_id: str) -> Dict[str, Any]:
        """
        Format error response according to MCP protocol.
        
        Args:
            error: Exception that occurred
            request_id: Original request ID
            
        Returns:
            Formatted MCP error response dictionary
        """
        try:
            # Determine appropriate error code and message
            error_code, error_message, error_data = self._categorize_error(error)
            
            response = MCPResponse.create_error(
                request_id,
                error_code,
                error_message,
                error_data
            )
            
            formatted = response.to_dict()
            logger.debug(f"Formatted error response for request {request_id}: {error_message}")
            return formatted
            
        except Exception as format_error:
            logger.error(f"Error formatting error response: {str(format_error)}")
            # Fallback error response
            fallback_response = MCPResponse.create_error(
                request_id,
                self.ERROR_CODES["INTERNAL_ERROR"],
                "Internal server error"
            )
            return fallback_response.to_dict()
    
    def _validate_client_info(self, client_info: Dict[str, Any]) -> None:
        """Validate client information structure."""
        if not isinstance(client_info, dict):
            raise ValidationError("Client info must be a dictionary")
        
        if "protocolVersion" not in client_info:
            raise ValidationError("Client info must include protocolVersion")
        
        if not isinstance(client_info["protocolVersion"], str):
            raise ValidationError("Protocol version must be a string")
    
    def _is_protocol_version_supported(self, version: str) -> bool:
        """Check if the client protocol version is supported."""
        # For now, we support the current version
        # In the future, this could support multiple versions
        return version == self.PROTOCOL_VERSION
    
    def _is_method_supported(self, method: str) -> bool:
        """Check if the requested method is supported."""
        supported_methods = {tool["name"] for tool in self.SERVER_CAPABILITIES["tools"]}
        return method in supported_methods
    
    def _validate_method_params(self, method: str, params: Dict[str, Any]) -> Optional[str]:
        """
        Validate parameters for a specific method.
        
        Returns:
            Error message if validation fails, None if valid
        """
        # Find the tool schema for this method
        tool_schema = None
        for tool in self.SERVER_CAPABILITIES["tools"]:
            if tool["name"] == method:
                tool_schema = tool["inputSchema"]
                break
        
        if not tool_schema:
            return f"No schema found for method: {method}"
        
        # Validate required parameters
        required_params = tool_schema.get("properties", {})
        required_fields = tool_schema.get("required", [])
        
        for field in required_fields:
            if field not in params:
                return f"Missing required parameter: {field}"
        
        # Validate parameter types and constraints
        for param_name, param_value in params.items():
            if param_name in required_params:
                param_schema = required_params[param_name]
                error = self._validate_param_value(param_name, param_value, param_schema)
                if error:
                    return error
        
        return None
    
    def _validate_param_value(self, name: str, value: Any, schema: Dict[str, Any]) -> Optional[str]:
        """Validate a single parameter value against its schema."""
        param_type = schema.get("type")
        
        # Type validation
        if param_type == "string" and not isinstance(value, str):
            return f"Parameter '{name}' must be a string"
        elif param_type == "integer" and not isinstance(value, int):
            return f"Parameter '{name}' must be an integer"
        elif param_type == "boolean" and not isinstance(value, bool):
            return f"Parameter '{name}' must be a boolean"
        
        # Constraint validation for integers
        if param_type == "integer":
            minimum = schema.get("minimum")
            maximum = schema.get("maximum")
            
            if minimum is not None and value < minimum:
                return f"Parameter '{name}' must be at least {minimum}"
            if maximum is not None and value > maximum:
                return f"Parameter '{name}' must be at most {maximum}"
        
        # String length validation
        if param_type == "string":
            if name == "query" and len(value.strip()) == 0:
                return f"Parameter '{name}' cannot be empty"
        
        return None
    
    def _categorize_error(self, error: Exception) -> tuple[int, str, Optional[Dict[str, Any]]]:
        """
        Categorize an error and return appropriate MCP error code, message, and data.
        
        Returns:
            Tuple of (error_code, error_message, error_data)
        """
        error_type = type(error).__name__
        error_message = str(error)
        
        # Categorize based on exception type
        if isinstance(error, ValidationError):
            return (
                self.ERROR_CODES["VALIDATION_ERROR"],
                error_message,
                {"type": error_type}
            )
        elif "outlook" in error_message.lower() or "com" in error_message.lower():
            if "connection" in error_message.lower():
                return (
                    self.ERROR_CODES["OUTLOOK_CONNECTION_ERROR"],
                    "Failed to connect to Outlook",
                    {"type": error_type, "details": error_message}
                )
            else:
                return (
                    self.ERROR_CODES["OUTLOOK_ACCESS_ERROR"],
                    "Outlook access error",
                    {"type": error_type, "details": error_message}
                )
        else:
            return (
                self.ERROR_CODES["SERVER_ERROR"],
                "Server error occurred",
                {"type": error_type, "details": error_message}
            )
    
    def get_server_info(self) -> Dict[str, Any]:
        """Get server information and capabilities."""
        return {
            "name": "outlook-mcp-server",
            "version": "1.0.0",
            "protocolVersion": self.PROTOCOL_VERSION,
            "capabilities": self.SERVER_CAPABILITIES,
            "description": "MCP server for Microsoft Outlook email operations"
        }
    
    def is_session_active(self) -> bool:
        """Check if there's an active client session."""
        return self.session_active
    
    def close_session(self) -> None:
        """Close the current client session."""
        self.client_info = None
        self.session_active = False
        logger.info("Client session closed")