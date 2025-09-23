"""Request router for MCP method dispatch and parameter validation."""

import logging
from typing import Any, Dict, Callable, Optional, List
from ..models.mcp_models import MCPRequest, MCPResponse
from ..models.exceptions import ValidationError, MethodNotFoundError


logger = logging.getLogger(__name__)


class RequestRouter:
    """Router for MCP requests with method dispatch and parameter validation."""
    
    def __init__(self):
        """Initialize the request router."""
        self._handlers: Dict[str, Callable] = {}
        self._parameter_schemas: Dict[str, Dict[str, Any]] = {}
        self._setup_parameter_schemas()
    
    def _setup_parameter_schemas(self) -> None:
        """Set up parameter validation schemas for each method."""
        self._parameter_schemas = {
            "list_inbox_emails": {
                "unread_only": {"type": bool, "required": False, "default": False},
                "limit": {"type": int, "required": False, "default": 50, "min": 1, "max": 1000}
            },
            "list_emails": {
                "folder_id": {"type": str, "required": True},
                "unread_only": {"type": bool, "required": False, "default": False},
                "limit": {"type": int, "required": False, "default": 50, "min": 1, "max": 1000}
            },
            "get_email": {
                "email_id": {"type": str, "required": True}
            },
            "search_emails": {
                "query": {"type": str, "required": True, "min_length": 1, "max_length": 1000},
                "folder_id": {"type": str, "required": False, "default": None},
                "limit": {"type": int, "required": False, "default": 50, "min": 1, "max": 1000}
            },
            "get_folders": {
                # No parameters required for get_folders
            },
            "send_email": {
                "to_recipients": {"type": list, "required": True, "min_length": 1},
                "subject": {"type": str, "required": True, "min_length": 1, "max_length": 255},
                "body": {"type": str, "required": True, "min_length": 1},
                "cc_recipients": {"type": list, "required": False, "default": None},
                "bcc_recipients": {"type": list, "required": False, "default": None},
                "body_format": {"type": str, "required": False, "default": "html"},
                "importance": {"type": str, "required": False, "default": "normal"},
                "attachments": {"type": list, "required": False, "default": None},
                "save_to_sent_items": {"type": bool, "required": False, "default": True}
            },
            "debug_folder_names": {
                # No parameters required for debug_folder_names
            }
        }
    
    def register_handler(self, method: str, handler: Callable) -> None:
        """
        Register a handler for a specific method.
        
        Args:
            method: The method name to register
            handler: The callable handler function
            
        Raises:
            ValidationError: If method name is invalid or handler is not callable
        """
        if not method or not isinstance(method, str):
            raise ValidationError("Method name must be a non-empty string")
        
        if not callable(handler):
            raise ValidationError("Handler must be callable")
        
        logger.debug(f"Registering handler for method: {method}")
        self._handlers[method] = handler
    
    def route_request(self, request: MCPRequest) -> Any:
        """
        Route an MCP request to the appropriate handler.
        
        Args:
            request: The MCP request to route
            
        Returns:
            Any: The result from the handler
            
        Raises:
            MethodNotFoundError: If method is not registered
            ValidationError: If parameters are invalid
        """
        logger.debug(f"Routing request for method: {request.method}")
        
        # Check if method is registered
        if request.method not in self._handlers:
            available_methods = list(self._handlers.keys())
            raise MethodNotFoundError(
                request.method, 
                f"Method '{request.method}' not found. Available methods: {available_methods}"
            )
        
        # Validate and process parameters
        validated_params = self.validate_params(request.method, request.params)
        
        # Get the handler and call it
        handler = self._handlers[request.method]
        
        try:
            logger.debug(f"Calling handler for {request.method} with params: {validated_params}")
            result = handler(**validated_params)
            logger.debug(f"Handler for {request.method} completed successfully")
            return result
            
        except Exception as e:
            logger.error(f"Handler for {request.method} failed: {str(e)}")
            raise
    
    def validate_params(self, method: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Validate parameters for a specific method.
        
        Args:
            method: The method name
            params: The parameters to validate
            
        Returns:
            Dict[str, Any]: Validated and processed parameters
            
        Raises:
            ValidationError: If parameters are invalid
            MethodNotFoundError: If method schema is not found
        """
        if method not in self._parameter_schemas:
            raise MethodNotFoundError(method, f"No parameter schema found for method '{method}'")
        
        schema = self._parameter_schemas[method]
        validated_params = {}
        
        # Check for unknown parameters
        unknown_params = set(params.keys()) - set(schema.keys())
        if unknown_params:
            raise ValidationError(
                f"Unknown parameters for method '{method}': {list(unknown_params)}. "
                f"Valid parameters: {list(schema.keys())}"
            )
        
        # Validate each parameter in the schema
        for param_name, param_config in schema.items():
            param_value = params.get(param_name)
            
            # Check if required parameter is missing
            if param_config.get("required", False) and param_value is None:
                raise ValidationError(f"Required parameter '{param_name}' is missing for method '{method}'")
            
            # Use default value if parameter is not provided
            if param_value is None and "default" in param_config:
                validated_params[param_name] = param_config["default"]
                continue
            
            # Skip validation if parameter is None and not required
            if param_value is None:
                validated_params[param_name] = None
                continue
            
            # Validate parameter type
            expected_type = param_config["type"]
            if not isinstance(param_value, expected_type):
                raise ValidationError(
                    f"Parameter '{param_name}' for method '{method}' must be of type {expected_type.__name__}, "
                    f"got {type(param_value).__name__}"
                )
            
            # Validate string parameters
            if expected_type == str and param_value is not None:
                self._validate_string_param(method, param_name, param_value, param_config)
            
            # Validate integer parameters
            elif expected_type == int and param_value is not None:
                self._validate_int_param(method, param_name, param_value, param_config)
            
            # Validate list parameters
            elif expected_type == list and param_value is not None:
                self._validate_list_param(method, param_name, param_value, param_config)
            
            validated_params[param_name] = param_value
        
        logger.debug(f"Parameters validated for method '{method}': {validated_params}")
        return validated_params
    
    def _validate_string_param(self, method: str, param_name: str, value: str, config: Dict[str, Any]) -> None:
        """Validate string parameter constraints."""
        # Check minimum length
        if "min_length" in config and len(value) < config["min_length"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must be at least "
                f"{config['min_length']} characters long"
            )
        
        # Check maximum length
        if "max_length" in config and len(value) > config["max_length"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must be at most "
                f"{config['max_length']} characters long"
            )
        
        # Validate specific string formats
        if param_name == "email_id":
            self._validate_email_id(value)
        elif param_name == "folder_id":
            self._validate_folder_id(value)
        elif param_name == "query":
            self._validate_search_query(value)
        elif param_name == "body_format":
            self._validate_body_format(value)
        elif param_name == "importance":
            self._validate_importance(value)
    
    def _validate_int_param(self, method: str, param_name: str, value: int, config: Dict[str, Any]) -> None:
        """Validate integer parameter constraints."""
        # Check minimum value
        if "min" in config and value < config["min"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must be at least {config['min']}"
            )
        
        # Check maximum value
        if "max" in config and value > config["max"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must be at most {config['max']}"
            )
    
    def _validate_email_id(self, email_id: str) -> None:
        """Validate email ID format."""
        if not email_id or not email_id.strip():
            raise ValidationError("Email ID cannot be empty")
        
        # Email IDs should be reasonable length and not contain certain characters
        if len(email_id) > 500:
            raise ValidationError("Email ID is too long")
        
        # Check for potentially dangerous characters
        dangerous_chars = ['<', '>', '"', "'", '&', '\n', '\r', '\t']
        if any(char in email_id for char in dangerous_chars):
            raise ValidationError("Email ID contains invalid characters")
    
    def _validate_folder_name(self, folder_name: str) -> None:
        """Validate folder name format - allows Unicode characters including Chinese."""
        if not folder_name or not folder_name.strip():
            raise ValidationError("Folder name cannot be empty")
        
        # Folder names should be reasonable length
        if len(folder_name) > 255:
            raise ValidationError("Folder name is too long")
        
        # Check for invalid folder name characters (Windows/Outlook restrictions)
        # Only reject characters that are actually problematic for Outlook
        # Allow Unicode characters (including Chinese, Japanese, Korean, etc.)
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        problematic_chars = [char for char in folder_name if char in invalid_chars]
        if problematic_chars:
            raise ValidationError(f"Folder name contains invalid characters: {problematic_chars}")
        
        # Check for control characters (ASCII 0-31) but allow Unicode
        control_chars = [char for char in folder_name if ord(char) < 32]
        if control_chars:
            raise ValidationError(f"Folder name contains control characters")
    
    def _validate_folder_id(self, folder_id: str) -> None:
        """Validate folder ID format."""
        if not folder_id or not folder_id.strip():
            raise ValidationError("Folder ID cannot be empty")
        
        # Folder IDs should be reasonable length (Outlook folder IDs are typically long hex strings)
        if len(folder_id) > 500:
            raise ValidationError("Folder ID is too long")
        
        # Check for potentially dangerous characters
        dangerous_chars = ['<', '>', '"', "'", '&', '\n', '\r', '\t']
        if any(char in folder_id for char in dangerous_chars):
            raise ValidationError("Folder ID contains invalid characters")
    
    def _validate_search_query(self, query: str) -> None:
        """Validate search query format."""
        query = query.strip()
        if not query:
            raise ValidationError("Search query cannot be empty")
        
        # Check for reasonable query length
        if len(query) > 1000:
            raise ValidationError("Search query is too long")
        
        # Basic validation - could be enhanced with more sophisticated query parsing
        if len(query) < 1:
            raise ValidationError("Search query must be at least 1 character long")
    
    def _validate_list_param(self, method: str, param_name: str, value: list, config: Dict[str, Any]) -> None:
        """Validate list parameter constraints."""
        # Check minimum length
        if "min_length" in config and len(value) < config["min_length"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must have at least "
                f"{config['min_length']} items"
            )
        
        # Check maximum length
        if "max_length" in config and len(value) > config["max_length"]:
            raise ValidationError(
                f"Parameter '{param_name}' for method '{method}' must have at most "
                f"{config['max_length']} items"
            )
        
        # Validate email recipient lists
        if param_name in ["to_recipients", "cc_recipients", "bcc_recipients"]:
            self._validate_email_list(param_name, value)
        
        # Validate attachment lists
        elif param_name == "attachments":
            self._validate_attachment_list(value)
    
    def _validate_email_list(self, param_name: str, email_list: list) -> None:
        """Validate list of email addresses."""
        if not email_list:
            return
        
        if len(email_list) > 100:  # Reasonable limit
            raise ValidationError(f"{param_name} cannot have more than 100 recipients")
        
        import re
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        
        for i, email in enumerate(email_list):
            if not isinstance(email, str):
                raise ValidationError(f"{param_name}[{i}] must be a string")
            
            email = email.strip()
            if not email:
                raise ValidationError(f"{param_name}[{i}] cannot be empty")
            
            if not re.match(email_pattern, email):
                raise ValidationError(f"{param_name}[{i}] is not a valid email address: {email}")
    
    def _validate_attachment_list(self, attachment_list: list) -> None:
        """Validate list of attachment file paths."""
        if not attachment_list:
            return
        
        if len(attachment_list) > 20:  # Reasonable limit
            raise ValidationError("Cannot attach more than 20 files")
        
        for i, attachment in enumerate(attachment_list):
            if not isinstance(attachment, str):
                raise ValidationError(f"attachments[{i}] must be a string file path")
            
            if not attachment.strip():
                raise ValidationError(f"attachments[{i}] cannot be empty")
            
            # Basic path validation
            if len(attachment) > 500:
                raise ValidationError(f"attachments[{i}] path is too long")
            
            # Check for dangerous characters
            dangerous_chars = ['<', '>', '"', '|', '\n', '\r', '\t']
            if any(char in attachment for char in dangerous_chars):
                raise ValidationError(f"attachments[{i}] contains invalid characters")
    
    def _validate_body_format(self, body_format: str) -> None:
        """Validate email body format."""
        valid_formats = ["html", "text", "rtf"]
        if body_format.lower() not in valid_formats:
            raise ValidationError(f"body_format must be one of: {valid_formats}")
    
    def _validate_importance(self, importance: str) -> None:
        """Validate email importance level."""
        valid_importance = ["low", "normal", "high"]
        if importance.lower() not in valid_importance:
            raise ValidationError(f"importance must be one of: {valid_importance}")
    
    def get_registered_methods(self) -> List[str]:
        """
        Get list of registered method names.
        
        Returns:
            List[str]: List of registered method names
        """
        return list(self._handlers.keys())
    
    def is_method_registered(self, method: str) -> bool:
        """
        Check if a method is registered.
        
        Args:
            method: The method name to check
            
        Returns:
            bool: True if method is registered, False otherwise
        """
        return method in self._handlers
    
    def get_method_schema(self, method: str) -> Optional[Dict[str, Any]]:
        """
        Get parameter schema for a method.
        
        Args:
            method: The method name
            
        Returns:
            Optional[Dict[str, Any]]: Parameter schema or None if method not found
        """
        return self._parameter_schemas.get(method)
    
    def unregister_handler(self, method: str) -> bool:
        """
        Unregister a handler for a method.
        
        Args:
            method: The method name to unregister
            
        Returns:
            bool: True if method was unregistered, False if it wasn't registered
        """
        if method in self._handlers:
            del self._handlers[method]
            logger.debug(f"Unregistered handler for method: {method}")
            return True
        return False