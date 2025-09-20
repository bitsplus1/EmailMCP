"""MCP protocol models for request and response handling."""

from dataclasses import dataclass
from typing import Any, Optional, Dict
import re
from .exceptions import ValidationError


@dataclass
class MCPRequest:
    """Model for MCP protocol requests."""
    
    jsonrpc: str
    id: str
    method: str
    params: Dict[str, Any]

    def __post_init__(self):
        """Validate MCP request after initialization."""
        self.validate()

    def validate(self) -> None:
        """Validate MCP request fields."""
        if self.jsonrpc != "2.0":
            raise ValidationError("JSON-RPC version must be '2.0'")
        
        if not self.id or not isinstance(self.id, str):
            raise ValidationError("Request ID must be a non-empty string")
        
        if not self.method or not isinstance(self.method, str):
            raise ValidationError("Method must be a non-empty string")
        
        if not isinstance(self.params, dict):
            raise ValidationError("Params must be a dictionary")
        
        # Validate method name format
        if not self._is_valid_method_name(self.method):
            raise ValidationError("Invalid method name format")

    @staticmethod
    def _is_valid_method_name(method: str) -> bool:
        """Validate method name format."""
        if not method or not isinstance(method, str):
            return False
        
        # Method names should contain only alphanumeric characters, underscores, and dots
        pattern = r'^[a-zA-Z][a-zA-Z0-9_\.]*$'
        return re.match(pattern, method) is not None

    @staticmethod
    def validate_search_query(query: str) -> bool:
        """Validate search query format."""
        if not query or not isinstance(query, str):
            return False
        
        # Search queries should be non-empty and have reasonable length
        query = query.strip()
        return len(query) > 0 and len(query) <= 1000

    def to_dict(self) -> dict:
        """Convert request to dictionary for JSON serialization."""
        return {
            "jsonrpc": self.jsonrpc,
            "id": self.id,
            "method": self.method,
            "params": self.params
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'MCPRequest':
        """Create MCPRequest instance from dictionary."""
        return cls(
            jsonrpc=data["jsonrpc"],
            id=data["id"],
            method=data["method"],
            params=data.get("params", {})
        )


@dataclass
class MCPResponse:
    """Model for MCP protocol responses."""
    
    jsonrpc: str
    id: str
    result: Optional[Any] = None
    error: Optional[Dict[str, Any]] = None

    def __post_init__(self):
        """Validate MCP response after initialization."""
        self.validate()

    def validate(self) -> None:
        """Validate MCP response fields."""
        if self.jsonrpc != "2.0":
            raise ValidationError("JSON-RPC version must be '2.0'")
        
        if not self.id or not isinstance(self.id, str):
            raise ValidationError("Response ID must be a non-empty string")
        
        # Either result or error must be present, but not both
        if self.result is not None and self.error is not None:
            raise ValidationError("Response cannot have both result and error")
        
        if self.result is None and self.error is None:
            raise ValidationError("Response must have either result or error")
        
        # Validate error format if present
        if self.error is not None:
            self._validate_error_format()

    def _validate_error_format(self) -> None:
        """Validate error object format."""
        if not isinstance(self.error, dict):
            raise ValidationError("Error must be a dictionary")
        
        if "code" not in self.error or not isinstance(self.error["code"], int):
            raise ValidationError("Error must have an integer code")
        
        if "message" not in self.error or not isinstance(self.error["message"], str):
            raise ValidationError("Error must have a string message")

    def to_dict(self) -> dict:
        """Convert response to dictionary for JSON serialization."""
        response = {
            "jsonrpc": self.jsonrpc,
            "id": self.id
        }
        
        if self.result is not None:
            response["result"] = self.result
        
        if self.error is not None:
            response["error"] = self.error
        
        return response

    @classmethod
    def from_dict(cls, data: dict) -> 'MCPResponse':
        """Create MCPResponse instance from dictionary."""
        return cls(
            jsonrpc=data["jsonrpc"],
            id=data["id"],
            result=data.get("result"),
            error=data.get("error")
        )

    @classmethod
    def create_success(cls, request_id: str, result: Any) -> 'MCPResponse':
        """Create a successful response."""
        return cls(
            jsonrpc="2.0",
            id=request_id,
            result=result
        )

    @classmethod
    def create_error(cls, request_id: str, code: int, message: str, data: Optional[Dict[str, Any]] = None) -> 'MCPResponse':
        """Create an error response."""
        error = {
            "code": code,
            "message": message
        }
        
        if data is not None:
            error["data"] = data
        
        return cls(
            jsonrpc="2.0",
            id=request_id,
            error=error
        )