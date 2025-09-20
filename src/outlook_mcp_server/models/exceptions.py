"""Exception classes for Outlook MCP Server."""


class OutlookMCPError(Exception):
    """Base exception class for Outlook MCP Server errors."""
    
    def __init__(self, message: str, error_code: int = -32000, details: dict = None):
        super().__init__(message)
        self.message = message
        self.error_code = error_code
        self.details = details or {}

    def to_dict(self) -> dict:
        """Convert exception to dictionary for MCP error response."""
        return {
            "code": self.error_code,
            "message": self.message,
            "data": {
                "type": self.__class__.__name__,
                "details": self.details
            }
        }


class ValidationError(OutlookMCPError):
    """Exception raised for data validation errors."""
    
    def __init__(self, message: str, field: str = None):
        super().__init__(message, error_code=-32602)
        if field:
            self.details["field"] = field


class OutlookConnectionError(OutlookMCPError):
    """Exception raised when Outlook connection fails."""
    
    def __init__(self, message: str = "Could not establish connection to Outlook"):
        super().__init__(message, error_code=-32001)


class EmailNotFoundError(OutlookMCPError):
    """Exception raised when requested email is not found."""
    
    def __init__(self, email_id: str):
        message = f"Email with ID '{email_id}' not found"
        super().__init__(message, error_code=-32003)
        self.details["email_id"] = email_id


class FolderNotFoundError(OutlookMCPError):
    """Exception raised when requested folder is not found."""
    
    def __init__(self, folder_name: str):
        message = f"Folder '{folder_name}' not found"
        super().__init__(message, error_code=-32004)
        self.details["folder_name"] = folder_name


class InvalidParameterError(OutlookMCPError):
    """Exception raised for invalid method parameters."""
    
    def __init__(self, parameter: str, message: str = None):
        if message is None:
            message = f"Invalid parameter: {parameter}"
        super().__init__(message, error_code=-32602)
        self.details["parameter"] = parameter


class SearchError(OutlookMCPError):
    """Exception raised when email search fails."""
    
    def __init__(self, query: str, message: str = "Search operation failed"):
        super().__init__(message, error_code=-32005)
        self.details["query"] = query


class PermissionError(OutlookMCPError):
    """Exception raised when access to resource is denied."""
    
    def __init__(self, resource: str, message: str = None):
        if message is None:
            message = f"Access denied to resource: {resource}"
        super().__init__(message, error_code=-32006)
        self.details["resource"] = resource


class TimeoutError(OutlookMCPError):
    """Exception raised when operation times out."""
    
    def __init__(self, operation: str, timeout_seconds: int):
        message = f"Operation '{operation}' timed out after {timeout_seconds} seconds"
        super().__init__(message, error_code=-32007)
        self.details["operation"] = operation
        self.details["timeout_seconds"] = timeout_seconds


class MethodNotFoundError(OutlookMCPError):
    """Exception raised when requested method is not found."""
    
    def __init__(self, method_name: str, message: str = None):
        if message is None:
            message = f"Method '{method_name}' not found"
        super().__init__(message, error_code=-32601)
        self.details["method_name"] = method_name