"""Comprehensive error handling system for Outlook MCP Server."""

import logging
import traceback
import time
from typing import Dict, Any, Optional, Callable
from dataclasses import dataclass
from enum import Enum

from .models.exceptions import (
    OutlookMCPError,
    OutlookConnectionError,
    ValidationError,
    EmailNotFoundError,
    FolderNotFoundError,
    InvalidParameterError,
    SearchError,
    PermissionError,
    TimeoutError,
    MethodNotFoundError
)


class ErrorSeverity(Enum):
    """Error severity levels for categorization."""
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


@dataclass
class ErrorContext:
    """Context information for error handling."""
    request_id: str
    method: str
    parameters: Dict[str, Any]
    timestamp: float
    user_agent: Optional[str] = None
    client_info: Optional[Dict[str, Any]] = None


class ErrorHandler:
    """Comprehensive error handling system with categorized error processing."""
    
    def __init__(self, logger: Optional[logging.Logger] = None):
        """Initialize the error handler.
        
        Args:
            logger: Optional logger instance. If not provided, creates a new one.
        """
        self.logger = logger or self._setup_logger()
        self._retry_strategies: Dict[type, Callable] = {}
        self._error_categories: Dict[type, ErrorSeverity] = self._setup_error_categories()
        self._error_stats = {
            "total_errors": 0,
            "errors_by_type": {},
            "errors_by_severity": {"low": 0, "medium": 0, "high": 0, "critical": 0},
            "recovery_attempts": 0,
            "successful_recoveries": 0
        }
        
    def _setup_logger(self) -> logging.Logger:
        """Set up structured logging for error handling."""
        logger = logging.getLogger("outlook_mcp_server.error_handler")
        if not logger.handlers:
            handler = logging.StreamHandler()
            # Use structured JSON-like formatting for better parsing
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)
        return logger
    
    def configure_structured_logging(self, use_json: bool = True, log_file: Optional[str] = None) -> None:
        """Configure structured logging with optional JSON format and file output.
        
        Args:
            use_json: Whether to use JSON formatting for log messages
            log_file: Optional file path for log output
        """
        # Remove existing handlers
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
        
        # Create new handlers
        handlers = []
        
        # Console handler
        console_handler = logging.StreamHandler()
        handlers.append(console_handler)
        
        # File handler if specified
        if log_file:
            file_handler = logging.FileHandler(log_file)
            handlers.append(file_handler)
        
        # Configure formatters
        if use_json:
            import json
            
            class JSONFormatter(logging.Formatter):
                def format(self, record):
                    log_entry = {
                        "timestamp": self.formatTime(record),
                        "level": record.levelname,
                        "logger": record.name,
                        "message": record.getMessage(),
                    }
                    
                    # Add extra fields if present
                    if hasattr(record, 'extra') and record.extra:
                        log_entry.update(record.extra)
                    
                    return json.dumps(log_entry)
            
            formatter = JSONFormatter()
        else:
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
        
        # Apply formatter to all handlers
        for handler in handlers:
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        
        self.logger.setLevel(logging.INFO)
    
    def _setup_error_categories(self) -> Dict[type, ErrorSeverity]:
        """Set up error severity categories."""
        return {
            ValidationError: ErrorSeverity.LOW,
            InvalidParameterError: ErrorSeverity.LOW,
            MethodNotFoundError: ErrorSeverity.LOW,
            EmailNotFoundError: ErrorSeverity.MEDIUM,
            FolderNotFoundError: ErrorSeverity.MEDIUM,
            SearchError: ErrorSeverity.MEDIUM,
            PermissionError: ErrorSeverity.HIGH,
            TimeoutError: ErrorSeverity.HIGH,
            OutlookConnectionError: ErrorSeverity.CRITICAL,
        }
    
    def register_retry_strategy(self, error_type: type, strategy: Callable) -> None:
        """Register a retry strategy for specific error types.
        
        Args:
            error_type: The exception type to handle
            strategy: Callable that implements retry logic
        """
        self._retry_strategies[error_type] = strategy
    
    def handle_error(self, error: Exception, context: ErrorContext) -> Dict[str, Any]:
        """Handle an error with comprehensive processing.
        
        Args:
            error: The exception that occurred
            context: Context information about the error
            
        Returns:
            Structured error response dictionary
        """
        # Update error statistics
        self._update_error_stats(error)
        
        # Log the error with context
        self._log_error(error, context)
        
        # Attempt error recovery if applicable
        recovery_result = self._attempt_recovery(error, context)
        if recovery_result:
            self._error_stats["successful_recoveries"] += 1
            return recovery_result
        
        # Generate structured error response
        return self._generate_error_response(error, context)
    
    def _log_error(self, error: Exception, context: ErrorContext) -> None:
        """Log error with contextual information.
        
        Args:
            error: The exception that occurred
            context: Context information about the error
        """
        severity = self._get_error_severity(error)
        
        log_data = {
            "error_type": error.__class__.__name__,
            "error_message": str(error),
            "severity": severity.value,
            "request_id": context.request_id,
            "method": context.method,
            "parameters": context.parameters,
            "timestamp": context.timestamp,
            "traceback": traceback.format_exc() if severity in [ErrorSeverity.HIGH, ErrorSeverity.CRITICAL] else None
        }
        
        if context.user_agent:
            log_data["user_agent"] = context.user_agent
        if context.client_info:
            log_data["client_info"] = context.client_info
        
        # Log at appropriate level based on severity
        if severity == ErrorSeverity.CRITICAL:
            self.logger.critical("Critical error occurred", extra=log_data)
        elif severity == ErrorSeverity.HIGH:
            self.logger.error("High severity error occurred", extra=log_data)
        elif severity == ErrorSeverity.MEDIUM:
            self.logger.warning("Medium severity error occurred", extra=log_data)
        else:
            self.logger.info("Low severity error occurred", extra=log_data)
    
    def _get_error_severity(self, error: Exception) -> ErrorSeverity:
        """Get the severity level for an error.
        
        Args:
            error: The exception to categorize
            
        Returns:
            ErrorSeverity level
        """
        error_type = type(error)
        return self._error_categories.get(error_type, ErrorSeverity.MEDIUM)
    
    def _attempt_recovery(self, error: Exception, context: ErrorContext) -> Optional[Dict[str, Any]]:
        """Attempt to recover from transient errors.
        
        Args:
            error: The exception that occurred
            context: Context information about the error
            
        Returns:
            Recovery result if successful, None otherwise
        """
        self._error_stats["recovery_attempts"] += 1
        error_type = type(error)
        
        # Check if we have a retry strategy for this error type
        if error_type in self._retry_strategies:
            try:
                strategy = self._retry_strategies[error_type]
                return strategy(error, context)
            except Exception as retry_error:
                self.logger.warning(
                    f"Retry strategy failed for {error_type.__name__}: {retry_error}",
                    extra={"original_error": str(error), "request_id": context.request_id}
                )
        
        # Built-in recovery for specific error types
        if isinstance(error, OutlookConnectionError):
            return self._recover_outlook_connection(error, context)
        elif isinstance(error, TimeoutError):
            return self._recover_timeout(error, context)
        
        return None
    
    def _recover_outlook_connection(self, error: OutlookConnectionError, context: ErrorContext) -> Optional[Dict[str, Any]]:
        """Attempt to recover from Outlook connection errors.
        
        Args:
            error: The connection error
            context: Error context
            
        Returns:
            Recovery result if successful, None otherwise
        """
        self.logger.info(
            "Attempting Outlook connection recovery",
            extra={"request_id": context.request_id}
        )
        
        # In a real implementation, this would attempt to reconnect to Outlook
        # For now, we'll just log the attempt and return None to indicate failure
        return None
    
    def _recover_timeout(self, error: TimeoutError, context: ErrorContext) -> Optional[Dict[str, Any]]:
        """Attempt to recover from timeout errors.
        
        Args:
            error: The timeout error
            context: Error context
            
        Returns:
            Recovery result if successful, None otherwise
        """
        self.logger.info(
            "Timeout recovery not implemented for this operation",
            extra={"request_id": context.request_id, "operation": error.details.get("operation")}
        )
        
        # Timeout recovery would depend on the specific operation
        return None
    
    def _generate_error_response(self, error: Exception, context: ErrorContext) -> Dict[str, Any]:
        """Generate a structured error response.
        
        Args:
            error: The exception that occurred
            context: Context information about the error
            
        Returns:
            Structured error response dictionary
        """
        # If it's already an OutlookMCPError, use its structured format
        if isinstance(error, OutlookMCPError):
            error_dict = error.to_dict()
        else:
            # Convert generic exceptions to structured format
            error_dict = {
                "code": -32000,  # Generic server error
                "message": str(error),
                "data": {
                    "type": error.__class__.__name__,
                    "details": {}
                }
            }
        
        # Add context information
        error_dict["data"]["context"] = {
            "request_id": context.request_id,
            "method": context.method,
            "timestamp": context.timestamp
        }
        
        # Add severity information
        severity = self._get_error_severity(error)
        error_dict["data"]["severity"] = severity.value
        
        return {
            "jsonrpc": "2.0",
            "id": context.request_id,
            "error": error_dict
        }
    
    def create_context(self, request_id: str, method: str, parameters: Dict[str, Any], 
                      user_agent: Optional[str] = None, 
                      client_info: Optional[Dict[str, Any]] = None) -> ErrorContext:
        """Create an error context for a request.
        
        Args:
            request_id: Unique identifier for the request
            method: The method being called
            parameters: Method parameters
            user_agent: Optional user agent string
            client_info: Optional client information
            
        Returns:
            ErrorContext instance
        """
        return ErrorContext(
            request_id=request_id,
            method=method,
            parameters=parameters,
            timestamp=time.time(),
            user_agent=user_agent,
            client_info=client_info
        )
    
    def _update_error_stats(self, error: Exception) -> None:
        """Update error statistics.
        
        Args:
            error: The exception that occurred
        """
        self._error_stats["total_errors"] += 1
        
        # Update error type statistics
        error_type = error.__class__.__name__
        self._error_stats["errors_by_type"][error_type] = (
            self._error_stats["errors_by_type"].get(error_type, 0) + 1
        )
        
        # Update severity statistics
        severity = self._get_error_severity(error)
        self._error_stats["errors_by_severity"][severity.value] += 1
    
    def get_error_statistics(self) -> Dict[str, Any]:
        """Get error statistics for monitoring.
        
        Returns:
            Dictionary containing error statistics
        """
        return self._error_stats.copy()
    
    def reset_error_statistics(self) -> None:
        """Reset error statistics counters."""
        self._error_stats = {
            "total_errors": 0,
            "errors_by_type": {},
            "errors_by_severity": {"low": 0, "medium": 0, "high": 0, "critical": 0},
            "recovery_attempts": 0,
            "successful_recoveries": 0
        }


# Default retry strategies
def outlook_connection_retry_strategy(error: OutlookConnectionError, context: ErrorContext, 
                                    max_retries: int = 3, base_delay: float = 1.0) -> Optional[Dict[str, Any]]:
    """Default retry strategy for Outlook connection errors with exponential backoff.
    
    Args:
        error: The connection error
        context: Error context
        max_retries: Maximum number of retry attempts
        base_delay: Base delay in seconds for exponential backoff
        
    Returns:
        Recovery result if successful, None otherwise
    """
    import time
    
    # Check if we've already exceeded retry attempts (would need to track this in context)
    retry_count = context.parameters.get('_retry_count', 0)
    
    if retry_count >= max_retries:
        return None
    
    # Calculate delay with exponential backoff
    delay = base_delay * (2 ** retry_count)
    
    # In a real implementation, this would:
    # 1. Wait for the calculated delay
    # 2. Attempt to reconnect to Outlook
    # 3. Return success result if connection is restored
    
    # For now, simulate the delay and return None to indicate failure
    # time.sleep(delay)  # Commented out to avoid blocking in tests
    
    return None


def timeout_retry_strategy(error: TimeoutError, context: ErrorContext,
                         max_retries: int = 2, timeout_multiplier: float = 1.5) -> Optional[Dict[str, Any]]:
    """Default retry strategy for timeout errors with increased timeout.
    
    Args:
        error: The timeout error
        context: Error context
        max_retries: Maximum number of retry attempts
        timeout_multiplier: Multiplier for increasing timeout on retry
        
    Returns:
        Recovery result if successful, None otherwise
    """
    retry_count = context.parameters.get('_retry_count', 0)
    
    if retry_count >= max_retries:
        return None
    
    # Calculate new timeout
    original_timeout = error.details.get('timeout_seconds', 30)
    new_timeout = int(original_timeout * (timeout_multiplier ** (retry_count + 1)))
    
    # In a real implementation, this would:
    # 1. Retry the operation with increased timeout
    # 2. Return success result if operation completes
    
    # For now, return None to indicate no recovery
    return None


def create_exponential_backoff_strategy(max_retries: int = 3, base_delay: float = 1.0, 
                                      max_delay: float = 60.0) -> Callable:
    """Create a configurable exponential backoff retry strategy.
    
    Args:
        max_retries: Maximum number of retry attempts
        base_delay: Base delay in seconds
        max_delay: Maximum delay in seconds
        
    Returns:
        Retry strategy function
    """
    def strategy(error: Exception, context: ErrorContext) -> Optional[Dict[str, Any]]:
        retry_count = context.parameters.get('_retry_count', 0)
        
        if retry_count >= max_retries:
            return None
        
        # Calculate delay with exponential backoff, capped at max_delay
        delay = min(base_delay * (2 ** retry_count), max_delay)
        
        # In a real implementation, this would implement the actual retry logic
        # For now, return None to indicate no recovery
        return None
    
    return strategy