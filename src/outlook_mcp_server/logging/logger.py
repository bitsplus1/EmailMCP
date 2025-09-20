"""
Comprehensive logging system with structured JSON output and performance metrics.

This module implements the logging requirements from the design document:
- Structured JSON logging for all server activities
- Log rotation and performance metrics logging
- Different log levels for debugging and monitoring
- Requirements: 5.2, 5.5, 8.4
"""

import json
import logging
import logging.handlers
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Union
from contextlib import contextmanager


class JSONFormatter(logging.Formatter):
    """Custom formatter that outputs structured JSON log entries."""
    
    def format(self, record: logging.LogRecord) -> str:
        """Format log record as structured JSON."""
        log_entry = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "module": record.module,
            "function": record.funcName,
            "line": record.lineno,
            "thread": record.thread,
            "thread_name": record.threadName,
        }
        
        # Add exception information if present
        if record.exc_info and record.exc_info != (None, None, None):
            exc_type, exc_value, exc_traceback = record.exc_info
            log_entry["exception"] = {
                "type": exc_type.__name__ if exc_type else None,
                "message": str(exc_value) if exc_value else None,
                "traceback": self.formatException(record.exc_info) if exc_traceback else None
            }
        
        # Add extra fields from the record
        extra_fields = {}
        for key, value in record.__dict__.items():
            if key not in {
                'name', 'msg', 'args', 'levelname', 'levelno', 'pathname', 'filename',
                'module', 'exc_info', 'exc_text', 'stack_info', 'lineno', 'funcName',
                'created', 'msecs', 'relativeCreated', 'thread', 'threadName',
                'processName', 'process', 'getMessage'
            }:
                extra_fields[key] = value
        
        if extra_fields:
            log_entry["extra"] = extra_fields
            
        return json.dumps(log_entry, default=str, ensure_ascii=False)


class PerformanceLogger:
    """Logger for performance metrics and timing information."""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        
    def log_request_timing(self, method: str, duration: float, success: bool = True, **kwargs):
        """Log request timing information."""
        self.logger.info(
            f"Request completed: {method}",
            extra={
                "performance": {
                    "method": method,
                    "duration_ms": round(duration * 1000, 2),
                    "success": success,
                    **kwargs
                }
            }
        )
    
    def log_connection_timing(self, operation: str, duration: float, success: bool = True):
        """Log connection operation timing."""
        self.logger.info(
            f"Connection operation: {operation}",
            extra={
                "performance": {
                    "operation": operation,
                    "duration_ms": round(duration * 1000, 2),
                    "success": success,
                    "type": "connection"
                }
            }
        )
    
    def log_resource_usage(self, memory_mb: float, cpu_percent: float = None):
        """Log resource usage metrics."""
        metrics = {
            "memory_mb": round(memory_mb, 2),
            "type": "resource_usage"
        }
        if cpu_percent is not None:
            metrics["cpu_percent"] = round(cpu_percent, 2)
            
        self.logger.info(
            "Resource usage metrics",
            extra={"performance": metrics}
        )


class Logger:
    """Enhanced logger with structured output and performance tracking."""
    
    def __init__(self, name: str, logger: logging.Logger):
        self.name = name
        self._logger = logger
        self.performance = PerformanceLogger(logger)
    
    def debug(self, message: str, **kwargs):
        """Log debug message with optional extra fields."""
        self._logger.debug(message, extra=kwargs)
    
    def info(self, message: str, **kwargs):
        """Log info message with optional extra fields."""
        self._logger.info(message, extra=kwargs)
    
    def warning(self, message: str, **kwargs):
        """Log warning message with optional extra fields."""
        self._logger.warning(message, extra=kwargs)
    
    def error(self, message: str, exc_info: bool = False, **kwargs):
        """Log error message with optional exception info and extra fields."""
        self._logger.error(message, exc_info=exc_info, extra=kwargs)
    
    def critical(self, message: str, exc_info: bool = False, **kwargs):
        """Log critical message with optional exception info and extra fields."""
        self._logger.critical(message, exc_info=exc_info, extra=kwargs)
    
    def log_mcp_request(self, request_id: str, method: str, params: Dict[str, Any]):
        """Log MCP request with structured data."""
        self.info(
            f"MCP request received: {method}",
            mcp={
                "request_id": request_id,
                "method": method,
                "params": params,
                "type": "request"
            }
        )
    
    def log_mcp_response(self, request_id: str, method: str, success: bool, duration: float):
        """Log MCP response with timing information."""
        level = "info" if success else "error"
        message = f"MCP response sent: {method} ({'success' if success else 'error'})"
        
        mcp_data = {
            "request_id": request_id,
            "method": method,
            "success": success,
            "duration_ms": round(duration * 1000, 2),
            "type": "response"
        }
        
        if level == "info":
            self.info(message, mcp=mcp_data)
        else:
            self.error(message, mcp=mcp_data)
    
    def log_outlook_operation(self, operation: str, success: bool, duration: float = None, **kwargs):
        """Log Outlook COM operation with timing and context."""
        message = f"Outlook operation: {operation} ({'success' if success else 'failed'})"
        
        outlook_data = {
            "operation": operation,
            "success": success,
            "type": "com_operation",
            **kwargs
        }
        
        if duration is not None:
            outlook_data["duration_ms"] = round(duration * 1000, 2)
        
        if success:
            self.info(message, outlook=outlook_data)
        else:
            self.error(message, outlook=outlook_data)
    
    def log_connection_status(self, connected: bool, details: str = None):
        """Log Outlook connection status (Requirement 5.5)."""
        message = f"Outlook connection {'established' if connected else 'failed'}"
        if details:
            message += f": {details}"
            
        outlook_data = {
            "connected": connected,
            "details": details,
            "type": "connection_status"
        }
        
        if connected:
            self.info(message, outlook=outlook_data)
        else:
            self.error(message, outlook=outlook_data)
    
    @contextmanager
    def time_operation(self, operation_name: str, log_success: bool = True):
        """Context manager to time operations and log performance."""
        start_time = time.time()
        success = True
        exception = None
        
        try:
            yield
        except Exception as e:
            success = False
            exception = e
            raise
        finally:
            duration = time.time() - start_time
            
            if log_success or not success:
                if success:
                    self.performance.log_request_timing(operation_name, duration, success=True)
                else:
                    self.performance.log_request_timing(
                        operation_name, 
                        duration, 
                        success=False,
                        error=str(exception) if exception else "Unknown error"
                    )


def configure_logging(
    log_level: str = "INFO",
    log_dir: str = "logs",
    max_bytes: int = 10 * 1024 * 1024,  # 10MB
    backup_count: int = 5,
    console_output: bool = True
) -> None:
    """
    Configure the logging system with structured JSON output and rotation.
    
    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_dir: Directory for log files
        max_bytes: Maximum size of each log file before rotation
        backup_count: Number of backup files to keep
        console_output: Whether to also output to console
    """
    # Create log directory if it doesn't exist
    log_path = Path(log_dir)
    log_path.mkdir(exist_ok=True)
    
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, log_level.upper()))
    
    # Clear existing handlers
    root_logger.handlers.clear()
    
    # Create JSON formatter
    json_formatter = JSONFormatter()
    
    # File handler with rotation
    file_handler = logging.handlers.RotatingFileHandler(
        filename=log_path / "outlook_mcp_server.log",
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    file_handler.setFormatter(json_formatter)
    root_logger.addHandler(file_handler)
    
    # Console handler (optional)
    if console_output:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(json_formatter)
        root_logger.addHandler(console_handler)
    
    # Configure specific loggers to prevent duplicate messages
    logging.getLogger("outlook_mcp_server").propagate = True


def get_logger(name: str) -> Logger:
    """
    Get a logger instance with the specified name.
    
    Args:
        name: Logger name (typically __name__)
        
    Returns:
        Logger instance with structured output capabilities
    """
    python_logger = logging.getLogger(name)
    return Logger(name, python_logger)