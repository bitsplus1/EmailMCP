"""
Logging configuration module for Outlook MCP Server.

This module provides configuration classes and utilities for the logging system.
"""

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass
class LoggingConfig:
    """Configuration class for logging system."""
    
    # Log levels
    level: str = "INFO"
    
    # File logging
    log_dir: str = "logs"
    max_file_size_mb: int = 10
    backup_count: int = 5
    
    # Console output
    console_output: bool = True
    
    # Performance logging
    enable_performance_logging: bool = True
    log_request_timing: bool = True
    log_resource_usage: bool = True
    
    # Structured logging
    json_format: bool = True
    include_thread_info: bool = True
    include_process_info: bool = True
    
    @property
    def max_bytes(self) -> int:
        """Convert max file size from MB to bytes."""
        return self.max_file_size_mb * 1024 * 1024
    
    @classmethod
    def from_environment(cls) -> 'LoggingConfig':
        """Create configuration from environment variables."""
        return cls(
            level=os.getenv("LOG_LEVEL", "INFO").upper(),
            log_dir=os.getenv("LOG_DIR", "logs"),
            max_file_size_mb=int(os.getenv("LOG_MAX_FILE_SIZE_MB", "10")),
            backup_count=int(os.getenv("LOG_BACKUP_COUNT", "5")),
            console_output=os.getenv("LOG_CONSOLE_OUTPUT", "true").lower() == "true",
            enable_performance_logging=os.getenv("LOG_PERFORMANCE", "true").lower() == "true",
            log_request_timing=os.getenv("LOG_REQUEST_TIMING", "true").lower() == "true",
            log_resource_usage=os.getenv("LOG_RESOURCE_USAGE", "true").lower() == "true",
            json_format=os.getenv("LOG_JSON_FORMAT", "true").lower() == "true",
            include_thread_info=os.getenv("LOG_THREAD_INFO", "true").lower() == "true",
            include_process_info=os.getenv("LOG_PROCESS_INFO", "true").lower() == "true",
        )
    
    def validate(self) -> None:
        """Validate configuration values."""
        valid_levels = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}
        if self.level not in valid_levels:
            raise ValueError(f"Invalid log level: {self.level}. Must be one of {valid_levels}")
        
        if self.max_file_size_mb <= 0:
            raise ValueError("Max file size must be positive")
        
        if self.backup_count < 0:
            raise ValueError("Backup count must be non-negative")
        
        # Ensure log directory is valid
        try:
            Path(self.log_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            raise ValueError(f"Cannot create log directory '{self.log_dir}': {e}")


# Default configuration instance
DEFAULT_CONFIG = LoggingConfig()