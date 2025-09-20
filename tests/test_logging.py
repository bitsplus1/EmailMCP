"""
Tests for the logging system.

This module tests the comprehensive logging functionality including:
- Structured JSON output
- Log rotation
- Performance metrics
- Different log levels
- Requirements: 5.2, 5.5, 8.4
"""

import json
import logging
import os
import tempfile
import time
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from src.outlook_mcp_server.logging import Logger, get_logger, configure_logging
from src.outlook_mcp_server.logging.logger import JSONFormatter, PerformanceLogger
from src.outlook_mcp_server.logging.config import LoggingConfig


class TestJSONFormatter:
    """Test the JSON formatter for structured logging."""
    
    def test_basic_formatting(self):
        """Test basic log record formatting to JSON."""
        formatter = JSONFormatter()
        record = logging.LogRecord(
            name="test_logger",
            level=logging.INFO,
            pathname="/test/path.py",
            lineno=42,
            msg="Test message",
            args=(),
            exc_info=None
        )
        
        result = formatter.format(record)
        log_data = json.loads(result)
        
        assert log_data["level"] == "INFO"
        assert log_data["logger"] == "test_logger"
        assert log_data["message"] == "Test message"
        assert log_data["line"] == 42
        assert "timestamp" in log_data
        assert log_data["timestamp"].endswith("Z")
    
    def test_exception_formatting(self):
        """Test formatting with exception information."""
        formatter = JSONFormatter()
        
        try:
            raise ValueError("Test exception")
        except ValueError:
            import sys
            record = logging.LogRecord(
                name="test_logger",
                level=logging.ERROR,
                pathname="/test/path.py",
                lineno=42,
                msg="Error occurred",
                args=(),
                exc_info=sys.exc_info()
            )
        
        result = formatter.format(record)
        log_data = json.loads(result)
        
        assert "exception" in log_data
        assert log_data["exception"]["type"] == "ValueError"
        assert log_data["exception"]["message"] == "Test exception"
        assert "traceback" in log_data["exception"]
    
    def test_extra_fields(self):
        """Test formatting with extra fields."""
        formatter = JSONFormatter()
        record = logging.LogRecord(
            name="test_logger",
            level=logging.INFO,
            pathname="/test/path.py",
            lineno=42,
            msg="Test message",
            args=(),
            exc_info=None
        )
        
        # Add extra fields
        record.request_id = "test-123"
        record.user_id = "user-456"
        
        result = formatter.format(record)
        log_data = json.loads(result)
        
        assert "extra" in log_data
        assert log_data["extra"]["request_id"] == "test-123"
        assert log_data["extra"]["user_id"] == "user-456"


class TestPerformanceLogger:
    """Test the performance logging functionality."""
    
    def test_request_timing(self):
        """Test logging request timing information."""
        mock_logger = MagicMock()
        perf_logger = PerformanceLogger(mock_logger)
        
        perf_logger.log_request_timing("test_method", 0.123, success=True, param1="value1")
        
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args
        
        assert "Request completed: test_method" in call_args[0][0]
        extra_data = call_args[1]["extra"]
        assert extra_data["performance"]["method"] == "test_method"
        assert extra_data["performance"]["duration_ms"] == 123.0
        assert extra_data["performance"]["success"] is True
        assert extra_data["performance"]["param1"] == "value1"
    
    def test_connection_timing(self):
        """Test logging connection operation timing."""
        mock_logger = MagicMock()
        perf_logger = PerformanceLogger(mock_logger)
        
        perf_logger.log_connection_timing("connect", 0.456, success=False)
        
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args
        
        extra_data = call_args[1]["extra"]
        assert extra_data["performance"]["operation"] == "connect"
        assert extra_data["performance"]["duration_ms"] == 456.0
        assert extra_data["performance"]["success"] is False
        assert extra_data["performance"]["type"] == "connection"
    
    def test_resource_usage(self):
        """Test logging resource usage metrics."""
        mock_logger = MagicMock()
        perf_logger = PerformanceLogger(mock_logger)
        
        perf_logger.log_resource_usage(128.5, cpu_percent=45.2)
        
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args
        
        extra_data = call_args[1]["extra"]
        assert extra_data["performance"]["memory_mb"] == 128.5
        assert extra_data["performance"]["cpu_percent"] == 45.2
        assert extra_data["performance"]["type"] == "resource_usage"


class TestLogger:
    """Test the enhanced Logger class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.mock_python_logger = MagicMock()
        self.logger = Logger("test_logger", self.mock_python_logger)
    
    def test_basic_logging_methods(self):
        """Test basic logging methods."""
        self.logger.debug("Debug message", extra_field="value")
        self.logger.info("Info message")
        self.logger.warning("Warning message")
        self.logger.error("Error message", exc_info=True)
        self.logger.critical("Critical message")
        
        assert self.mock_python_logger.debug.called
        assert self.mock_python_logger.info.called
        assert self.mock_python_logger.warning.called
        assert self.mock_python_logger.error.called
        assert self.mock_python_logger.critical.called
    
    def test_mcp_request_logging(self):
        """Test MCP request logging."""
        params = {"folder": "inbox", "limit": 10}
        self.logger.log_mcp_request("req-123", "list_emails", params)
        
        self.mock_python_logger.info.assert_called_once()
        call_args = self.mock_python_logger.info.call_args
        
        assert "MCP request received: list_emails" in call_args[0][0]
        # The extra data is passed as keyword arguments
        assert "extra" in call_args[1]
        extra_data = call_args[1]["extra"]
        assert extra_data["mcp"]["request_id"] == "req-123"
        assert extra_data["mcp"]["method"] == "list_emails"
        assert extra_data["mcp"]["params"] == params
        assert extra_data["mcp"]["type"] == "request"
    
    def test_mcp_response_logging_success(self):
        """Test MCP response logging for successful requests."""
        self.logger.log_mcp_response("req-123", "list_emails", success=True, duration=0.234)
        
        self.mock_python_logger.info.assert_called_once()
        call_args = self.mock_python_logger.info.call_args
        
        extra_data = call_args[1]["extra"]
        assert extra_data["mcp"]["success"] is True
        assert extra_data["mcp"]["duration_ms"] == 234.0
    
    def test_mcp_response_logging_error(self):
        """Test MCP response logging for failed requests."""
        self.logger.log_mcp_response("req-123", "list_emails", success=False, duration=0.1)
        
        self.mock_python_logger.error.assert_called_once()
    
    def test_outlook_operation_logging(self):
        """Test Outlook operation logging."""
        self.logger.log_outlook_operation(
            "get_folders", 
            success=True, 
            duration=0.5, 
            folder_count=5
        )
        
        self.mock_python_logger.info.assert_called_once()
        call_args = self.mock_python_logger.info.call_args
        
        extra_data = call_args[1]["extra"]
        assert extra_data["outlook"]["operation"] == "get_folders"
        assert extra_data["outlook"]["success"] is True
        assert extra_data["outlook"]["duration_ms"] == 500.0
        assert extra_data["outlook"]["folder_count"] == 5
    
    def test_connection_status_logging(self):
        """Test connection status logging (Requirement 5.5)."""
        # Test successful connection
        self.logger.log_connection_status(True, "Connected to Outlook")
        self.mock_python_logger.info.assert_called_once()
        
        # Test failed connection
        self.mock_python_logger.reset_mock()
        self.logger.log_connection_status(False, "COM connection failed")
        self.mock_python_logger.error.assert_called_once()
    
    def test_time_operation_context_manager_success(self):
        """Test timing context manager for successful operations."""
        with patch.object(self.logger.performance, 'log_request_timing') as mock_log:
            with self.logger.time_operation("test_operation"):
                time.sleep(0.01)  # Small delay to ensure measurable time
        
        mock_log.assert_called_once()
        # Check positional arguments
        call_args = mock_log.call_args[0]
        assert call_args[0] == "test_operation"
        assert call_args[1] > 0  # Duration should be positive
        # Check keyword arguments
        call_kwargs = mock_log.call_args[1]
        assert call_kwargs["success"] is True
    
    def test_time_operation_context_manager_exception(self):
        """Test timing context manager when exception occurs."""
        with patch.object(self.logger.performance, 'log_request_timing') as mock_log:
            with pytest.raises(ValueError):
                with self.logger.time_operation("test_operation"):
                    raise ValueError("Test error")
        
        mock_log.assert_called_once()
        call_kwargs = mock_log.call_args[1]
        assert call_kwargs["success"] is False
        assert "error" in call_kwargs


class TestLoggingConfiguration:
    """Test logging configuration and setup."""
    
    def test_configure_logging_basic(self):
        """Test basic logging configuration."""
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                configure_logging(
                    log_level="DEBUG",
                    log_dir=temp_dir,
                    console_output=False
                )
                
                # Check that log file is created
                log_file = Path(temp_dir) / "outlook_mcp_server.log"
                
                # Test logging
                logger = get_logger("test")
                logger.info("Test message")
                
                # Force flush and close handlers
                for handler in logging.getLogger().handlers:
                    handler.flush()
                
                # Verify log file exists and has content
                assert log_file.exists()
            finally:
                # Clean up handlers to release file locks
                root_logger = logging.getLogger()
                for handler in root_logger.handlers[:]:
                    handler.close()
                    root_logger.removeHandler(handler)
    
    def test_get_logger(self):
        """Test getting logger instances."""
        logger1 = get_logger("test.module1")
        logger2 = get_logger("test.module2")
        
        assert isinstance(logger1, Logger)
        assert isinstance(logger2, Logger)
        assert logger1.name == "test.module1"
        assert logger2.name == "test.module2"
    
    def test_log_rotation(self):
        """Test log file rotation functionality."""
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                # Configure with very small file size for testing
                configure_logging(
                    log_dir=temp_dir,
                    max_bytes=100,  # Very small for testing
                    backup_count=2,
                    console_output=False
                )
                
                logger = get_logger("test")
                
                # Generate enough log entries to trigger rotation
                for i in range(50):
                    logger.info(f"Test message {i} with some additional content to make it longer")
                
                # Force flush handlers
                for handler in logging.getLogger().handlers:
                    handler.flush()
                
                # Check that rotation occurred
                log_files = list(Path(temp_dir).glob("outlook_mcp_server.log*"))
                assert len(log_files) > 1  # Should have main log + rotated files
            finally:
                # Clean up handlers to release file locks
                root_logger = logging.getLogger()
                for handler in root_logger.handlers[:]:
                    handler.close()
                    root_logger.removeHandler(handler)


class TestLoggingConfig:
    """Test the LoggingConfig class."""
    
    def test_default_config(self):
        """Test default configuration values."""
        config = LoggingConfig()
        
        assert config.level == "INFO"
        assert config.log_dir == "logs"
        assert config.max_file_size_mb == 10
        assert config.backup_count == 5
        assert config.console_output is True
        assert config.max_bytes == 10 * 1024 * 1024
    
    def test_from_environment(self):
        """Test configuration from environment variables."""
        env_vars = {
            "LOG_LEVEL": "DEBUG",
            "LOG_DIR": "/custom/logs",
            "LOG_MAX_FILE_SIZE_MB": "20",
            "LOG_BACKUP_COUNT": "10",
            "LOG_CONSOLE_OUTPUT": "false",
            "LOG_PERFORMANCE": "false"
        }
        
        with patch.dict(os.environ, env_vars):
            config = LoggingConfig.from_environment()
        
        assert config.level == "DEBUG"
        assert config.log_dir == "/custom/logs"
        assert config.max_file_size_mb == 20
        assert config.backup_count == 10
        assert config.console_output is False
        assert config.enable_performance_logging is False
    
    def test_config_validation(self):
        """Test configuration validation."""
        # Test invalid log level
        config = LoggingConfig(level="INVALID")
        with pytest.raises(ValueError, match="Invalid log level"):
            config.validate()
        
        # Test invalid file size
        config = LoggingConfig(max_file_size_mb=-1)
        with pytest.raises(ValueError, match="Max file size must be positive"):
            config.validate()
        
        # Test invalid backup count
        config = LoggingConfig(backup_count=-1)
        with pytest.raises(ValueError, match="Backup count must be non-negative"):
            config.validate()


class TestIntegration:
    """Integration tests for the complete logging system."""
    
    def test_structured_json_output(self):
        """Test that logs are properly formatted as JSON."""
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                configure_logging(
                    log_dir=temp_dir,
                    console_output=False
                )
                
                logger = get_logger("integration_test")
                
                # Log various types of messages
                logger.info("Simple info message")
                logger.log_mcp_request("req-123", "test_method", {"param": "value"})
                logger.log_outlook_operation("connect", True, 0.5)
                
                # Force flush handlers
                for handler in logging.getLogger().handlers:
                    handler.flush()
                
                # Read and verify log file content
                log_file = Path(temp_dir) / "outlook_mcp_server.log"
                assert log_file.exists()
                
                with open(log_file, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                # Verify each line is valid JSON
                for line in lines:
                    if line.strip():
                        log_data = json.loads(line)
                        assert "timestamp" in log_data
                        assert "level" in log_data
                        assert "message" in log_data
            finally:
                # Clean up handlers to release file locks
                root_logger = logging.getLogger()
                for handler in root_logger.handlers[:]:
                    handler.close()
                    root_logger.removeHandler(handler)
    
    def test_performance_logging_integration(self):
        """Test performance logging integration."""
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                configure_logging(log_dir=temp_dir, console_output=False)
                
                logger = get_logger("perf_test")
                
                # Test timing context manager
                with logger.time_operation("test_operation"):
                    time.sleep(0.01)
                
                # Test direct performance logging
                logger.performance.log_request_timing("manual_test", 0.123)
                
                # Force flush handlers
                for handler in logging.getLogger().handlers:
                    handler.flush()
                
                # Verify logs contain performance data
                log_file = Path(temp_dir) / "outlook_mcp_server.log"
                with open(log_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                assert "performance" in content
                assert "duration_ms" in content
            finally:
                # Clean up handlers to release file locks
                root_logger = logging.getLogger()
                for handler in root_logger.handlers[:]:
                    handler.close()
                    root_logger.removeHandler(handler)