"""Unit tests for the ErrorHandler class."""

import pytest
import logging
import time
from unittest.mock import Mock, patch, MagicMock
from typing import Dict, Any

from src.outlook_mcp_server.error_handler import (
    ErrorHandler,
    ErrorContext,
    ErrorSeverity,
    outlook_connection_retry_strategy,
    timeout_retry_strategy,
    create_exponential_backoff_strategy
)
from src.outlook_mcp_server.models.exceptions import (
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


class TestErrorHandler:
    """Test cases for ErrorHandler class."""
    
    @pytest.fixture
    def mock_logger(self):
        """Create a mock logger for testing."""
        return Mock(spec=logging.Logger)
    
    @pytest.fixture
    def error_handler(self, mock_logger):
        """Create an ErrorHandler instance for testing."""
        return ErrorHandler(logger=mock_logger)
    
    @pytest.fixture
    def sample_context(self):
        """Create a sample error context for testing."""
        return ErrorContext(
            request_id="test-123",
            method="list_emails",
            parameters={"folder": "Inbox", "limit": 10},
            timestamp=time.time(),
            user_agent="TestClient/1.0",
            client_info={"version": "1.0", "platform": "test"}
        )
    
    def test_error_handler_initialization(self):
        """Test ErrorHandler initialization."""
        handler = ErrorHandler()
        assert handler.logger is not None
        assert isinstance(handler._error_categories, dict)
        assert isinstance(handler._retry_strategies, dict)
    
    def test_error_handler_with_custom_logger(self, mock_logger):
        """Test ErrorHandler initialization with custom logger."""
        handler = ErrorHandler(logger=mock_logger)
        assert handler.logger is mock_logger
    
    def test_error_severity_categorization(self, error_handler):
        """Test error severity categorization."""
        # Test low severity errors
        validation_error = ValidationError("Invalid input")
        assert error_handler._get_error_severity(validation_error) == ErrorSeverity.LOW
        
        invalid_param_error = InvalidParameterError("param1")
        assert error_handler._get_error_severity(invalid_param_error) == ErrorSeverity.LOW
        
        # Test medium severity errors
        email_not_found = EmailNotFoundError("email123")
        assert error_handler._get_error_severity(email_not_found) == ErrorSeverity.MEDIUM
        
        folder_not_found = FolderNotFoundError("NonExistent")
        assert error_handler._get_error_severity(folder_not_found) == ErrorSeverity.MEDIUM
        
        # Test high severity errors
        permission_error = PermissionError("restricted_folder")
        assert error_handler._get_error_severity(permission_error) == ErrorSeverity.HIGH
        
        timeout_error = TimeoutError("search_operation", 30)
        assert error_handler._get_error_severity(timeout_error) == ErrorSeverity.HIGH
        
        # Test critical severity errors
        connection_error = OutlookConnectionError()
        assert error_handler._get_error_severity(connection_error) == ErrorSeverity.CRITICAL
        
        # Test unknown error type (should default to medium)
        generic_error = Exception("Unknown error")
        assert error_handler._get_error_severity(generic_error) == ErrorSeverity.MEDIUM
    
    def test_register_retry_strategy(self, error_handler):
        """Test registering retry strategies."""
        def custom_strategy(error, context):
            return {"recovered": True}
        
        error_handler.register_retry_strategy(OutlookConnectionError, custom_strategy)
        assert OutlookConnectionError in error_handler._retry_strategies
        assert error_handler._retry_strategies[OutlookConnectionError] == custom_strategy
    
    def test_create_context(self, error_handler):
        """Test creating error context."""
        context = error_handler.create_context(
            request_id="test-456",
            method="get_email",
            parameters={"email_id": "123"},
            user_agent="TestAgent/2.0",
            client_info={"version": "2.0"}
        )
        
        assert context.request_id == "test-456"
        assert context.method == "get_email"
        assert context.parameters == {"email_id": "123"}
        assert context.user_agent == "TestAgent/2.0"
        assert context.client_info == {"version": "2.0"}
        assert isinstance(context.timestamp, float)
    
    def test_log_error_low_severity(self, error_handler, mock_logger, sample_context):
        """Test logging low severity errors."""
        error = ValidationError("Invalid email format")
        error_handler._log_error(error, sample_context)
        
        mock_logger.info.assert_called_once()
        call_args = mock_logger.info.call_args
        assert "Low severity error occurred" in call_args[0][0]
        
        extra_data = call_args[1]["extra"]
        assert extra_data["error_type"] == "ValidationError"
        assert extra_data["severity"] == "low"
        assert extra_data["request_id"] == "test-123"
    
    def test_log_error_critical_severity(self, error_handler, mock_logger, sample_context):
        """Test logging critical severity errors."""
        error = OutlookConnectionError("Connection failed")
        error_handler._log_error(error, sample_context)
        
        mock_logger.critical.assert_called_once()
        call_args = mock_logger.critical.call_args
        assert "Critical error occurred" in call_args[0][0]
        
        extra_data = call_args[1]["extra"]
        assert extra_data["error_type"] == "OutlookConnectionError"
        assert extra_data["severity"] == "critical"
        assert extra_data["traceback"] is not None
    
    def test_generate_error_response_outlook_mcp_error(self, error_handler, sample_context):
        """Test generating error response for OutlookMCPError."""
        error = EmailNotFoundError("email123")
        response = error_handler._generate_error_response(error, sample_context)
        
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test-123"
        assert "error" in response
        
        error_data = response["error"]
        assert error_data["code"] == -32003
        assert "Email with ID 'email123' not found" in error_data["message"]
        assert error_data["data"]["type"] == "EmailNotFoundError"
        assert error_data["data"]["severity"] == "medium"
        assert error_data["data"]["context"]["request_id"] == "test-123"
    
    def test_generate_error_response_generic_error(self, error_handler, sample_context):
        """Test generating error response for generic exceptions."""
        error = ValueError("Invalid value provided")
        response = error_handler._generate_error_response(error, sample_context)
        
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test-123"
        assert "error" in response
        
        error_data = response["error"]
        assert error_data["code"] == -32000
        assert error_data["message"] == "Invalid value provided"
        assert error_data["data"]["type"] == "ValueError"
        assert error_data["data"]["severity"] == "medium"
    
    def test_handle_error_without_recovery(self, error_handler, mock_logger, sample_context):
        """Test handling error without recovery."""
        error = ValidationError("Invalid input")
        
        with patch.object(error_handler, '_attempt_recovery', return_value=None):
            response = error_handler.handle_error(error, sample_context)
        
        # Should log the error
        mock_logger.info.assert_called_once()
        
        # Should return structured error response
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test-123"
        assert "error" in response
    
    def test_handle_error_with_successful_recovery(self, error_handler, mock_logger, sample_context):
        """Test handling error with successful recovery."""
        error = OutlookConnectionError("Connection lost")
        recovery_result = {"jsonrpc": "2.0", "id": "test-123", "result": "recovered"}
        
        with patch.object(error_handler, '_attempt_recovery', return_value=recovery_result):
            response = error_handler.handle_error(error, sample_context)
        
        # Should log the error
        mock_logger.critical.assert_called_once()
        
        # Should return recovery result
        assert response == recovery_result
    
    def test_attempt_recovery_with_registered_strategy(self, error_handler, sample_context):
        """Test recovery attempt with registered strategy."""
        def mock_strategy(error, context):
            return {"recovered": True, "method": "custom_strategy"}
        
        error_handler.register_retry_strategy(OutlookConnectionError, mock_strategy)
        error = OutlookConnectionError("Connection failed")
        
        result = error_handler._attempt_recovery(error, sample_context)
        assert result == {"recovered": True, "method": "custom_strategy"}
    
    def test_attempt_recovery_strategy_failure(self, error_handler, mock_logger, sample_context):
        """Test recovery when strategy itself fails."""
        def failing_strategy(error, context):
            raise Exception("Strategy failed")
        
        error_handler.register_retry_strategy(ValidationError, failing_strategy)
        error = ValidationError("Invalid input")
        
        result = error_handler._attempt_recovery(error, sample_context)
        assert result is None
        mock_logger.warning.assert_called_once()
    
    def test_recover_outlook_connection(self, error_handler, mock_logger, sample_context):
        """Test Outlook connection recovery attempt."""
        error = OutlookConnectionError("Connection lost")
        result = error_handler._recover_outlook_connection(error, sample_context)
        
        # Current implementation returns None (no recovery)
        assert result is None
        mock_logger.info.assert_called_once()
        
        call_args = mock_logger.info.call_args
        assert "Attempting Outlook connection recovery" in call_args[0][0]
    
    def test_recover_timeout(self, error_handler, mock_logger, sample_context):
        """Test timeout recovery attempt."""
        error = TimeoutError("search_operation", 30)
        result = error_handler._recover_timeout(error, sample_context)
        
        # Current implementation returns None (no recovery)
        assert result is None
        mock_logger.info.assert_called_once()
        
        call_args = mock_logger.info.call_args
        assert "Timeout recovery not implemented" in call_args[0][0]
    
    def test_get_error_statistics(self, error_handler):
        """Test getting error statistics."""
        stats = error_handler.get_error_statistics()
        
        assert isinstance(stats, dict)
        assert "total_errors" in stats
        assert "errors_by_type" in stats
        assert "errors_by_severity" in stats
        assert "recovery_attempts" in stats
        assert "successful_recoveries" in stats
        
        # Check severity breakdown
        severity_stats = stats["errors_by_severity"]
        assert "low" in severity_stats
        assert "medium" in severity_stats
        assert "high" in severity_stats
        assert "critical" in severity_stats
    
    def test_error_statistics_tracking(self, error_handler, sample_context):
        """Test that error statistics are properly tracked."""
        # Initial stats should be zero
        initial_stats = error_handler.get_error_statistics()
        assert initial_stats["total_errors"] == 0
        
        # Handle a validation error
        validation_error = ValidationError("Invalid input")
        error_handler.handle_error(validation_error, sample_context)
        
        # Check updated stats
        stats = error_handler.get_error_statistics()
        assert stats["total_errors"] == 1
        assert stats["errors_by_type"]["ValidationError"] == 1
        assert stats["errors_by_severity"]["low"] == 1
        assert stats["recovery_attempts"] == 1  # Attempt is made even if no strategy exists
        
        # Handle another error of different type
        connection_error = OutlookConnectionError("Connection failed")
        error_handler.handle_error(connection_error, sample_context)
        
        # Check updated stats
        stats = error_handler.get_error_statistics()
        assert stats["total_errors"] == 2
        assert stats["errors_by_type"]["OutlookConnectionError"] == 1
        assert stats["errors_by_severity"]["critical"] == 1
        assert stats["recovery_attempts"] == 2
    
    def test_reset_error_statistics(self, error_handler, sample_context):
        """Test resetting error statistics."""
        # Generate some errors first
        error_handler.handle_error(ValidationError("Test"), sample_context)
        error_handler.handle_error(EmailNotFoundError("123"), sample_context)
        
        # Verify stats are not zero
        stats = error_handler.get_error_statistics()
        assert stats["total_errors"] > 0
        
        # Reset stats
        error_handler.reset_error_statistics()
        
        # Verify stats are reset
        stats = error_handler.get_error_statistics()
        assert stats["total_errors"] == 0
        assert stats["errors_by_type"] == {}
        assert all(count == 0 for count in stats["errors_by_severity"].values())
        assert stats["recovery_attempts"] == 0
        assert stats["successful_recoveries"] == 0
    
    def test_configure_structured_logging(self):
        """Test configuring structured logging."""
        # Create error handler without mock logger for this test
        error_handler = ErrorHandler()
        
        # Test basic configuration
        error_handler.configure_structured_logging(use_json=False)
        assert len(error_handler.logger.handlers) > 0
        
        # Test JSON configuration
        error_handler.configure_structured_logging(use_json=True)
        assert len(error_handler.logger.handlers) > 0
        
        # Verify logger is properly configured
        assert error_handler.logger.level == logging.INFO


class TestErrorContext:
    """Test cases for ErrorContext class."""
    
    def test_error_context_creation(self):
        """Test ErrorContext creation."""
        timestamp = time.time()
        context = ErrorContext(
            request_id="test-789",
            method="search_emails",
            parameters={"query": "test", "limit": 5},
            timestamp=timestamp,
            user_agent="TestClient/3.0",
            client_info={"platform": "windows"}
        )
        
        assert context.request_id == "test-789"
        assert context.method == "search_emails"
        assert context.parameters == {"query": "test", "limit": 5}
        assert context.timestamp == timestamp
        assert context.user_agent == "TestClient/3.0"
        assert context.client_info == {"platform": "windows"}
    
    def test_error_context_optional_fields(self):
        """Test ErrorContext with optional fields."""
        context = ErrorContext(
            request_id="test-456",
            method="get_folders",
            parameters={},
            timestamp=time.time()
        )
        
        assert context.user_agent is None
        assert context.client_info is None


class TestRetryStrategies:
    """Test cases for retry strategies."""
    
    def test_outlook_connection_retry_strategy(self):
        """Test Outlook connection retry strategy."""
        error = OutlookConnectionError("Connection failed")
        context = ErrorContext(
            request_id="test-retry",
            method="list_emails",
            parameters={},
            timestamp=time.time()
        )
        
        # Current implementation returns None
        result = outlook_connection_retry_strategy(error, context)
        assert result is None
    
    def test_timeout_retry_strategy(self):
        """Test timeout retry strategy."""
        error = TimeoutError("search_operation", 30)
        context = ErrorContext(
            request_id="test-timeout",
            method="search_emails",
            parameters={"query": "test"},
            timestamp=time.time()
        )
        
        # Current implementation returns None
        result = timeout_retry_strategy(error, context)
        assert result is None
    
    def test_outlook_connection_retry_strategy_with_retries(self):
        """Test Outlook connection retry strategy with retry count."""
        error = OutlookConnectionError("Connection failed")
        
        # Test with no previous retries
        context = ErrorContext(
            request_id="test-retry",
            method="list_emails",
            parameters={},
            timestamp=time.time()
        )
        result = outlook_connection_retry_strategy(error, context)
        assert result is None
        
        # Test with max retries exceeded
        context_max_retries = ErrorContext(
            request_id="test-retry-max",
            method="list_emails",
            parameters={"_retry_count": 5},
            timestamp=time.time()
        )
        result = outlook_connection_retry_strategy(error, context_max_retries, max_retries=3)
        assert result is None
    
    def test_timeout_retry_strategy_with_retries(self):
        """Test timeout retry strategy with retry count."""
        error = TimeoutError("search_operation", 30)
        
        # Test with no previous retries
        context = ErrorContext(
            request_id="test-timeout",
            method="search_emails",
            parameters={"query": "test"},
            timestamp=time.time()
        )
        result = timeout_retry_strategy(error, context)
        assert result is None
        
        # Test with max retries exceeded
        context_max_retries = ErrorContext(
            request_id="test-timeout-max",
            method="search_emails",
            parameters={"query": "test", "_retry_count": 3},
            timestamp=time.time()
        )
        result = timeout_retry_strategy(error, context_max_retries, max_retries=2)
        assert result is None
    
    def test_create_exponential_backoff_strategy(self):
        """Test creating exponential backoff strategy."""
        strategy = create_exponential_backoff_strategy(max_retries=2, base_delay=0.5, max_delay=10.0)
        
        error = ValidationError("Test error")
        context = ErrorContext(
            request_id="test-backoff",
            method="test_method",
            parameters={},
            timestamp=time.time()
        )
        
        # Test strategy execution
        result = strategy(error, context)
        assert result is None  # Current implementation returns None
        
        # Test with max retries exceeded
        context_max = ErrorContext(
            request_id="test-backoff-max",
            method="test_method",
            parameters={"_retry_count": 5},
            timestamp=time.time()
        )
        result = strategy(error, context_max)
        assert result is None


class TestErrorSeverity:
    """Test cases for ErrorSeverity enum."""
    
    def test_error_severity_values(self):
        """Test ErrorSeverity enum values."""
        assert ErrorSeverity.LOW.value == "low"
        assert ErrorSeverity.MEDIUM.value == "medium"
        assert ErrorSeverity.HIGH.value == "high"
        assert ErrorSeverity.CRITICAL.value == "critical"
    
    def test_error_severity_comparison(self):
        """Test ErrorSeverity enum comparison."""
        assert ErrorSeverity.LOW != ErrorSeverity.MEDIUM
        assert ErrorSeverity.HIGH != ErrorSeverity.CRITICAL
        assert ErrorSeverity.LOW == ErrorSeverity.LOW


if __name__ == "__main__":
    pytest.main([__file__])