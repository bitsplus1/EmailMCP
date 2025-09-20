"""Unit tests for RequestRouter class."""

import pytest
from unittest.mock import Mock, MagicMock
from src.outlook_mcp_server.routing.request_router import RequestRouter
from src.outlook_mcp_server.models.mcp_models import MCPRequest
from src.outlook_mcp_server.models.exceptions import (
    ValidationError, 
    MethodNotFoundError
)


class TestRequestRouter:
    """Test cases for RequestRouter class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.router = RequestRouter()
    
    def test_init(self):
        """Test router initialization."""
        assert isinstance(self.router._handlers, dict)
        assert isinstance(self.router._parameter_schemas, dict)
        assert len(self.router._handlers) == 0
        
        # Check that parameter schemas are set up
        expected_methods = ["list_emails", "get_email", "search_emails", "get_folders"]
        for method in expected_methods:
            assert method in self.router._parameter_schemas
    
    def test_register_handler_success(self):
        """Test successful handler registration."""
        mock_handler = Mock()
        
        self.router.register_handler("test_method", mock_handler)
        
        assert "test_method" in self.router._handlers
        assert self.router._handlers["test_method"] == mock_handler
    
    def test_register_handler_invalid_method(self):
        """Test handler registration with invalid method name."""
        mock_handler = Mock()
        
        with pytest.raises(ValidationError, match="Method name must be a non-empty string"):
            self.router.register_handler("", mock_handler)
        
        with pytest.raises(ValidationError, match="Method name must be a non-empty string"):
            self.router.register_handler(None, mock_handler)
    
    def test_register_handler_invalid_handler(self):
        """Test handler registration with invalid handler."""
        with pytest.raises(ValidationError, match="Handler must be callable"):
            self.router.register_handler("test_method", "not_callable")
        
        with pytest.raises(ValidationError, match="Handler must be callable"):
            self.router.register_handler("test_method", None)
    
    def test_route_request_success(self):
        """Test successful request routing."""
        mock_handler = Mock(return_value={"result": "success"})
        self.router.register_handler("get_folders", mock_handler)
        
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="get_folders",
            params={}
        )
        
        result = self.router.route_request(request)
        
        assert result == {"result": "success"}
        mock_handler.assert_called_once_with()
    
    def test_route_request_method_not_found(self):
        """Test routing request for unregistered method."""
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="unknown_method",
            params={}
        )
        
        with pytest.raises(MethodNotFoundError, match="Method 'unknown_method' not found"):
            self.router.route_request(request)
    
    def test_route_request_with_parameters(self):
        """Test routing request with parameters."""
        mock_handler = Mock(return_value={"emails": []})
        self.router.register_handler("list_emails", mock_handler)
        
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="list_emails",
            params={"folder": "Inbox", "limit": 10}
        )
        
        result = self.router.route_request(request)
        
        assert result == {"emails": []}
        mock_handler.assert_called_once_with(folder="Inbox", unread_only=False, limit=10)
    
    def test_validate_params_list_emails_success(self):
        """Test parameter validation for list_emails method."""
        params = {"folder": "Inbox", "unread_only": True, "limit": 25}
        
        validated = self.router.validate_params("list_emails", params)
        
        assert validated == {"folder": "Inbox", "unread_only": True, "limit": 25}
    
    def test_validate_params_list_emails_defaults(self):
        """Test parameter validation with defaults for list_emails."""
        params = {}
        
        validated = self.router.validate_params("list_emails", params)
        
        assert validated == {"folder": None, "unread_only": False, "limit": 50}
    
    def test_validate_params_get_email_success(self):
        """Test parameter validation for get_email method."""
        params = {"email_id": "test-email-123"}
        
        validated = self.router.validate_params("get_email", params)
        
        assert validated == {"email_id": "test-email-123"}
    
    def test_validate_params_get_email_missing_required(self):
        """Test parameter validation for get_email with missing required parameter."""
        params = {}
        
        with pytest.raises(ValidationError, match="Required parameter 'email_id' is missing"):
            self.router.validate_params("get_email", params)
    
    def test_validate_params_search_emails_success(self):
        """Test parameter validation for search_emails method."""
        params = {"query": "test search", "folder": "Inbox", "limit": 20}
        
        validated = self.router.validate_params("search_emails", params)
        
        assert validated == {"query": "test search", "folder": "Inbox", "limit": 20}
    
    def test_validate_params_search_emails_defaults(self):
        """Test parameter validation for search_emails with defaults."""
        params = {"query": "test search"}
        
        validated = self.router.validate_params("search_emails", params)
        
        assert validated == {"query": "test search", "folder": None, "limit": 50}
    
    def test_validate_params_get_folders_success(self):
        """Test parameter validation for get_folders method."""
        params = {}
        
        validated = self.router.validate_params("get_folders", params)
        
        assert validated == {}
    
    def test_validate_params_unknown_method(self):
        """Test parameter validation for unknown method."""
        with pytest.raises(MethodNotFoundError, match="No parameter schema found for method 'unknown'"):
            self.router.validate_params("unknown", {})
    
    def test_validate_params_unknown_parameter(self):
        """Test parameter validation with unknown parameter."""
        params = {"unknown_param": "value"}
        
        with pytest.raises(ValidationError, match="Unknown parameters for method 'get_folders'"):
            self.router.validate_params("get_folders", params)
    
    def test_validate_params_wrong_type(self):
        """Test parameter validation with wrong parameter type."""
        params = {"limit": "not_an_int"}
        
        with pytest.raises(ValidationError, match="Parameter 'limit' for method 'list_emails' must be of type int"):
            self.router.validate_params("list_emails", params)
    
    def test_validate_params_limit_constraints(self):
        """Test parameter validation for limit constraints."""
        # Test minimum limit
        params = {"limit": 0}
        with pytest.raises(ValidationError, match="Parameter 'limit' for method 'list_emails' must be at least 1"):
            self.router.validate_params("list_emails", params)
        
        # Test maximum limit
        params = {"limit": 1001}
        with pytest.raises(ValidationError, match="Parameter 'limit' for method 'list_emails' must be at most 1000"):
            self.router.validate_params("list_emails", params)
    
    def test_validate_email_id_empty(self):
        """Test email ID validation with empty value."""
        params = {"email_id": ""}
        
        with pytest.raises(ValidationError, match="Email ID cannot be empty"):
            self.router.validate_params("get_email", params)
    
    def test_validate_email_id_too_long(self):
        """Test email ID validation with too long value."""
        params = {"email_id": "x" * 501}
        
        with pytest.raises(ValidationError, match="Email ID is too long"):
            self.router.validate_params("get_email", params)
    
    def test_validate_email_id_dangerous_chars(self):
        """Test email ID validation with dangerous characters."""
        dangerous_chars = ['<', '>', '"', "'", '&', '\n', '\r', '\t']
        
        for char in dangerous_chars:
            params = {"email_id": f"test{char}id"}
            with pytest.raises(ValidationError, match="Email ID contains invalid characters"):
                self.router.validate_params("get_email", params)
    
    def test_validate_folder_name_empty(self):
        """Test folder name validation with empty value."""
        params = {"folder": ""}
        
        with pytest.raises(ValidationError, match="Folder name cannot be empty"):
            self.router.validate_params("list_emails", params)
    
    def test_validate_folder_name_too_long(self):
        """Test folder name validation with too long value."""
        params = {"folder": "x" * 256}
        
        with pytest.raises(ValidationError, match="Folder name is too long"):
            self.router.validate_params("list_emails", params)
    
    def test_validate_folder_name_invalid_chars(self):
        """Test folder name validation with invalid characters."""
        invalid_chars = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
        
        for char in invalid_chars:
            params = {"folder": f"test{char}folder"}
            with pytest.raises(ValidationError, match="Folder name contains invalid characters"):
                self.router.validate_params("list_emails", params)
    
    def test_validate_search_query_empty(self):
        """Test search query validation with empty value."""
        params = {"query": ""}
        
        with pytest.raises(ValidationError, match="Parameter 'query' for method 'search_emails' must be at least 1 characters long"):
            self.router.validate_params("search_emails", params)
    
    def test_validate_search_query_too_long(self):
        """Test search query validation with too long value."""
        params = {"query": "x" * 1001}
        
        with pytest.raises(ValidationError, match="Parameter 'query' for method 'search_emails' must be at most 1000 characters long"):
            self.router.validate_params("search_emails", params)
    
    def test_get_registered_methods(self):
        """Test getting list of registered methods."""
        mock_handler1 = Mock()
        mock_handler2 = Mock()
        
        self.router.register_handler("method1", mock_handler1)
        self.router.register_handler("method2", mock_handler2)
        
        methods = self.router.get_registered_methods()
        
        assert set(methods) == {"method1", "method2"}
    
    def test_is_method_registered(self):
        """Test checking if method is registered."""
        mock_handler = Mock()
        self.router.register_handler("test_method", mock_handler)
        
        assert self.router.is_method_registered("test_method") is True
        assert self.router.is_method_registered("unknown_method") is False
    
    def test_get_method_schema(self):
        """Test getting method parameter schema."""
        schema = self.router.get_method_schema("list_emails")
        
        assert schema is not None
        assert "folder" in schema
        assert "unread_only" in schema
        assert "limit" in schema
        
        # Test unknown method
        assert self.router.get_method_schema("unknown_method") is None
    
    def test_unregister_handler(self):
        """Test unregistering a handler."""
        mock_handler = Mock()
        self.router.register_handler("test_method", mock_handler)
        
        # Verify it's registered
        assert self.router.is_method_registered("test_method") is True
        
        # Unregister it
        result = self.router.unregister_handler("test_method")
        assert result is True
        assert self.router.is_method_registered("test_method") is False
        
        # Try to unregister again
        result = self.router.unregister_handler("test_method")
        assert result is False
    
    def test_handler_exception_propagation(self):
        """Test that handler exceptions are properly propagated."""
        def failing_handler():
            raise ValueError("Handler failed")
        
        # Register handler and add schema for the test method
        self.router.register_handler("failing_method", failing_handler)
        self.router._parameter_schemas["failing_method"] = {}
        
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="failing_method",
            params={}
        )
        
        with pytest.raises(ValueError, match="Handler failed"):
            self.router.route_request(request)


class TestRequestRouterIntegration:
    """Integration tests for RequestRouter with realistic scenarios."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.router = RequestRouter()
        
        # Mock handlers for all methods
        self.mock_email_service = Mock()
        self.mock_folder_service = Mock()
        
        # Register handlers
        self.router.register_handler("list_emails", self.mock_email_service.list_emails)
        self.router.register_handler("get_email", self.mock_email_service.get_email)
        self.router.register_handler("search_emails", self.mock_email_service.search_emails)
        self.router.register_handler("get_folders", self.mock_folder_service.get_folders)
    
    def test_complete_list_emails_flow(self):
        """Test complete flow for list_emails method."""
        # Set up mock response
        expected_result = [{"id": "1", "subject": "Test Email"}]
        self.mock_email_service.list_emails.return_value = expected_result
        
        # Create request
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="list_emails",
            params={"folder": "Inbox", "unread_only": True, "limit": 10}
        )
        
        # Route request
        result = self.router.route_request(request)
        
        # Verify result and call
        assert result == expected_result
        self.mock_email_service.list_emails.assert_called_once_with(
            folder="Inbox", 
            unread_only=True, 
            limit=10
        )
    
    def test_complete_get_email_flow(self):
        """Test complete flow for get_email method."""
        # Set up mock response
        expected_result = {"id": "123", "subject": "Test Email", "body": "Test content"}
        self.mock_email_service.get_email.return_value = expected_result
        
        # Create request
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="get_email",
            params={"email_id": "test-email-123"}
        )
        
        # Route request
        result = self.router.route_request(request)
        
        # Verify result and call
        assert result == expected_result
        self.mock_email_service.get_email.assert_called_once_with(email_id="test-email-123")
    
    def test_complete_search_emails_flow(self):
        """Test complete flow for search_emails method."""
        # Set up mock response
        expected_result = [{"id": "1", "subject": "Search Result"}]
        self.mock_email_service.search_emails.return_value = expected_result
        
        # Create request
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="search_emails",
            params={"query": "important meeting", "folder": "Inbox"}
        )
        
        # Route request
        result = self.router.route_request(request)
        
        # Verify result and call
        assert result == expected_result
        self.mock_email_service.search_emails.assert_called_once_with(
            query="important meeting", 
            folder="Inbox", 
            limit=50
        )
    
    def test_complete_get_folders_flow(self):
        """Test complete flow for get_folders method."""
        # Set up mock response
        expected_result = [{"id": "1", "name": "Inbox"}, {"id": "2", "name": "Sent Items"}]
        self.mock_folder_service.get_folders.return_value = expected_result
        
        # Create request
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="get_folders",
            params={}
        )
        
        # Route request
        result = self.router.route_request(request)
        
        # Verify result and call
        assert result == expected_result
        self.mock_folder_service.get_folders.assert_called_once_with()
    
    def test_multiple_requests_different_methods(self):
        """Test handling multiple requests for different methods."""
        # Set up mock responses
        self.mock_folder_service.get_folders.return_value = [{"name": "Inbox"}]
        self.mock_email_service.list_emails.return_value = [{"id": "1"}]
        
        # Create requests
        folders_request = MCPRequest(
            jsonrpc="2.0", id="1", method="get_folders", params={}
        )
        emails_request = MCPRequest(
            jsonrpc="2.0", id="2", method="list_emails", params={"limit": 5}
        )
        
        # Route requests
        folders_result = self.router.route_request(folders_request)
        emails_result = self.router.route_request(emails_request)
        
        # Verify results
        assert folders_result == [{"name": "Inbox"}]
        assert emails_result == [{"id": "1"}]
        
        # Verify calls
        self.mock_folder_service.get_folders.assert_called_once_with()
        self.mock_email_service.list_emails.assert_called_once_with(
            folder=None, unread_only=False, limit=5
        )