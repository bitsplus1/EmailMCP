"""Unit tests for MCP Protocol Handler."""

import pytest
from unittest.mock import Mock, patch
from src.outlook_mcp_server.protocol.mcp_protocol_handler import MCPProtocolHandler
from src.outlook_mcp_server.models.mcp_models import MCPRequest, MCPResponse
from src.outlook_mcp_server.models.exceptions import ValidationError


class TestMCPProtocolHandler:
    """Test cases for MCPProtocolHandler class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.handler = MCPProtocolHandler()
    
    def test_initialization(self):
        """Test protocol handler initialization."""
        assert self.handler.client_info is None
        assert self.handler.session_active is False
        assert self.handler.PROTOCOL_VERSION == "2024-11-05"
        assert "tools" in self.handler.SERVER_CAPABILITIES
        assert len(self.handler.SERVER_CAPABILITIES["tools"]) == 4
    
    def test_handle_handshake_success(self):
        """Test successful handshake."""
        client_info = {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {
                "name": "test-client",
                "version": "1.0.0"
            }
        }
        
        response = self.handler.handle_handshake(client_info)
        
        assert self.handler.session_active is True
        assert self.handler.client_info == client_info
        assert response["protocolVersion"] == "2024-11-05"
        assert "capabilities" in response
        assert "serverInfo" in response
        assert response["serverInfo"]["name"] == "outlook-mcp-server"
    
    def test_handle_handshake_invalid_client_info(self):
        """Test handshake with invalid client info."""
        # Test with non-dict client info
        with pytest.raises(ValidationError, match="Client info must be a dictionary"):
            self.handler.handle_handshake("invalid")
        
        # Test with missing protocol version
        with pytest.raises(ValidationError, match="Client info must include protocolVersion"):
            self.handler.handle_handshake({})
        
        # Test with invalid protocol version type
        with pytest.raises(ValidationError, match="Protocol version must be a string"):
            self.handler.handle_handshake({"protocolVersion": 123})
    
    def test_handle_handshake_unsupported_version(self):
        """Test handshake with unsupported protocol version."""
        client_info = {
            "protocolVersion": "1.0.0",
            "capabilities": {}
        }
        
        with pytest.raises(ValidationError, match="Unsupported protocol version: 1.0.0"):
            self.handler.handle_handshake(client_info)
        
        assert self.handler.session_active is False
    
    def test_process_request_no_session(self):
        """Test processing request without active session."""
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="list_emails",
            params={}
        )
        
        response = self.handler.process_request(request)
        
        assert isinstance(response, MCPResponse)
        assert response.error is not None
        assert response.error["code"] == -32000
        assert "No active session" in response.error["message"]
    
    def test_process_request_with_session(self):
        """Test processing request with active session."""
        # Set up session
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="list_emails",
            params={}
        )
        
        response = self.handler.process_request(request)
        
        # Should return None for valid requests (handled by router)
        assert response is None
    
    def test_process_request_invalid_method(self):
        """Test processing request with invalid method."""
        # Set up session
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="invalid_method",
            params={}
        )
        
        response = self.handler.process_request(request)
        
        assert isinstance(response, MCPResponse)
        assert response.error is not None
        assert response.error["code"] == -32601
        assert "Method 'invalid_method' not found" in response.error["message"]
    
    def test_process_request_invalid_params(self):
        """Test processing request with invalid parameters."""
        # Set up session
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        
        # Test missing required parameter
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="get_email",
            params={}  # Missing required email_id
        )
        
        response = self.handler.process_request(request)
        
        assert isinstance(response, MCPResponse)
        assert response.error is not None
        assert response.error["code"] == -32602
        assert "Missing required parameter: email_id" in response.error["message"]
    
    def test_process_request_validation_error(self):
        """Test processing request with validation error."""
        # Set up session
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        
        # Create a request that will fail validation during processing
        # We'll mock the validate method to raise an error during process_request
        request = MCPRequest(
            jsonrpc="2.0",
            id="test-1",
            method="list_emails",
            params={}
        )
        
        # Mock the validate method to raise ValidationError during processing
        with patch.object(request, 'validate', side_effect=ValidationError("Invalid request")):
            response = self.handler.process_request(request)
            
            assert isinstance(response, MCPResponse)
            assert response.error is not None
            assert response.error["code"] == -32003
    
    def test_format_response_success(self):
        """Test formatting successful response."""
        data = {"emails": [{"id": "1", "subject": "Test"}]}
        request_id = "test-1"
        
        formatted = self.handler.format_response(data, request_id)
        
        assert formatted["jsonrpc"] == "2.0"
        assert formatted["id"] == request_id
        assert formatted["result"] == data
        assert "error" not in formatted
    
    def test_format_response_error_handling(self):
        """Test response formatting error handling."""
        # Test with data that might cause formatting issues
        with patch('src.outlook_mcp_server.protocol.mcp_protocol_handler.MCPResponse.create_success', 
                  side_effect=Exception("Formatting error")):
            formatted = self.handler.format_response({}, "test-1")
            
            assert formatted["jsonrpc"] == "2.0"
            assert formatted["id"] == "test-1"
            assert "error" in formatted
            assert formatted["error"]["code"] == -32603
    
    def test_format_error_validation_error(self):
        """Test formatting validation error."""
        error = ValidationError("Invalid parameter")
        request_id = "test-1"
        
        formatted = self.handler.format_error(error, request_id)
        
        assert formatted["jsonrpc"] == "2.0"
        assert formatted["id"] == request_id
        assert "error" in formatted
        assert formatted["error"]["code"] == -32003
        assert formatted["error"]["message"] == "Invalid parameter"
        assert formatted["error"]["data"]["type"] == "ValidationError"
    
    def test_format_error_outlook_connection_error(self):
        """Test formatting Outlook connection error."""
        error = Exception("Outlook connection failed")
        request_id = "test-1"
        
        formatted = self.handler.format_error(error, request_id)
        
        assert formatted["error"]["code"] == -32001
        assert "Failed to connect to Outlook" in formatted["error"]["message"]
        assert formatted["error"]["data"]["details"] == "Outlook connection failed"
    
    def test_format_error_outlook_access_error(self):
        """Test formatting Outlook access error."""
        error = Exception("Outlook access denied")
        request_id = "test-1"
        
        formatted = self.handler.format_error(error, request_id)
        
        assert formatted["error"]["code"] == -32002
        assert "Outlook access error" in formatted["error"]["message"]
    
    def test_format_error_generic_error(self):
        """Test formatting generic error."""
        error = Exception("Generic error")
        request_id = "test-1"
        
        formatted = self.handler.format_error(error, request_id)
        
        assert formatted["error"]["code"] == -32000
        assert "Server error occurred" in formatted["error"]["message"]
    
    def test_format_error_exception_handling(self):
        """Test error formatting exception handling."""
        error = Exception("Test error")
        
        # Mock the _categorize_error method to raise an exception
        with patch.object(self.handler, '_categorize_error', side_effect=Exception("Format error")):
            formatted = self.handler.format_error(error, "test-1")
            
            assert formatted["jsonrpc"] == "2.0"
            assert formatted["id"] == "test-1"
            assert "error" in formatted
            assert formatted["error"]["code"] == -32603
    
    def test_validate_method_params_list_emails(self):
        """Test parameter validation for list_emails method."""
        # Valid parameters
        assert self.handler._validate_method_params("list_emails", {}) is None
        assert self.handler._validate_method_params("list_emails", {"folder": "Inbox"}) is None
        assert self.handler._validate_method_params("list_emails", {"unread_only": True}) is None
        assert self.handler._validate_method_params("list_emails", {"limit": 10}) is None
        
        # Invalid parameters
        error = self.handler._validate_method_params("list_emails", {"limit": "invalid"})
        assert "must be an integer" in error
        
        error = self.handler._validate_method_params("list_emails", {"limit": 0})
        assert "must be at least 1" in error
        
        error = self.handler._validate_method_params("list_emails", {"limit": 2000})
        assert "must be at most 1000" in error
    
    def test_validate_method_params_get_email(self):
        """Test parameter validation for get_email method."""
        # Valid parameters
        assert self.handler._validate_method_params("get_email", {"email_id": "123"}) is None
        
        # Invalid parameters
        error = self.handler._validate_method_params("get_email", {})
        assert "Missing required parameter: email_id" in error
        
        error = self.handler._validate_method_params("get_email", {"email_id": 123})
        assert "must be a string" in error
    
    def test_validate_method_params_search_emails(self):
        """Test parameter validation for search_emails method."""
        # Valid parameters
        assert self.handler._validate_method_params("search_emails", {"query": "test"}) is None
        assert self.handler._validate_method_params("search_emails", {"query": "test", "folder": "Inbox"}) is None
        
        # Invalid parameters
        error = self.handler._validate_method_params("search_emails", {})
        assert "Missing required parameter: query" in error
        
        error = self.handler._validate_method_params("search_emails", {"query": ""})
        assert "cannot be empty" in error
        
        error = self.handler._validate_method_params("search_emails", {"query": "   "})
        assert "cannot be empty" in error
    
    def test_validate_method_params_get_folders(self):
        """Test parameter validation for get_folders method."""
        # Valid parameters (no parameters required)
        assert self.handler._validate_method_params("get_folders", {}) is None
        assert self.handler._validate_method_params("get_folders", {"extra": "ignored"}) is None
    
    def test_is_method_supported(self):
        """Test method support checking."""
        assert self.handler._is_method_supported("list_emails") is True
        assert self.handler._is_method_supported("get_email") is True
        assert self.handler._is_method_supported("search_emails") is True
        assert self.handler._is_method_supported("get_folders") is True
        assert self.handler._is_method_supported("invalid_method") is False
    
    def test_is_protocol_version_supported(self):
        """Test protocol version support checking."""
        assert self.handler._is_protocol_version_supported("2024-11-05") is True
        assert self.handler._is_protocol_version_supported("1.0.0") is False
        assert self.handler._is_protocol_version_supported("") is False
    
    def test_get_server_info(self):
        """Test getting server information."""
        info = self.handler.get_server_info()
        
        assert info["name"] == "outlook-mcp-server"
        assert info["version"] == "1.0.0"
        assert info["protocolVersion"] == "2024-11-05"
        assert "capabilities" in info
        assert "description" in info
    
    def test_session_management(self):
        """Test session management methods."""
        # Initially no session
        assert self.handler.is_session_active() is False
        
        # Start session
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        assert self.handler.is_session_active() is True
        
        # Close session
        self.handler.close_session()
        assert self.handler.is_session_active() is False
        assert self.handler.client_info is None
    
    def test_server_capabilities_structure(self):
        """Test server capabilities structure."""
        capabilities = self.handler.SERVER_CAPABILITIES
        
        assert "tools" in capabilities
        tools = capabilities["tools"]
        
        # Check all required tools are present
        tool_names = {tool["name"] for tool in tools}
        expected_tools = {"list_emails", "get_email", "search_emails", "get_folders"}
        assert tool_names == expected_tools
        
        # Check each tool has required fields
        for tool in tools:
            assert "name" in tool
            assert "description" in tool
            assert "inputSchema" in tool
            assert "type" in tool["inputSchema"]
            assert "properties" in tool["inputSchema"]
    
    def test_error_codes_defined(self):
        """Test that all required error codes are defined."""
        expected_codes = [
            "PARSE_ERROR", "INVALID_REQUEST", "METHOD_NOT_FOUND",
            "INVALID_PARAMS", "INTERNAL_ERROR", "SERVER_ERROR",
            "OUTLOOK_CONNECTION_ERROR", "OUTLOOK_ACCESS_ERROR", "VALIDATION_ERROR"
        ]
        
        for code in expected_codes:
            assert code in self.handler.ERROR_CODES
            assert isinstance(self.handler.ERROR_CODES[code], int)


class TestMCPProtocolHandlerIntegration:
    """Integration tests for MCP Protocol Handler."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.handler = MCPProtocolHandler()
    
    def test_full_handshake_and_request_flow(self):
        """Test complete handshake and request processing flow."""
        # Step 1: Handshake
        client_info = {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "test-client"}
        }
        
        handshake_response = self.handler.handle_handshake(client_info)
        assert self.handler.is_session_active()
        assert "capabilities" in handshake_response
        
        # Step 2: Valid request
        request = MCPRequest(
            jsonrpc="2.0",
            id="req-1",
            method="list_emails",
            params={"folder": "Inbox", "limit": 10}
        )
        
        response = self.handler.process_request(request)
        assert response is None  # Valid request should return None for router handling
        
        # Step 3: Format successful response
        mock_data = {"emails": [{"id": "1", "subject": "Test Email"}]}
        formatted_response = self.handler.format_response(mock_data, "req-1")
        
        assert formatted_response["jsonrpc"] == "2.0"
        assert formatted_response["id"] == "req-1"
        assert formatted_response["result"] == mock_data
        
        # Step 4: Close session
        self.handler.close_session()
        assert not self.handler.is_session_active()
    
    def test_error_handling_flow(self):
        """Test error handling throughout the request flow."""
        # Step 1: Request without handshake
        request = MCPRequest(
            jsonrpc="2.0",
            id="req-1",
            method="list_emails",
            params={}
        )
        
        response = self.handler.process_request(request)
        assert response.error is not None
        assert "No active session" in response.error["message"]
        
        # Step 2: Handshake and invalid method
        client_info = {"protocolVersion": "2024-11-05"}
        self.handler.handle_handshake(client_info)
        
        invalid_request = MCPRequest(
            jsonrpc="2.0",
            id="req-2",
            method="invalid_method",
            params={}
        )
        
        response = self.handler.process_request(invalid_request)
        assert response.error is not None
        assert response.error["code"] == -32601
        
        # Step 3: Format error response
        test_error = ValidationError("Test validation error")
        formatted_error = self.handler.format_error(test_error, "req-3")
        
        assert formatted_error["jsonrpc"] == "2.0"
        assert formatted_error["id"] == "req-3"
        assert formatted_error["error"]["code"] == -32003
        assert "Test validation error" in formatted_error["error"]["message"]