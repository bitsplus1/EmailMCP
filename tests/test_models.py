"""Unit tests for data models."""

import pytest
from datetime import datetime
from src.outlook_mcp_server.models import (
    EmailData,
    FolderData,
    MCPRequest,
    MCPResponse,
    ValidationError,
    OutlookMCPError,
    EmailNotFoundError,
    FolderNotFoundError,
    InvalidParameterError
)


class TestEmailData:
    """Test cases for EmailData model."""

    def test_valid_email_data_creation(self):
        """Test creating valid EmailData instance."""
        email = EmailData(
            id="test-id-123",
            subject="Test Subject",
            sender="John Doe",
            sender_email="john@example.com",
            recipients=["jane@example.com"],
            body="Test body",
            is_read=True,
            size=1024
        )
        
        assert email.id == "test-id-123"
        assert email.subject == "Test Subject"
        assert email.sender == "John Doe"
        assert email.sender_email == "john@example.com"
        assert email.recipients == ["jane@example.com"]
        assert email.is_read is True
        assert email.size == 1024

    def test_email_data_validation_empty_id(self):
        """Test validation fails for empty email ID."""
        with pytest.raises(ValidationError, match="Email ID must be a non-empty string"):
            EmailData(
                id="",
                subject="Test",
                sender="John",
                sender_email="john@example.com"
            )

    def test_email_data_validation_invalid_sender_email(self):
        """Test validation fails for invalid sender email."""
        with pytest.raises(ValidationError, match="Sender email must be a valid email address"):
            EmailData(
                id="test-id",
                subject="Test",
                sender="John",
                sender_email="invalid-email"
            )

    def test_email_data_validation_invalid_recipient_email(self):
        """Test validation fails for invalid recipient email."""
        with pytest.raises(ValidationError, match="Invalid recipient email address"):
            EmailData(
                id="test-id",
                subject="Test",
                sender="John",
                sender_email="john@example.com",
                recipients=["invalid-email"]
            )

    def test_email_data_validation_invalid_importance(self):
        """Test validation fails for invalid importance level."""
        with pytest.raises(ValidationError, match="Importance must be 'Low', 'Normal', or 'High'"):
            EmailData(
                id="test-id",
                subject="Test",
                sender="John",
                sender_email="john@example.com",
                importance="Invalid"
            )

    def test_email_data_validation_negative_size(self):
        """Test validation fails for negative size."""
        with pytest.raises(ValidationError, match="Email size cannot be negative"):
            EmailData(
                id="test-id",
                subject="Test",
                sender="John",
                sender_email="john@example.com",
                size=-100
            )

    def test_email_id_validation(self):
        """Test email ID validation method."""
        assert EmailData.validate_email_id("valid-id-123") is True
        assert EmailData.validate_email_id("") is False
        assert EmailData.validate_email_id(None) is False
        assert EmailData.validate_email_id("a" * 256) is False

    def test_email_data_to_dict(self):
        """Test converting EmailData to dictionary."""
        received_time = datetime(2023, 1, 1, 12, 0, 0)
        email = EmailData(
            id="test-id",
            subject="Test",
            sender="John",
            sender_email="john@example.com",
            received_time=received_time
        )
        
        result = email.to_dict()
        
        assert result["id"] == "test-id"
        assert result["subject"] == "Test"
        assert result["received_time"] == "2023-01-01T12:00:00"

    def test_email_data_from_dict(self):
        """Test creating EmailData from dictionary."""
        data = {
            "id": "test-id",
            "subject": "Test",
            "sender": "John",
            "sender_email": "john@example.com",
            "received_time": "2023-01-01T12:00:00",
            "is_read": True
        }
        
        email = EmailData.from_dict(data)
        
        assert email.id == "test-id"
        assert email.subject == "Test"
        assert email.received_time == datetime(2023, 1, 1, 12, 0, 0)
        assert email.is_read is True


class TestFolderData:
    """Test cases for FolderData model."""

    def test_valid_folder_data_creation(self):
        """Test creating valid FolderData instance."""
        folder = FolderData(
            id="folder-id-123",
            name="Inbox",
            full_path="/Inbox",
            item_count=50,
            unread_count=10
        )
        
        assert folder.id == "folder-id-123"
        assert folder.name == "Inbox"
        assert folder.full_path == "/Inbox"
        assert folder.item_count == 50
        assert folder.unread_count == 10

    def test_folder_data_validation_empty_id(self):
        """Test validation fails for empty folder ID."""
        with pytest.raises(ValidationError, match="Folder ID must be a non-empty string"):
            FolderData(
                id="",
                name="Inbox",
                full_path="/Inbox"
            )

    def test_folder_data_validation_negative_item_count(self):
        """Test validation fails for negative item count."""
        with pytest.raises(ValidationError, match="Item count cannot be negative"):
            FolderData(
                id="folder-id",
                name="Inbox",
                full_path="/Inbox",
                item_count=-1
            )

    def test_folder_data_validation_unread_exceeds_total(self):
        """Test validation fails when unread count exceeds total."""
        with pytest.raises(ValidationError, match="Unread count cannot exceed total item count"):
            FolderData(
                id="folder-id",
                name="Inbox",
                full_path="/Inbox",
                item_count=10,
                unread_count=15
            )

    def test_folder_data_validation_invalid_name_characters(self):
        """Test validation fails for invalid folder name characters."""
        with pytest.raises(ValidationError, match="Folder name contains invalid characters"):
            FolderData(
                id="folder-id",
                name="Invalid<Name",
                full_path="/Invalid<Name"
            )

    def test_folder_name_validation(self):
        """Test folder name validation method."""
        assert FolderData.validate_folder_name("Valid Folder") is True
        assert FolderData.validate_folder_name("Inbox") is True
        assert FolderData.validate_folder_name("") is False
        assert FolderData.validate_folder_name("Invalid<Name") is False
        assert FolderData.validate_folder_name("Invalid|Name") is False
        assert FolderData.validate_folder_name("a" * 256) is False

    def test_folder_data_to_dict(self):
        """Test converting FolderData to dictionary."""
        folder = FolderData(
            id="folder-id",
            name="Inbox",
            full_path="/Inbox",
            item_count=50
        )
        
        result = folder.to_dict()
        
        assert result["id"] == "folder-id"
        assert result["name"] == "Inbox"
        assert result["item_count"] == 50

    def test_folder_data_from_dict(self):
        """Test creating FolderData from dictionary."""
        data = {
            "id": "folder-id",
            "name": "Inbox",
            "full_path": "/Inbox",
            "item_count": 50,
            "unread_count": 10
        }
        
        folder = FolderData.from_dict(data)
        
        assert folder.id == "folder-id"
        assert folder.name == "Inbox"
        assert folder.item_count == 50
        assert folder.unread_count == 10


class TestMCPRequest:
    """Test cases for MCPRequest model."""

    def test_valid_mcp_request_creation(self):
        """Test creating valid MCPRequest instance."""
        request = MCPRequest(
            jsonrpc="2.0",
            id="req-123",
            method="list_emails",
            params={"folder": "Inbox"}
        )
        
        assert request.jsonrpc == "2.0"
        assert request.id == "req-123"
        assert request.method == "list_emails"
        assert request.params == {"folder": "Inbox"}

    def test_mcp_request_validation_invalid_jsonrpc(self):
        """Test validation fails for invalid JSON-RPC version."""
        with pytest.raises(ValidationError, match="JSON-RPC version must be '2.0'"):
            MCPRequest(
                jsonrpc="1.0",
                id="req-123",
                method="list_emails",
                params={}
            )

    def test_mcp_request_validation_empty_id(self):
        """Test validation fails for empty request ID."""
        with pytest.raises(ValidationError, match="Request ID must be a non-empty string"):
            MCPRequest(
                jsonrpc="2.0",
                id="",
                method="list_emails",
                params={}
            )

    def test_mcp_request_validation_invalid_method_name(self):
        """Test validation fails for invalid method name."""
        with pytest.raises(ValidationError, match="Invalid method name format"):
            MCPRequest(
                jsonrpc="2.0",
                id="req-123",
                method="123invalid",
                params={}
            )

    def test_search_query_validation(self):
        """Test search query validation method."""
        assert MCPRequest.validate_search_query("valid query") is True
        assert MCPRequest.validate_search_query("") is False
        assert MCPRequest.validate_search_query(None) is False
        assert MCPRequest.validate_search_query("a" * 1001) is False

    def test_mcp_request_to_dict(self):
        """Test converting MCPRequest to dictionary."""
        request = MCPRequest(
            jsonrpc="2.0",
            id="req-123",
            method="list_emails",
            params={"folder": "Inbox"}
        )
        
        result = request.to_dict()
        
        assert result["jsonrpc"] == "2.0"
        assert result["id"] == "req-123"
        assert result["method"] == "list_emails"
        assert result["params"] == {"folder": "Inbox"}

    def test_mcp_request_from_dict(self):
        """Test creating MCPRequest from dictionary."""
        data = {
            "jsonrpc": "2.0",
            "id": "req-123",
            "method": "list_emails",
            "params": {"folder": "Inbox"}
        }
        
        request = MCPRequest.from_dict(data)
        
        assert request.jsonrpc == "2.0"
        assert request.id == "req-123"
        assert request.method == "list_emails"
        assert request.params == {"folder": "Inbox"}


class TestMCPResponse:
    """Test cases for MCPResponse model."""

    def test_valid_mcp_response_creation_with_result(self):
        """Test creating valid MCPResponse with result."""
        response = MCPResponse(
            jsonrpc="2.0",
            id="req-123",
            result={"emails": []}
        )
        
        assert response.jsonrpc == "2.0"
        assert response.id == "req-123"
        assert response.result == {"emails": []}
        assert response.error is None

    def test_valid_mcp_response_creation_with_error(self):
        """Test creating valid MCPResponse with error."""
        error = {"code": -32000, "message": "Test error"}
        response = MCPResponse(
            jsonrpc="2.0",
            id="req-123",
            error=error
        )
        
        assert response.jsonrpc == "2.0"
        assert response.id == "req-123"
        assert response.result is None
        assert response.error == error

    def test_mcp_response_validation_both_result_and_error(self):
        """Test validation fails when both result and error are present."""
        with pytest.raises(ValidationError, match="Response cannot have both result and error"):
            MCPResponse(
                jsonrpc="2.0",
                id="req-123",
                result={"data": "test"},
                error={"code": -32000, "message": "error"}
            )

    def test_mcp_response_validation_neither_result_nor_error(self):
        """Test validation fails when neither result nor error are present."""
        with pytest.raises(ValidationError, match="Response must have either result or error"):
            MCPResponse(
                jsonrpc="2.0",
                id="req-123"
            )

    def test_mcp_response_validation_invalid_error_format(self):
        """Test validation fails for invalid error format."""
        with pytest.raises(ValidationError, match="Error must have an integer code"):
            MCPResponse(
                jsonrpc="2.0",
                id="req-123",
                error={"message": "error without code"}
            )

    def test_create_success_response(self):
        """Test creating success response."""
        response = MCPResponse.create_success("req-123", {"data": "test"})
        
        assert response.jsonrpc == "2.0"
        assert response.id == "req-123"
        assert response.result == {"data": "test"}
        assert response.error is None

    def test_create_error_response(self):
        """Test creating error response."""
        response = MCPResponse.create_error("req-123", -32000, "Test error", {"detail": "info"})
        
        assert response.jsonrpc == "2.0"
        assert response.id == "req-123"
        assert response.result is None
        assert response.error["code"] == -32000
        assert response.error["message"] == "Test error"
        assert response.error["data"] == {"detail": "info"}

    def test_mcp_response_to_dict(self):
        """Test converting MCPResponse to dictionary."""
        response = MCPResponse(
            jsonrpc="2.0",
            id="req-123",
            result={"data": "test"}
        )
        
        result = response.to_dict()
        
        assert result["jsonrpc"] == "2.0"
        assert result["id"] == "req-123"
        assert result["result"] == {"data": "test"}
        assert "error" not in result


class TestExceptions:
    """Test cases for exception classes."""

    def test_outlook_mcp_error_creation(self):
        """Test creating OutlookMCPError."""
        error = OutlookMCPError("Test error", -32000, {"detail": "info"})
        
        assert str(error) == "Test error"
        assert error.error_code == -32000
        assert error.details == {"detail": "info"}

    def test_outlook_mcp_error_to_dict(self):
        """Test converting OutlookMCPError to dictionary."""
        error = OutlookMCPError("Test error", -32000, {"detail": "info"})
        result = error.to_dict()
        
        assert result["code"] == -32000
        assert result["message"] == "Test error"
        assert result["data"]["type"] == "OutlookMCPError"
        assert result["data"]["details"] == {"detail": "info"}

    def test_validation_error_with_field(self):
        """Test ValidationError with field information."""
        error = ValidationError("Invalid field", "email")
        
        assert error.error_code == -32602
        assert error.details["field"] == "email"

    def test_email_not_found_error(self):
        """Test EmailNotFoundError."""
        error = EmailNotFoundError("test-email-id")
        
        assert "test-email-id" in str(error)
        assert error.details["email_id"] == "test-email-id"
        assert error.error_code == -32003

    def test_folder_not_found_error(self):
        """Test FolderNotFoundError."""
        error = FolderNotFoundError("TestFolder")
        
        assert "TestFolder" in str(error)
        assert error.details["folder_name"] == "TestFolder"
        assert error.error_code == -32004

    def test_invalid_parameter_error(self):
        """Test InvalidParameterError."""
        error = InvalidParameterError("limit", "Limit must be positive")
        
        assert "Limit must be positive" in str(error)
        assert error.details["parameter"] == "limit"
        assert error.error_code == -32602