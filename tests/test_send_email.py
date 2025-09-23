"""Unit tests for send_email functionality."""

import pytest
import asyncio
from unittest.mock import Mock, patch, MagicMock
from src.outlook_mcp_server.services.email_service import EmailService
from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
from src.outlook_mcp_server.models.exceptions import (
    ValidationError,
    PermissionError,
    OutlookConnectionError
)


class TestSendEmail:
    """Test cases for send_email functionality."""
    
    @pytest.fixture
    def mock_outlook_adapter(self):
        """Create a mock Outlook adapter."""
        adapter = Mock(spec=OutlookAdapter)
        adapter.is_connected.return_value = True
        adapter.send_email.return_value = "test_email_id_123"
        return adapter
    
    @pytest.fixture
    def email_service(self, mock_outlook_adapter):
        """Create an EmailService instance with mocked adapter."""
        return EmailService(mock_outlook_adapter)
    
    @pytest.mark.asyncio
    async def test_send_email_success(self, email_service, mock_outlook_adapter):
        """Test successful email sending."""
        # Arrange
        to_recipients = ["test@example.com"]
        subject = "Test Subject"
        body = "Test body content"
        
        # Act
        result = await email_service.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body
        )
        
        # Assert
        assert result["status"] == "sent"
        assert result["email_id"] == "test_email_id_123"
        assert result["recipients"]["to"] == to_recipients
        assert result["subject"] == subject
        assert result["body_format"] == "html"
        assert result["importance"] == "normal"
        assert result["attachments_count"] == 0
        
        # Verify adapter was called correctly
        mock_outlook_adapter.send_email.assert_called_once_with(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=None,
            bcc_recipients=None,
            body_format="html",
            importance="normal",
            attachments=None,
            save_to_sent_items=True
        )
    
    @pytest.mark.asyncio
    async def test_send_email_with_all_parameters(self, email_service, mock_outlook_adapter):
        """Test email sending with all parameters."""
        # Arrange
        to_recipients = ["test1@example.com", "test2@example.com"]
        cc_recipients = ["cc@example.com"]
        bcc_recipients = ["bcc@example.com"]
        subject = "Test Subject"
        body = "Test body content"
        attachments = ["C:\\test\\file1.pdf", "C:\\test\\file2.docx"]
        
        # Act
        result = await email_service.send_email(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            body_format="text",
            importance="high",
            attachments=attachments,
            save_to_sent_items=False
        )
        
        # Assert
        assert result["status"] == "sent"
        assert result["recipients"]["to"] == to_recipients
        assert result["recipients"]["cc"] == cc_recipients
        assert result["recipients"]["bcc"] == bcc_recipients
        assert result["body_format"] == "text"
        assert result["importance"] == "high"
        assert result["attachments_count"] == 2
        
        # Verify adapter was called correctly
        mock_outlook_adapter.send_email.assert_called_once_with(
            to_recipients=to_recipients,
            subject=subject,
            body=body,
            cc_recipients=cc_recipients,
            bcc_recipients=bcc_recipients,
            body_format="text",
            importance="high",
            attachments=attachments,
            save_to_sent_items=False
        )
    
    @pytest.mark.asyncio
    async def test_send_email_validation_errors(self, email_service):
        """Test validation errors for send_email."""
        # Test empty recipients
        with pytest.raises(ValidationError) as exc_info:
            await email_service.send_email(
                to_recipients=[],
                subject="Test",
                body="Test body"
            )
        assert "At least one recipient is required" in str(exc_info.value)
        
        # Test invalid recipients type
        with pytest.raises(ValidationError) as exc_info:
            await email_service.send_email(
                to_recipients="not_a_list",
                subject="Test",
                body="Test body"
            )
        assert "At least one recipient is required" in str(exc_info.value)
        
        # Test empty subject
        with pytest.raises(ValidationError) as exc_info:
            await email_service.send_email(
                to_recipients=["test@example.com"],
                subject="",
                body="Test body"
            )
        assert "Subject is required" in str(exc_info.value)
        
        # Test empty body
        with pytest.raises(ValidationError) as exc_info:
            await email_service.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body=""
            )
        assert "Body is required" in str(exc_info.value)
        
        # Test invalid email address
        with pytest.raises(ValidationError) as exc_info:
            await email_service.send_email(
                to_recipients=["invalid_email"],
                subject="Test",
                body="Test body"
            )
        assert "Invalid email address" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_send_email_outlook_not_connected(self, email_service, mock_outlook_adapter):
        """Test send_email when Outlook is not connected."""
        # Arrange
        mock_outlook_adapter.is_connected.return_value = False
        
        # Act & Assert
        with pytest.raises(OutlookConnectionError) as exc_info:
            await email_service.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body"
            )
        assert "Not connected to Outlook" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_send_email_permission_error(self, email_service, mock_outlook_adapter):
        """Test send_email when permission is denied."""
        # Arrange
        mock_outlook_adapter.send_email.side_effect = Exception("Access denied by policy")
        
        # Act & Assert
        with pytest.raises(PermissionError) as exc_info:
            await email_service.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body"
            )
        assert "Permission denied to send email" in str(exc_info.value)
    
    @pytest.mark.asyncio
    async def test_send_email_connection_error(self, email_service, mock_outlook_adapter):
        """Test send_email when connection fails."""
        # Arrange
        mock_outlook_adapter.send_email.side_effect = Exception("Connection lost")
        
        # Act & Assert
        with pytest.raises(OutlookConnectionError) as exc_info:
            await email_service.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body"
            )
        assert "Failed to send email" in str(exc_info.value)
    
    def test_validate_email_address(self, email_service):
        """Test email address validation."""
        # Valid email addresses
        assert email_service._validate_email_address("test@example.com") == True
        assert email_service._validate_email_address("user.name+tag@domain.co.uk") == True
        assert email_service._validate_email_address("test123@test-domain.org") == True
        
        # Invalid email addresses
        assert email_service._validate_email_address("") == False
        assert email_service._validate_email_address("invalid") == False
        assert email_service._validate_email_address("@domain.com") == False
        assert email_service._validate_email_address("user@") == False
        assert email_service._validate_email_address("user@domain") == False
        assert email_service._validate_email_address(None) == False
        assert email_service._validate_email_address(123) == False


class TestOutlookAdapterSendEmail:
    """Test cases for OutlookAdapter send_email functionality."""
    
    @pytest.fixture
    def outlook_adapter(self):
        """Create an OutlookAdapter instance."""
        adapter = OutlookAdapter()
        adapter._connected = True
        # Mock the is_connected method to return True
        adapter.is_connected = Mock(return_value=True)
        return adapter
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_send_email_success(self, mock_win32com, mock_pythoncom, outlook_adapter):
        """Test successful email sending in OutlookAdapter."""
        # Arrange
        mock_outlook_app = Mock()
        mock_mail_item = Mock()
        mock_recipients = Mock()
        mock_attachments = Mock()
        
        outlook_adapter._outlook_app = mock_outlook_app
        mock_outlook_app.CreateItem.return_value = mock_mail_item
        mock_mail_item.Recipients = mock_recipients
        mock_mail_item.Attachments = mock_attachments
        mock_mail_item.EntryID = "test_entry_id"
        
        # Act
        result = outlook_adapter.send_email(
            to_recipients=["test@example.com"],
            subject="Test Subject",
            body="Test body"
        )
        
        # Assert
        assert result == "test_entry_id"
        mock_outlook_app.CreateItem.assert_called_once_with(0)  # olMailItem
        mock_recipients.Add.assert_called_once_with("test@example.com")
        mock_recipients.ResolveAll.assert_called_once()
        mock_mail_item.Send.assert_called_once()
        
        # Verify email properties were set
        assert mock_mail_item.Subject == "Test Subject"
        assert mock_mail_item.HTMLBody == "Test body"
        assert mock_mail_item.Importance == 1  # normal importance
    
    def test_send_email_validation_errors(self, outlook_adapter):
        """Test validation errors in OutlookAdapter."""
        # Test empty recipients
        with pytest.raises(ValidationError):
            outlook_adapter.send_email(
                to_recipients=[],
                subject="Test",
                body="Test body"
            )
        
        # Test invalid email format
        with pytest.raises(ValidationError):
            outlook_adapter.send_email(
                to_recipients=["invalid_email"],
                subject="Test",
                body="Test body"
            )
        
        # Test invalid body format
        with pytest.raises(ValidationError):
            outlook_adapter.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body",
                body_format="invalid"
            )
        
        # Test invalid importance
        with pytest.raises(ValidationError):
            outlook_adapter.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body",
                importance="invalid"
            )
    
    def test_send_email_not_connected(self, outlook_adapter):
        """Test send_email when not connected."""
        # Arrange
        outlook_adapter._connected = False
        
        # Act & Assert
        with pytest.raises(OutlookConnectionError):
            outlook_adapter.send_email(
                to_recipients=["test@example.com"],
                subject="Test",
                body="Test body"
            )
    
    def test_validate_email_address(self, outlook_adapter):
        """Test email address validation in OutlookAdapter."""
        # Valid emails
        assert outlook_adapter._validate_email_address("test@example.com") == True
        assert outlook_adapter._validate_email_address("user+tag@domain.org") == True
        
        # Invalid emails
        assert outlook_adapter._validate_email_address("invalid") == False
        assert outlook_adapter._validate_email_address("") == False
        assert outlook_adapter._validate_email_address(None) == False


if __name__ == "__main__":
    pytest.main([__file__])