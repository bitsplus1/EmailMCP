"""Unit tests for OutlookAdapter with mocked COM objects."""

import pytest
from unittest.mock import Mock, patch, MagicMock, PropertyMock
import sys
from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
from src.outlook_mcp_server.models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    EmailNotFoundError,
    PermissionError
)


class TestOutlookAdapter:
    """Test cases for OutlookAdapter class."""
    
    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.adapter = OutlookAdapter()
    
    def teardown_method(self):
        """Clean up after each test method."""
        if hasattr(self.adapter, 'disconnect'):
            try:
                self.adapter.disconnect()
            except:
                pass
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_connect_success_existing_instance(self, mock_client, mock_pythoncom):
        """Test successful connection to existing Outlook instance."""
        # Mock successful connection to existing instance
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Test connection
        result = self.adapter.connect()
        
        # Assertions
        assert result is True
        assert self.adapter.is_connected() is True
        mock_pythoncom.CoInitialize.assert_called_once()
        mock_client.GetActiveObject.assert_called_once_with("Outlook.Application")
        mock_outlook_app.GetNamespace.assert_called_once_with("MAPI")
        mock_namespace.GetDefaultFolder.assert_called_with(6)  # Inbox
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_connect_success_new_instance(self, mock_client, mock_pythoncom):
        """Test successful connection by creating new Outlook instance."""
        # Mock failure to get existing instance, success with new instance
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.side_effect = Exception("No active object")
        mock_client.Dispatch.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Test connection
        result = self.adapter.connect()
        
        # Assertions
        assert result is True
        assert self.adapter.is_connected() is True
        mock_client.GetActiveObject.assert_called_once_with("Outlook.Application")
        mock_client.Dispatch.assert_called_once_with("Outlook.Application")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_connect_failure_no_outlook(self, mock_client, mock_pythoncom):
        """Test connection failure when Outlook is not available."""
        # Mock failure to connect
        mock_client.GetActiveObject.side_effect = Exception("No Outlook")
        mock_client.Dispatch.side_effect = Exception("Cannot create Outlook")
        
        # Test connection failure
        with pytest.raises(OutlookConnectionError) as exc_info:
            self.adapter.connect()
        
        assert "Failed to connect to Outlook" in str(exc_info.value)
        assert self.adapter.is_connected() is False
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_connect_failure_namespace_error(self, mock_client, mock_pythoncom):
        """Test connection failure when namespace cannot be accessed."""
        # Mock Outlook app creation success but namespace failure
        mock_outlook_app = Mock()
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.side_effect = Exception("Namespace error")
        
        # Test connection failure
        with pytest.raises(OutlookConnectionError):
            self.adapter.connect()
        
        assert self.adapter.is_connected() is False
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_connect_failure_inbox_access(self, mock_client, mock_pythoncom):
        """Test connection failure when inbox cannot be accessed."""
        # Mock successful app and namespace creation but inbox access failure
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = None  # No inbox access
        
        # Test connection failure
        with pytest.raises(OutlookConnectionError) as exc_info:
            self.adapter.connect()
        
        assert "Connection test failed" in str(exc_info.value)
        assert self.adapter.is_connected() is False
    
    def test_disconnect(self):
        """Test disconnection from Outlook."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        self.adapter._namespace = Mock()
        
        # Test disconnect
        self.adapter.disconnect()
        
        # Assertions
        assert self.adapter._connected is False
        assert self.adapter._outlook_app is None
        assert self.adapter._namespace is None
    
    def test_is_connected_false_when_not_connected(self):
        """Test is_connected returns False when not connected."""
        assert self.adapter.is_connected() is False
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_is_connected_true_when_connected(self, mock_client, mock_pythoncom):
        """Test is_connected returns True when properly connected."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test is_connected
        assert self.adapter.is_connected() is True
    
    def test_is_connected_false_when_connection_lost(self):
        """Test is_connected returns False when connection is lost."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        mock_namespace = Mock()
        self.adapter._namespace = mock_namespace
        
        # Mock connection test failure
        mock_namespace.GetDefaultFolder.side_effect = Exception("Connection lost")
        
        # Test is_connected
        assert self.adapter.is_connected() is False
        assert self.adapter._connected is False
    
    def test_get_namespace_when_not_connected(self):
        """Test get_namespace raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.get_namespace()
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_namespace_when_connected(self, mock_client, mock_pythoncom):
        """Test get_namespace returns namespace when connected."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test get_namespace
        result = self.adapter.get_namespace()
        assert result == mock_namespace
    
    def test_get_folder_by_name_when_not_connected(self):
        """Test get_folder_by_name raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.get_folder_by_name("Inbox")
    
    def test_get_folder_by_name_invalid_name(self):
        """Test get_folder_by_name with invalid folder name."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        self.adapter._namespace = Mock()
        
        # Test with None name
        with pytest.raises(FolderNotFoundError):
            self.adapter.get_folder_by_name(None)
        
        # Test with empty name
        with pytest.raises(FolderNotFoundError):
            self.adapter.get_folder_by_name("")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folder_by_name_default_folder(self, mock_client, mock_pythoncom):
        """Test get_folder_by_name with default folder names."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_folder  # Always return the folder
        
        self.adapter.connect()
        
        # Test getting default folder
        result = self.adapter.get_folder_by_name("Inbox")
        assert result == mock_folder
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folder_by_name_not_found(self, mock_client, mock_pythoncom):
        """Test get_folder_by_name when folder is not found."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folders = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.Folders = []  # No folders
        
        self.adapter.connect()
        
        # Test folder not found
        with pytest.raises(FolderNotFoundError):
            self.adapter.get_folder_by_name("NonExistentFolder")
    
    def test_get_email_by_id_when_not_connected(self):
        """Test get_email_by_id raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.get_email_by_id("some_id")
    
    def test_get_email_by_id_invalid_id(self):
        """Test get_email_by_id with invalid email ID."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        self.adapter._namespace = Mock()
        
        # Test with None ID
        with pytest.raises(EmailNotFoundError):
            self.adapter.get_email_by_id(None)
        
        # Test with empty ID
        with pytest.raises(EmailNotFoundError):
            self.adapter.get_email_by_id("")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_success(self, mock_client, mock_pythoncom):
        """Test successful email retrieval by ID."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create detailed mock email
        mock_email = Mock()
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "test_email_id"
        mock_email.Subject = "Test Email Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "test@example.com"
        mock_email.Body = "Test body content"
        mock_email.HTMLBody = "<p>Test body content</p>"
        mock_email.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email.SentOn = datetime(2023, 1, 15, 10, 25, 0)
        mock_email.UnRead = False
        mock_email.Importance = 1  # Normal
        mock_email.Size = 1024
        
        # Mock parent folder
        mock_parent = Mock()
        mock_parent.Name = "Inbox"
        mock_email.Parent = mock_parent
        
        # Mock attachments
        mock_attachments = Mock()
        mock_attachments.Count = 0
        mock_email.Attachments = mock_attachments
        
        # Mock recipients
        mock_email.Recipients = []
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = mock_email
        
        self.adapter.connect()
        
        # Test email retrieval - now returns EmailData object
        result = self.adapter.get_email_by_id("test_email_id")
        assert result.id == "test_email_id"
        assert result.subject == "Test Email Subject"
        assert result.sender == "Test Sender"
        assert result.sender_email == "test@example.com"
        mock_namespace.GetItemFromID.assert_called_once_with("test_email_id")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_not_found(self, mock_client, mock_pythoncom):
        """Test get_email_by_id when email is not found."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = None  # Email not found
        
        self.adapter.connect()
        
        # Test email not found
        with pytest.raises(EmailNotFoundError):
            self.adapter.get_email_by_id("nonexistent_id")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_detailed_success(self, mock_client, mock_pythoncom):
        """Test successful detailed email retrieval by ID."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create detailed mock email
        mock_email = Mock()
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "test_email_id"
        mock_email.Subject = "Test Email Subject"
        mock_email.SenderName = "John Doe"
        mock_email.SenderEmailAddress = "john.doe@example.com"
        mock_email.Body = "This is the email body content."
        mock_email.HTMLBody = "<html><body>This is the email body content.</body></html>"
        mock_email.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email.SentOn = datetime(2023, 1, 15, 10, 25, 0)
        mock_email.UnRead = False
        mock_email.Importance = 1  # Normal
        mock_email.Size = 1024
        
        # Mock parent folder
        mock_parent = Mock()
        mock_parent.Name = "Inbox"
        mock_email.Parent = mock_parent
        
        # Mock attachments
        mock_attachments = Mock()
        mock_attachments.Count = 2
        mock_email.Attachments = mock_attachments
        
        # Mock recipients
        mock_recipient1 = Mock()
        mock_recipient1.Address = "recipient1@example.com"
        mock_recipient1.Name = "Recipient One"
        mock_recipient1.Type = 1  # To
        
        mock_recipient2 = Mock()
        mock_recipient2.Address = "cc@example.com"
        mock_recipient2.Name = "CC Recipient"
        mock_recipient2.Type = 2  # CC
        
        mock_recipients = [mock_recipient1, mock_recipient2]
        mock_email.Recipients = mock_recipients
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = mock_email
        
        self.adapter.connect()
        
        # Test detailed email retrieval
        result = self.adapter.get_email_by_id("test_email_id")
        
        # Assertions
        assert result.id == "test_email_id"
        assert result.subject == "Test Email Subject"
        assert result.sender == "John Doe"
        assert result.sender_email == "john.doe@example.com"
        assert result.body == "This is the email body content."
        assert result.body_html == "<html><body>This is the email body content.</body></html>"
        assert result.received_time == datetime(2023, 1, 15, 10, 30, 0)
        assert result.sent_time == datetime(2023, 1, 15, 10, 25, 0)
        assert result.is_read is True  # UnRead = False means is_read = True
        assert result.has_attachments is True
        assert result.importance == "Normal"
        assert result.folder_name == "Inbox"
        assert result.size == 1024
        assert "recipient1@example.com" in result.recipients
        assert "cc@example.com" in result.cc_recipients
        
        mock_namespace.GetItemFromID.assert_called_once_with("test_email_id")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_invalid_item_type(self, mock_client, mock_pythoncom):
        """Test get_email_by_id when item is not a mail item."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create mock item that's not a mail item
        mock_item = Mock()
        mock_item.Class = 40  # Not olMail (43)
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = mock_item
        
        self.adapter.connect()
        
        # Test with non-mail item
        with pytest.raises(EmailNotFoundError):
            self.adapter.get_email_by_id("not_mail_item_id")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_permission_error(self, mock_client, mock_pythoncom):
        """Test get_email_by_id when access is denied."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.side_effect = Exception("Access denied to email")
        
        self.adapter.connect()
        
        # Test permission error
        with pytest.raises(PermissionError) as exc_info:
            self.adapter.get_email_by_id("restricted_email_id")
        
        assert "Access denied" in str(exc_info.value)
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_with_missing_properties(self, mock_client, mock_pythoncom):
        """Test get_email_by_id with email missing some properties."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create mock email with missing properties
        mock_email = Mock()
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "minimal_email_id"
        mock_email.Subject = ""  # Empty subject
        # Missing SenderName, SenderEmailAddress, Body, etc.
        del mock_email.SenderName
        del mock_email.SenderEmailAddress
        del mock_email.Body
        del mock_email.HTMLBody
        del mock_email.ReceivedTime
        del mock_email.SentOn
        del mock_email.CreationTime
        del mock_email.UnRead
        del mock_email.Importance
        del mock_email.Size
        del mock_email.Parent
        del mock_email.Attachments
        del mock_email.Recipients
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = mock_email
        
        self.adapter.connect()
        
        # Test with minimal email properties
        result = self.adapter.get_email_by_id("minimal_email_id")
        
        # Should handle missing properties gracefully
        assert result.id == "minimal_email_id"
        assert result.subject == "(No Subject)"  # Default for empty subject
        assert result.sender == "unknown"  # Default when missing
        assert result.sender_email == "unknown@unknown.com"
        assert result.body == ""
        assert result.body_html == ""
        assert result.received_time is None
        assert result.sent_time is None
        assert result.is_read is False  # Default
        assert result.has_attachments is False
        assert result.importance == "Normal"
        assert result.folder_name == "Unknown"
        assert result.size == 0
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_email_by_id_with_html_only_body(self, mock_client, mock_pythoncom):
        """Test get_email_by_id with HTML-only email body."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create mock email with HTML body only
        mock_email = Mock()
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "html_email_id"
        mock_email.Subject = "HTML Email"
        mock_email.SenderName = "HTML Sender"
        mock_email.SenderEmailAddress = "html@example.com"
        mock_email.Body = ""  # No plain text
        mock_email.HTMLBody = "<html><body><p>This is <b>HTML</b> content.</p></body></html>"
        mock_email.UnRead = True
        mock_email.Importance = 2  # High
        mock_email.Size = 512
        
        # Mock empty collections
        mock_email.Recipients = []
        mock_attachments = Mock()
        mock_attachments.Count = 0
        mock_email.Attachments = mock_attachments
        
        mock_parent = Mock()
        mock_parent.Name = "Test Folder"
        mock_email.Parent = mock_parent
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.GetItemFromID.return_value = mock_email
        
        self.adapter.connect()
        
        # Test HTML-only email
        result = self.adapter.get_email_by_id("html_email_id")
        
        # Should extract text from HTML for plain text body
        assert result.id == "html_email_id"
        assert result.subject == "HTML Email"
        assert result.body == "This is HTML content."  # Extracted from HTML
        assert result.body_html == "<html><body><p>This is <b>HTML</b> content.</p></body></html>"
        assert result.is_read is False  # UnRead = True means is_read = False
        assert result.importance == "High"
        assert result.folder_name == "Test Folder"
    
    def test_get_email_by_id_invalid_email_id_format(self):
        """Test get_email_by_id with invalid email ID format."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        self.adapter._namespace = Mock()
        
        # Test with invalid ID format (too long)
        invalid_id = "x" * 300  # Exceeds 255 character limit
        with pytest.raises(EmailNotFoundError):
            self.adapter.get_email_by_id(invalid_id)
    
    def test_search_emails_when_not_connected(self):
        """Test search_emails raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.search_emails("test query")
    
    def test_search_emails_invalid_query(self):
        """Test search_emails with invalid query."""
        # Set up connected state
        self.adapter._connected = True
        self.adapter._outlook_app = Mock()
        self.adapter._namespace = Mock()
        
        # Test with None query
        result = self.adapter.search_emails(None)
        assert result == []
        
        # Test with empty query
        result = self.adapter.search_emails("")
        assert result == []
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_success(self, mock_client, mock_pythoncom):
        """Test successful email search."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        
        # Create mock email objects with proper properties
        mock_email1 = Mock()
        mock_email1.Class = 43  # olMail
        mock_email1.EntryID = "email1_id"
        mock_email1.Subject = "Test Email 1"
        mock_email1.SenderName = "Sender 1"
        mock_email1.SenderEmailAddress = "sender1@example.com"
        mock_email1.Body = "Test body 1"
        mock_email1.HTMLBody = "<p>Test body 1</p>"
        mock_email1.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email1.SentOn = datetime(2023, 1, 15, 10, 25, 0)
        mock_email1.UnRead = False
        mock_email1.Importance = 1
        mock_email1.Size = 1024
        mock_email1.Recipients = []
        mock_attachments1 = Mock()
        mock_attachments1.Count = 0
        mock_email1.Attachments = mock_attachments1
        mock_parent1 = Mock()
        mock_parent1.Name = "Inbox"
        mock_email1.Parent = mock_parent1
        
        mock_email2 = Mock()
        mock_email2.Class = 43  # olMail
        mock_email2.EntryID = "email2_id"
        mock_email2.Subject = "Test Email 2"
        mock_email2.SenderName = "Sender 2"
        mock_email2.SenderEmailAddress = "sender2@example.com"
        mock_email2.Body = "Test body 2"
        mock_email2.HTMLBody = "<p>Test body 2</p>"
        mock_email2.ReceivedTime = datetime(2023, 1, 15, 11, 30, 0)
        mock_email2.SentOn = datetime(2023, 1, 15, 11, 25, 0)
        mock_email2.UnRead = True
        mock_email2.Importance = 1
        mock_email2.Size = 2048
        mock_email2.Recipients = []
        mock_attachments2 = Mock()
        mock_attachments2.Count = 0
        mock_email2.Attachments = mock_attachments2
        mock_parent2 = Mock()
        mock_parent2.Name = "Inbox"
        mock_email2.Parent = mock_parent2
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_inbox.Items = mock_items
        
        # Mock search results
        mock_items.Sort = Mock()
        mock_items.Find.return_value = mock_email1
        mock_items.FindNext.side_effect = [mock_email2, None]  # Two results, then None
        
        self.adapter.connect()
        
        # Test email search - should search in default folders since get_folders will fail
        result = self.adapter.search_emails("test query")
        
        # Should find results from multiple default folders (Inbox, Sent Items, Drafts)
        # Each folder returns 2 results, but they're sorted by received time (newest first)
        assert len(result) == 4  # 2 results from each of 2 folders that return results
        
        # Results should be sorted by received time (newest first)
        assert result[0].received_time >= result[1].received_time
        
        # Should contain our test emails
        email_ids = [email.id for email in result]
        assert "email1_id" in email_ids
        assert "email2_id" in email_ids
        
        # Verify the processed query was used (should wrap in subject/body search)
        expected_query = '(subject:"test query" OR body:"test query")'
        # Find should be called multiple times (once per default folder)
        assert mock_items.Find.call_count >= 1
        mock_items.Find.assert_called_with(expected_query)
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_with_specific_folder(self, mock_client, mock_pythoncom):
        """Test email search in specific folder."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_folder  # Return the folder for "Inbox"
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = None  # No results
        mock_items.FindNext.return_value = None
        
        self.adapter.connect()
        
        # Test search in specific folder by name
        result = self.adapter.search_emails("test query", "Inbox")
        assert result == []
        
        # Verify the processed query was used
        expected_query = '(subject:"test query" OR body:"test query")'
        mock_items.Find.assert_called_once_with(expected_query)
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_context_manager(self, mock_client, mock_pythoncom):
        """Test OutlookAdapter as context manager."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Test context manager
        with OutlookAdapter() as adapter:
            assert adapter.is_connected() is True
        
        # After context, should be disconnected
        assert adapter.is_connected() is False

    # Enhanced Search Functionality Tests
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_enhanced_with_folder_and_limit(self, mock_client, mock_pythoncom):
        """Test enhanced search_emails with folder and limit parameters."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        # Create mock search results
        mock_email1 = Mock()
        mock_email1.Class = 43  # olMail
        mock_email1.EntryID = "search_result_1"
        mock_email1.Subject = "Test Search Result 1"
        mock_email1.SenderName = "Sender One"
        mock_email1.SenderEmailAddress = "sender1@example.com"
        mock_email1.Body = "This email contains the search term"
        mock_email1.HTMLBody = "<p>This email contains the search term</p>"
        mock_email1.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email1.SentOn = datetime(2023, 1, 15, 10, 25, 0)
        mock_email1.UnRead = False
        mock_email1.Importance = 1
        mock_email1.Size = 1024
        mock_email1.Recipients = []
        mock_attachments1 = Mock()
        mock_attachments1.Count = 0
        mock_email1.Attachments = mock_attachments1
        mock_parent1 = Mock()
        mock_parent1.Name = "Test Folder"
        mock_email1.Parent = mock_parent1
        
        mock_email2 = Mock()
        mock_email2.Class = 43  # olMail
        mock_email2.EntryID = "search_result_2"
        mock_email2.Subject = "Test Search Result 2"
        mock_email2.SenderName = "Sender Two"
        mock_email2.SenderEmailAddress = "sender2@example.com"
        mock_email2.Body = "Another email with search term"
        mock_email2.HTMLBody = "<p>Another email with search term</p>"
        mock_email2.ReceivedTime = datetime(2023, 1, 14, 15, 20, 0)
        mock_email2.SentOn = datetime(2023, 1, 14, 15, 15, 0)
        mock_email2.UnRead = True
        mock_email2.Importance = 2
        mock_email2.Size = 2048
        mock_email2.Recipients = []
        mock_attachments2 = Mock()
        mock_attachments2.Count = 1
        mock_email2.Attachments = mock_attachments2
        mock_parent2 = Mock()
        mock_parent2.Name = "Test Folder"
        mock_email2.Parent = mock_parent2
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock folder search
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = mock_email1
        mock_items.FindNext.side_effect = [mock_email2, None]  # Two results, then None
        
        # Mock get_folder_by_name to return our test folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test enhanced search with folder and limit
        results = self.adapter.search_emails("search term", folder_name="Test Folder", limit=2)
        
        # Assertions
        assert len(results) == 2
        assert results[0].id == "search_result_1"
        assert results[0].subject == "Test Search Result 1"
        assert results[0].sender == "Sender One"
        assert results[0].folder_name == "Test Folder"
        assert results[1].id == "search_result_2"
        assert results[1].subject == "Test Search Result 2"
        assert results[1].sender == "Sender Two"
        assert results[1].folder_name == "Test Folder"
        
        # Verify search was called with processed query
        mock_items.Find.assert_called_once()
        mock_items.Sort.assert_called_once_with("[ReceivedTime]", True)
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_global_search(self, mock_client, mock_pythoncom):
        """Test search_emails with global search across all folders."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folders to return test folders
        from src.outlook_mcp_server.models.folder_data import FolderData
        test_folders = [
            FolderData(
                id="folder1_id",
                name="Inbox",
                full_path="Inbox",
                item_count=10,
                unread_count=2,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="folder2_id",
                name="Sent Items",
                full_path="Sent Items",
                item_count=5,
                unread_count=0,
                parent_folder="",
                folder_type="Mail"
            )
        ]
        
        # Mock the get_folders method
        self.adapter.get_folders = Mock(return_value=test_folders)
        
        # Mock search results for each folder
        mock_email1 = Mock()
        mock_email1.Class = 43
        mock_email1.EntryID = "global_result_1"
        mock_email1.Subject = "Global Search Result 1"
        mock_email1.SenderName = "Global Sender"
        mock_email1.SenderEmailAddress = "global@example.com"
        mock_email1.Body = "Global search content"
        mock_email1.HTMLBody = "<p>Global search content</p>"
        mock_email1.ReceivedTime = datetime(2023, 1, 16, 12, 0, 0)
        mock_email1.SentOn = datetime(2023, 1, 16, 11, 55, 0)
        mock_email1.UnRead = False
        mock_email1.Importance = 1
        mock_email1.Size = 512
        mock_email1.Recipients = []
        mock_attachments1 = Mock()
        mock_attachments1.Count = 0
        mock_email1.Attachments = mock_attachments1
        mock_parent1 = Mock()
        mock_parent1.Name = "Inbox"
        mock_email1.Parent = mock_parent1
        
        # Mock _search_in_folder to return results
        def mock_search_in_folder(query, folder_name, limit):
            if folder_name == "Inbox":
                return [self.adapter._transform_email_to_data(mock_email1, "Inbox")]
            return []
        
        self.adapter._search_in_folder = Mock(side_effect=mock_search_in_folder)
        
        self.adapter.connect()
        
        # Test global search (no folder specified)
        results = self.adapter.search_emails("global search", limit=10)
        
        # Assertions
        assert len(results) == 1
        assert results[0].subject == "Global Search Result 1"
        assert results[0].sender == "Global Sender"
        
        # Verify get_folders was called
        self.adapter.get_folders.assert_called_once()
        
        # Verify search was attempted in folders
        assert self.adapter._search_in_folder.call_count >= 1
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_empty_results(self, mock_client, mock_pythoncom):
        """Test search_emails with no matching results."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock folder with no search results
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = None  # No results
        mock_items.FindNext.return_value = None
        
        # Mock get_folder_by_name to return our test folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test search with no results
        results = self.adapter.search_emails("nonexistent term", folder_name="Test Folder")
        
        # Assertions
        assert results == []
        mock_items.Find.assert_called_once()
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_invalid_query_types(self, mock_client, mock_pythoncom):
        """Test search_emails with various invalid query types."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock the Folders collection to prevent iteration errors
        mock_namespace.Folders = []
        
        self.adapter.connect()
        
        # Test with None query - should return empty list immediately
        results = self.adapter.search_emails(None)
        assert results == []
        
        # Test with empty string query - should return empty list immediately
        results = self.adapter.search_emails("")
        assert results == []
        
        # Test with whitespace-only query - should return empty list immediately
        results = self.adapter.search_emails("   ")
        assert results == []
        
        # Test with non-string query - should return empty list immediately
        results = self.adapter.search_emails(123)
        assert results == []
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_folder_not_found(self, mock_client, mock_pythoncom):
        """Test search_emails with non-existent folder."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to raise FolderNotFoundError
        def mock_get_folder_by_name(name):
            raise FolderNotFoundError(name)
        
        self.adapter.get_folder_by_name = Mock(side_effect=mock_get_folder_by_name)
        
        self.adapter.connect()
        
        # Test search in non-existent folder
        with pytest.raises(FolderNotFoundError):
            self.adapter.search_emails("test query", folder_name="NonExistentFolder")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_permission_error(self, mock_client, mock_pythoncom):
        """Test search_emails with permission denied error."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to raise PermissionError
        def mock_get_folder_by_name(name):
            raise PermissionError(name, f"Access denied to folder '{name}'")
        
        self.adapter.get_folder_by_name = Mock(side_effect=mock_get_folder_by_name)
        
        self.adapter.connect()
        
        # Test search with permission error
        with pytest.raises(PermissionError):
            self.adapter.search_emails("test query", folder_name="RestrictedFolder")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_limit_enforcement(self, mock_client, mock_pythoncom):
        """Test search_emails properly enforces result limits."""
        from datetime import datetime
        
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        # Create multiple mock search results
        mock_emails = []
        for i in range(5):
            mock_email = Mock()
            mock_email.Class = 43  # olMail
            mock_email.EntryID = f"limit_test_{i}"
            mock_email.Subject = f"Limit Test Email {i}"
            mock_email.SenderName = f"Sender {i}"
            mock_email.SenderEmailAddress = f"sender{i}@example.com"
            mock_email.Body = f"Email content {i}"
            mock_email.HTMLBody = f"<p>Email content {i}</p>"
            mock_email.ReceivedTime = datetime(2023, 1, 15 + i, 10, 30, 0)
            mock_email.SentOn = datetime(2023, 1, 15 + i, 10, 25, 0)
            mock_email.UnRead = False
            mock_email.Importance = 1
            mock_email.Size = 1024
            mock_email.Recipients = []
            mock_attachments = Mock()
            mock_attachments.Count = 0
            mock_email.Attachments = mock_attachments
            mock_parent = Mock()
            mock_parent.Name = "Test Folder"
            mock_email.Parent = mock_parent
            mock_emails.append(mock_email)
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock folder search with multiple results
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = mock_emails[0]
        mock_items.FindNext.side_effect = mock_emails[1:] + [None]  # Return all emails, then None
        
        # Mock get_folder_by_name to return our test folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test search with limit of 3
        results = self.adapter.search_emails("limit test", folder_name="Test Folder", limit=3)
        
        # Assertions - should only return 3 results despite 5 being available
        assert len(results) == 3
        assert results[0].subject == "Limit Test Email 0"
        assert results[1].subject == "Limit Test Email 1"
        assert results[2].subject == "Limit Test Email 2"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_query_processing(self, mock_client, mock_pythoncom):
        """Test search_emails query processing functionality."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock folder with no results (we're testing query processing, not results)
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = None
        mock_items.FindNext.return_value = None
        
        # Mock get_folder_by_name to return our test folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test simple query processing
        self.adapter.search_emails("simple query", folder_name="Test Folder")
        
        # Verify that Find was called (query was processed)
        mock_items.Find.assert_called_once()
        
        # Test query with Outlook syntax (should not be modified)
        mock_items.Find.reset_mock()
        self.adapter.search_emails("subject:test query", folder_name="Test Folder")
        mock_items.Find.assert_called_once()
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_search_emails_non_mail_items_filtered(self, mock_client, mock_pythoncom):
        """Test search_emails filters out non-mail items."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        
        # Create mock items - one mail item and one non-mail item
        mock_mail_item = Mock()
        mock_mail_item.Class = 43  # olMail
        mock_mail_item.EntryID = "mail_item_id"
        mock_mail_item.Subject = "Mail Item"
        mock_mail_item.SenderName = "Mail Sender"
        mock_mail_item.SenderEmailAddress = "mail@example.com"
        mock_mail_item.Body = "Mail content"
        mock_mail_item.HTMLBody = "<p>Mail content</p>"
        mock_mail_item.UnRead = False
        mock_mail_item.Importance = 1
        mock_mail_item.Size = 1024
        mock_mail_item.Recipients = []
        mock_attachments = Mock()
        mock_attachments.Count = 0
        mock_mail_item.Attachments = mock_attachments
        mock_parent = Mock()
        mock_parent.Name = "Test Folder"
        mock_mail_item.Parent = mock_parent
        
        mock_non_mail_item = Mock()
        mock_non_mail_item.Class = 40  # Not olMail
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock folder search returning both items
        mock_folder.Items = mock_items
        mock_items.Sort = Mock()
        mock_items.Find.return_value = mock_mail_item
        mock_items.FindNext.side_effect = [mock_non_mail_item, None]  # Non-mail item, then None
        
        # Mock get_folder_by_name to return our test folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test search - should only return the mail item
        results = self.adapter.search_emails("test query", folder_name="Test Folder")
        
        # Assertions - should only return 1 result (the mail item)
        assert len(results) == 1
        assert results[0].subject == "Mail Item"
        assert results[0].sender == "Mail Sender"
    
    def test_process_search_query(self):
        """Test _process_search_query method."""
        # Test simple query processing
        result = self.adapter._process_search_query("simple query")
        assert 'subject:"simple query"' in result or 'body:"simple query"' in result
        
        # Test query with existing Outlook syntax
        result = self.adapter._process_search_query("subject:test")
        assert result == "subject:test"
        
        # Test query with body syntax
        result = self.adapter._process_search_query("body:content")
        assert result == "body:content"
        
        # Test query with from syntax
        result = self.adapter._process_search_query("from:sender@example.com")
        assert result == "from:sender@example.com"


class TestOutlookAdapterFolderOperations:
    """Test cases for OutlookAdapter folder operations."""
    
    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.adapter = OutlookAdapter()
    
    def teardown_method(self):
        """Clean up after each test method."""
        if hasattr(self.adapter, 'disconnect'):
            try:
                self.adapter.disconnect()
            except:
                pass
    
    def test_get_folders_when_not_connected(self):
        """Test get_folders raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.get_folders()
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folders_success(self, mock_client, mock_pythoncom):
        """Test successful folder retrieval."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create mock folders
        mock_root_folder = Mock()
        mock_root_folder.Name = "TestAccount"
        mock_root_folder.EntryID = "root_id"
        mock_root_folder.Items.Count = 0
        mock_root_folder.UnReadItemCount = 0
        mock_root_folder.DefaultItemType = 0  # Mail
        mock_root_folder.Folders = []
        
        mock_subfolder = Mock()
        mock_subfolder.Name = "Inbox"
        mock_subfolder.EntryID = "inbox_id"
        mock_subfolder.Items.Count = 10
        mock_subfolder.UnReadItemCount = 3
        mock_subfolder.DefaultItemType = 0  # Mail
        mock_subfolder.Folders = []
        
        mock_root_folder.Folders = [mock_subfolder]
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.Folders = [mock_root_folder]
        
        self.adapter.connect()
        
        # Test get_folders
        result = self.adapter.get_folders()
        
        # Assertions
        assert len(result) == 2  # Root folder + subfolder
        assert result[0].name == "TestAccount"
        assert result[0].id == "root_id"
        assert result[0].full_path == "TestAccount"
        assert result[0].parent_folder == ""
        assert result[1].name == "Inbox"
        assert result[1].id == "inbox_id"
        assert result[1].full_path == "TestAccount/Inbox"
        assert result[1].parent_folder == "TestAccount"
        assert result[1].item_count == 10
        assert result[1].unread_count == 3
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folders_with_nested_hierarchy(self, mock_client, mock_pythoncom):
        """Test folder retrieval with nested folder hierarchy."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create nested folder structure
        mock_root = Mock()
        mock_root.Name = "Root"
        mock_root.EntryID = "root_id"
        mock_root.Items.Count = 0
        mock_root.UnReadItemCount = 0
        mock_root.DefaultItemType = 0
        
        mock_level1 = Mock()
        mock_level1.Name = "Level1"
        mock_level1.EntryID = "level1_id"
        mock_level1.Items.Count = 5
        mock_level1.UnReadItemCount = 2
        mock_level1.DefaultItemType = 0
        
        mock_level2 = Mock()
        mock_level2.Name = "Level2"
        mock_level2.EntryID = "level2_id"
        mock_level2.Items.Count = 3
        mock_level2.UnReadItemCount = 1
        mock_level2.DefaultItemType = 0
        mock_level2.Folders = []
        
        mock_level1.Folders = [mock_level2]
        mock_root.Folders = [mock_level1]
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.Folders = [mock_root]
        
        self.adapter.connect()
        
        # Test get_folders
        result = self.adapter.get_folders()
        
        # Assertions
        assert len(result) == 3  # Root + Level1 + Level2
        assert result[0].name == "Root"
        assert result[0].full_path == "Root"
        assert result[1].name == "Level1"
        assert result[1].full_path == "Root/Level1"
        assert result[2].name == "Level2"
        assert result[2].full_path == "Root/Level1/Level2"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folders_with_access_error(self, mock_client, mock_pythoncom):
        """Test get_folders when some folders have access errors."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        # Create mock folder that raises access error
        mock_folder = Mock()
        mock_folder.Name = "RestrictedFolder"
        mock_folder.EntryID = "restricted_id"
        mock_folder.Items.Count = Mock(side_effect=Exception("Access denied"))
        mock_folder.UnReadItemCount = 0
        mock_folder.DefaultItemType = 0
        mock_folder.Folders = []
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.Folders = [mock_folder]
        
        self.adapter.connect()
        
        # Test get_folders - should handle access errors gracefully
        result = self.adapter.get_folders()
        
        # Should still return folder data with default values
        assert len(result) == 1
        assert result[0].name == "RestrictedFolder"
        assert result[0].item_count == 0  # Default value when access fails
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_get_folders_permission_error(self, mock_client, mock_pythoncom):
        """Test get_folders when permission error occurs."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock Folders property to raise exception when accessed
        type(mock_namespace).Folders = PropertyMock(side_effect=Exception("Access denied to folders"))
        
        self.adapter.connect()
        
        # Test get_folders with permission error
        with pytest.raises(PermissionError) as exc_info:
            self.adapter.get_folders()
        
        assert "Access denied to folders" in str(exc_info.value)
    
    def test_transform_folder_to_data_basic(self):
        """Test folder transformation with basic properties."""
        # Create mock folder
        mock_folder = Mock()
        mock_folder.Name = "TestFolder"
        mock_folder.EntryID = "test_id"
        mock_folder.Items.Count = 15
        mock_folder.UnReadItemCount = 5
        mock_folder.DefaultItemType = 0  # Mail
        
        # Test transformation
        result = self.adapter._transform_folder_to_data(mock_folder, "Parent")
        
        # Assertions
        assert result.name == "TestFolder"
        assert result.id == "test_id"
        assert result.full_path == "Parent/TestFolder"
        assert result.item_count == 15
        assert result.unread_count == 5
        assert result.parent_folder == "Parent"
        assert result.folder_type == "Mail"
    
    def test_transform_folder_to_data_root_folder(self):
        """Test folder transformation for root folder."""
        # Create mock folder
        mock_folder = Mock()
        mock_folder.Name = "RootFolder"
        mock_folder.EntryID = "root_id"
        mock_folder.Items.Count = 0
        mock_folder.UnReadItemCount = 0
        mock_folder.DefaultItemType = 1  # Contact
        
        # Test transformation with empty parent path
        result = self.adapter._transform_folder_to_data(mock_folder, "")
        
        # Assertions
        assert result.name == "RootFolder"
        assert result.full_path == "RootFolder"
        assert result.parent_folder == ""
        assert result.folder_type == "Contact"
    
    def test_transform_folder_to_data_missing_properties(self):
        """Test folder transformation when some properties are missing."""
        # Create mock folder with missing properties
        mock_folder = Mock()
        mock_folder.Name = "IncompleteFolder"
        # Missing EntryID, Items, UnReadItemCount, DefaultItemType
        del mock_folder.EntryID
        del mock_folder.Items
        del mock_folder.UnReadItemCount
        del mock_folder.DefaultItemType
        
        # Test transformation
        result = self.adapter._transform_folder_to_data(mock_folder, "Parent")
        
        # Should handle missing properties gracefully
        assert result.name == "IncompleteFolder"
        assert result.id.startswith("unknown_IncompleteFolder_")  # Generated ID when missing
        assert result.item_count == 0  # Default when missing
        assert result.unread_count == 0  # Default when missing
        assert result.folder_type == "Mail"  # Default when missing
    
    def test_get_folder_type_by_item_type(self):
        """Test folder type determination by DefaultItemType."""
        # Test different item types
        test_cases = [
            (0, "Mail"),
            (1, "Contact"),
            (2, "Task"),
            (3, "Journal"),
            (4, "Note"),
            (5, "Post"),
            (99, "Mail")  # Unknown type defaults to Mail
        ]
        
        for item_type, expected_type in test_cases:
            mock_folder = Mock()
            mock_folder.DefaultItemType = item_type
            
            result = self.adapter._get_folder_type(mock_folder)
            assert result == expected_type
    
    def test_get_folder_type_by_name(self):
        """Test folder type determination by folder name."""
        # Test different folder names
        test_cases = [
            ("Contacts", "Contact"),
            ("My Contacts", "Contact"),
            ("Calendar", "Calendar"),
            ("Tasks", "Task"),
            ("My Tasks", "Task"),
            ("Notes", "Note"),
            ("Inbox", "Mail"),  # Default
            ("Unknown Folder", "Mail")  # Default
        ]
        
        for folder_name, expected_type in test_cases:
            mock_folder = Mock()
            mock_folder.Name = folder_name
            # No DefaultItemType property
            del mock_folder.DefaultItemType
            
            result = self.adapter._get_folder_type(mock_folder)
            assert result == expected_type
    
    def test_validate_folder_access_when_not_connected(self):
        """Test validate_folder_access when not connected."""
        result = self.adapter.validate_folder_access("Inbox")
        assert result is False


class TestOutlookAdapterEmailListing:
    """Test cases for OutlookAdapter email listing functionality."""
    
    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.adapter = OutlookAdapter()
    
    def teardown_method(self):
        """Clean up after each test method."""
        if hasattr(self.adapter, 'disconnect'):
            try:
                self.adapter.disconnect()
            except:
                pass
    
    def test_list_emails_when_not_connected(self):
        """Test list_emails raises error when not connected."""
        with pytest.raises(OutlookConnectionError):
            self.adapter.list_emails()
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_default_inbox(self, mock_client, mock_pythoncom):
        """Test list_emails with default inbox folder."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        mock_email = Mock()
        
        # Configure email properties
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "test_email_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "sender@test.com"
        mock_email.UnRead = False
        mock_email.Body = "Test body"
        mock_email.HTMLBody = "<p>Test body</p>"
        mock_email.ReceivedTime = "2023-01-01 10:00:00"
        mock_email.SentOn = "2023-01-01 09:00:00"
        mock_email.Importance = 1  # Normal
        mock_email.Size = 1024
        mock_email.Recipients = []
        mock_email.Attachments.Count = 0
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([mock_email]))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test list_emails
        result = self.adapter.list_emails()
        
        # Assertions
        assert len(result) == 1
        assert result[0].id == "test_email_id"
        assert result[0].subject == "Test Subject"
        assert result[0].sender == "Test Sender"
        assert result[0].sender_email == "sender@test.com"
        assert result[0].is_read is True
        assert result[0].folder_name == "Inbox"
        mock_items.Sort.assert_called_once_with("[ReceivedTime]", True)
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_specific_folder(self, mock_client, mock_pythoncom):
        """Test list_emails with specific folder name."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        mock_items = Mock()
        mock_email = Mock()
        
        # Configure email properties
        mock_email.Class = 43  # olMail
        mock_email.EntryID = "test_email_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "sender@test.com"
        mock_email.UnRead = True
        mock_email.Body = "Test body"
        mock_email.HTMLBody = "<p>Test body</p>"
        mock_email.ReceivedTime = "2023-01-01 10:00:00"
        mock_email.SentOn = "2023-01-01 09:00:00"
        mock_email.Importance = 2  # High
        mock_email.Size = 2048
        mock_email.Recipients = []
        mock_email.Attachments.Count = 1
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([mock_email]))
        mock_folder.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to return specific folder
        self.adapter.get_folder_by_name = Mock(return_value=mock_folder)
        
        self.adapter.connect()
        
        # Test list_emails with specific folder
        result = self.adapter.list_emails(folder_name="Sent Items")
        
        # Assertions
        assert len(result) == 1
        assert result[0].id == "test_email_id"
        assert result[0].subject == "Test Subject"
        assert result[0].is_read is False  # UnRead = True means not read
        assert result[0].has_attachments is True
        assert result[0].importance == "High"
        assert result[0].folder_name == "Sent Items"
        self.adapter.get_folder_by_name.assert_called_once_with("Sent Items")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_unread_only_filter(self, mock_client, mock_pythoncom):
        """Test list_emails with unread_only filter."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        
        # Create read and unread emails
        mock_read_email = Mock()
        mock_read_email.Class = 43
        mock_read_email.UnRead = False
        mock_read_email.EntryID = "read_email_id"
        mock_read_email.Subject = "Read Email"
        mock_read_email.SenderName = "Sender"
        mock_read_email.SenderEmailAddress = "sender@test.com"
        
        mock_unread_email = Mock()
        mock_unread_email.Class = 43
        mock_unread_email.UnRead = True
        mock_unread_email.EntryID = "unread_email_id"
        mock_unread_email.Subject = "Unread Email"
        mock_unread_email.SenderName = "Sender"
        mock_unread_email.SenderEmailAddress = "sender@test.com"
        mock_unread_email.Body = ""
        mock_unread_email.HTMLBody = ""
        mock_unread_email.ReceivedTime = None
        mock_unread_email.SentOn = None
        mock_unread_email.Importance = 1
        mock_unread_email.Size = 0
        mock_unread_email.Recipients = []
        mock_unread_email.Attachments.Count = 0
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([mock_read_email, mock_unread_email]))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test list_emails with unread_only=True
        result = self.adapter.list_emails(unread_only=True)
        
        # Assertions - should only return unread email
        assert len(result) == 1
        assert result[0].id == "unread_email_id"
        assert result[0].subject == "Unread Email"
        assert result[0].is_read is False
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_with_limit(self, mock_client, mock_pythoncom):
        """Test list_emails with limit parameter."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        
        # Create multiple emails
        emails = []
        for i in range(5):
            mock_email = Mock()
            mock_email.Class = 43
            mock_email.EntryID = f"email_id_{i}"
            mock_email.Subject = f"Email {i}"
            mock_email.SenderName = "Sender"
            mock_email.SenderEmailAddress = "sender@test.com"
            mock_email.UnRead = False
            mock_email.Body = ""
            mock_email.HTMLBody = ""
            mock_email.ReceivedTime = None
            mock_email.SentOn = None
            mock_email.Importance = 1
            mock_email.Size = 0
            mock_email.Recipients = []
            mock_email.Attachments.Count = 0
            emails.append(mock_email)
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter(emails))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test list_emails with limit=3
        result = self.adapter.list_emails(limit=3)
        
        # Assertions - should only return 3 emails
        assert len(result) == 3
        assert result[0].id == "email_id_0"
        assert result[1].id == "email_id_1"
        assert result[2].id == "email_id_2"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_with_recipients(self, mock_client, mock_pythoncom):
        """Test list_emails with email containing recipients."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        mock_email = Mock()
        
        # Create mock recipients
        mock_to_recipient = Mock()
        mock_to_recipient.Address = "to@test.com"
        mock_to_recipient.Type = 1  # To
        
        mock_cc_recipient = Mock()
        mock_cc_recipient.Address = "cc@test.com"
        mock_cc_recipient.Type = 2  # CC
        
        mock_bcc_recipient = Mock()
        mock_bcc_recipient.Address = "bcc@test.com"
        mock_bcc_recipient.Type = 3  # BCC
        
        mock_recipients = [mock_to_recipient, mock_cc_recipient, mock_bcc_recipient]
        
        # Configure email properties
        mock_email.Class = 43
        mock_email.EntryID = "test_email_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "sender@test.com"
        mock_email.UnRead = False
        mock_email.Body = "Test body"
        mock_email.HTMLBody = "<p>Test body</p>"
        mock_email.ReceivedTime = "2023-01-01 10:00:00"
        mock_email.SentOn = "2023-01-01 09:00:00"
        mock_email.Importance = 0  # Low
        mock_email.Size = 512
        mock_email.Recipients = mock_recipients
        mock_email.Attachments.Count = 0
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([mock_email]))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test list_emails
        result = self.adapter.list_emails()
        
        # Assertions
        assert len(result) == 1
        assert result[0].recipients == ["to@test.com"]
        assert result[0].cc_recipients == ["cc@test.com"]
        assert result[0].bcc_recipients == ["bcc@test.com"]
        assert result[0].importance == "Low"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_non_mail_items_filtered(self, mock_client, mock_pythoncom):
        """Test list_emails filters out non-mail items."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        
        # Create mail and non-mail items
        mock_mail_item = Mock()
        mock_mail_item.Class = 43  # olMail
        mock_mail_item.EntryID = "mail_id"
        mock_mail_item.Subject = "Mail Item"
        mock_mail_item.SenderName = "Sender"
        mock_mail_item.SenderEmailAddress = "sender@test.com"
        mock_mail_item.UnRead = False
        mock_mail_item.Body = ""
        mock_mail_item.HTMLBody = ""
        mock_mail_item.ReceivedTime = None
        mock_mail_item.SentOn = None
        mock_mail_item.Importance = 1
        mock_mail_item.Size = 0
        mock_mail_item.Recipients = []
        mock_mail_item.Attachments.Count = 0
        
        mock_contact_item = Mock()
        mock_contact_item.Class = 40  # olContact
        
        mock_task_item = Mock()
        mock_task_item.Class = 48  # olTask
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([mock_contact_item, mock_mail_item, mock_task_item]))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test list_emails
        result = self.adapter.list_emails()
        
        # Assertions - should only return mail item
        assert len(result) == 1
        assert result[0].id == "mail_id"
        assert result[0].subject == "Mail Item"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_folder_not_found(self, mock_client, mock_pythoncom):
        """Test list_emails when specified folder is not found."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to raise FolderNotFoundError
        self.adapter.get_folder_by_name = Mock(side_effect=FolderNotFoundError("NonExistent"))
        
        self.adapter.connect()
        
        # Test list_emails with non-existent folder
        with pytest.raises(FolderNotFoundError):
            self.adapter.list_emails(folder_name="NonExistent")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_permission_error(self, mock_client, mock_pythoncom):
        """Test list_emails when permission error occurs."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to raise PermissionError
        self.adapter.get_folder_by_name = Mock(side_effect=PermissionError("Restricted", "Access denied"))
        
        self.adapter.connect()
        
        # Test list_emails with restricted folder
        with pytest.raises(PermissionError):
            self.adapter.list_emails(folder_name="Restricted")
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_list_emails_invalid_limit(self, mock_client, mock_pythoncom):
        """Test list_emails with invalid limit values."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_items = Mock()
        
        mock_items.Sort = Mock()
        mock_items.__iter__ = Mock(return_value=iter([]))
        mock_inbox.Items = mock_items
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        self.adapter.connect()
        
        # Test with negative limit - should use default
        result = self.adapter.list_emails(limit=-1)
        assert isinstance(result, list)
        
        # Test with zero limit - should use default
        result = self.adapter.list_emails(limit=0)
        assert isinstance(result, list)
    
    def test_transform_email_to_data_basic(self):
        """Test email transformation with basic properties."""
        # Create mock email
        mock_email = Mock()
        mock_email.EntryID = "test_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "sender@test.com"
        mock_email.Recipients = []
        mock_email.Body = "Test body"
        mock_email.HTMLBody = "<p>Test body</p>"
        mock_email.ReceivedTime = "2023-01-01 10:00:00"
        mock_email.SentOn = "2023-01-01 09:00:00"
        mock_email.UnRead = False
        mock_email.Attachments.Count = 1
        mock_email.Importance = 2  # High
        mock_email.Size = 2048
        
        # Test transformation
        result = self.adapter._transform_email_to_data(mock_email, "TestFolder")
        
        # Assertions
        assert result.id == "test_id"
        assert result.subject == "Test Subject"
        assert result.sender == "Test Sender"
        assert result.sender_email == "sender@test.com"
        assert result.body == "Test body"
        assert result.body_html == "<p>Test body</p>"
        assert result.is_read is True
        assert result.has_attachments is True
        assert result.importance == "High"
        assert result.folder_name == "TestFolder"
        assert result.size == 2048
    
    def test_transform_email_to_data_missing_properties(self):
        """Test email transformation when some properties are missing."""
        # Create mock email with minimal properties
        mock_email = Mock()
        mock_email.EntryID = "test_id"
        # Missing most properties
        del mock_email.Subject
        del mock_email.SenderName
        del mock_email.SenderEmailAddress
        del mock_email.Recipients
        del mock_email.Body
        del mock_email.HTMLBody
        del mock_email.ReceivedTime
        del mock_email.SentOn
        del mock_email.UnRead
        del mock_email.Attachments
        del mock_email.Importance
        del mock_email.Size
        
        # Test transformation
        result = self.adapter._transform_email_to_data(mock_email, "TestFolder")
        
        # Should handle missing properties gracefully
        assert result.id == "test_id"
        assert result.subject == "(No Subject)"
        assert result.sender == "unknown"
        assert result.sender_email == "unknown@unknown.com"
        assert result.body == ""
        assert result.body_html == ""
        assert result.is_read is False  # Default when UnRead is missing
        assert result.has_attachments is False
        assert result.importance == "Normal"
        assert result.folder_name == "TestFolder"
        assert result.size == 0
    
    def test_transform_email_to_data_invalid_sender_email(self):
        """Test email transformation with invalid sender email."""
        # Create mock email with invalid sender email
        mock_email = Mock()
        mock_email.EntryID = "test_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = "Test Sender"
        mock_email.SenderEmailAddress = "invalid_email"  # No @ symbol
        mock_email.Recipients = []
        mock_email.Body = ""
        mock_email.HTMLBody = ""
        mock_email.ReceivedTime = None
        mock_email.SentOn = None
        mock_email.UnRead = True
        mock_email.Attachments.Count = 0
        mock_email.Importance = 1
        mock_email.Size = 0
        
        # Test transformation
        result = self.adapter._transform_email_to_data(mock_email, "TestFolder")
        
        # Should fix invalid email
        assert result.sender_email == "TestSender@unknown.com"
    
    def test_transform_email_to_data_no_sender_info(self):
        """Test email transformation with no sender information."""
        # Create mock email with no sender info
        mock_email = Mock()
        mock_email.EntryID = "test_id"
        mock_email.Subject = "Test Subject"
        mock_email.SenderName = ""
        mock_email.SenderEmailAddress = ""
        mock_email.Recipients = []
        mock_email.Body = ""
        mock_email.HTMLBody = ""
        mock_email.ReceivedTime = None
        mock_email.SentOn = None
        mock_email.UnRead = True
        mock_email.Attachments.Count = 0
        mock_email.Importance = 1
        mock_email.Size = 0
        
        # Test transformation
        result = self.adapter._transform_email_to_data(mock_email, "TestFolder")
        
        # Should provide defaults
        assert result.sender == "unknown"
        assert result.sender_email == "unknown@unknown.com"
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_validate_folder_access_success(self, mock_client, mock_pythoncom):
        """Test successful folder access validation."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        mock_folder = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_folder  # Return folder for both connection test and folder lookup
        
        self.adapter.connect()
        
        # Test folder access validation
        result = self.adapter.validate_folder_access("Inbox")
        assert result is True
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_validate_folder_access_folder_not_found(self, mock_client, mock_pythoncom):
        """Test folder access validation when folder is not found."""
        # Set up successful connection
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_namespace.Folders = []  # No folders
        
        self.adapter.connect()
        
        # Test folder access validation for non-existent folder
        result = self.adapter.validate_folder_access("NonExistentFolder")
        assert result is False
    
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.pythoncom')
    @patch('src.outlook_mcp_server.adapters.outlook_adapter.win32com.client')
    def test_validate_folder_access_permission_error(self, mock_client, mock_pythoncom):
        """Test folder access validation when permission error occurs."""
        # Set up successful connection but folder access fails
        mock_outlook_app = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_client.GetActiveObject.return_value = mock_outlook_app
        mock_outlook_app.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        
        # Mock get_folder_by_name to raise PermissionError
        with patch.object(self.adapter, 'get_folder_by_name', side_effect=PermissionError("test", "Access denied")):
            self.adapter.connect()
            
            result = self.adapter.validate_folder_access("RestrictedFolder")
            assert result is False


class TestOutlookAdapterEmailRetrieval:
    """Test cases for OutlookAdapter detailed email retrieval functionality."""
    
    def setup_method(self):
        """Set up test fixtures before each test method."""
        self.adapter = OutlookAdapter()
    
    def teardown_method(self):
        """Clean up after each test method."""
        if hasattr(self.adapter, 'disconnect'):
            try:
                self.adapter.disconnect()
            except:
                pass
    
    def test_get_email_property_success(self):
        """Test _get_email_property with valid property."""
        mock_email = Mock()
        mock_email.TestProperty = "test_value"
        
        result = self.adapter._get_email_property(mock_email, 'TestProperty', 'default')
        assert result == "test_value"
    
    def test_get_email_property_missing(self):
        """Test _get_email_property with missing property."""
        mock_email = Mock()
        del mock_email.TestProperty
        
        result = self.adapter._get_email_property(mock_email, 'TestProperty', 'default')
        assert result == "default"
    
    def test_get_email_property_none_value(self):
        """Test _get_email_property with None value."""
        mock_email = Mock()
        mock_email.TestProperty = None
        
        result = self.adapter._get_email_property(mock_email, 'TestProperty', 'default')
        assert result == "default"
    
    def test_get_email_property_exception(self):
        """Test _get_email_property when property access raises exception."""
        mock_email = Mock()
        type(mock_email).TestProperty = PropertyMock(side_effect=Exception("Access error"))
        
        result = self.adapter._get_email_property(mock_email, 'TestProperty', 'default')
        assert result == "default"
    
    def test_extract_recipients_success(self):
        """Test _extract_recipients with valid recipients."""
        mock_email = Mock()
        
        # Create mock recipients
        mock_to_recipient = Mock()
        mock_to_recipient.Address = "to@example.com"
        mock_to_recipient.Name = "To Recipient"
        mock_to_recipient.Type = 1  # To
        
        mock_cc_recipient = Mock()
        mock_cc_recipient.Address = "cc@example.com"
        mock_cc_recipient.Name = "CC Recipient"
        mock_cc_recipient.Type = 2  # CC
        
        mock_bcc_recipient = Mock()
        mock_bcc_recipient.Address = "bcc@example.com"
        mock_bcc_recipient.Name = "BCC Recipient"
        mock_bcc_recipient.Type = 3  # BCC
        
        mock_email.Recipients = [mock_to_recipient, mock_cc_recipient, mock_bcc_recipient]
        
        recipients, cc_recipients, bcc_recipients = self.adapter._extract_recipients(mock_email)
        
        assert recipients == ["to@example.com"]
        assert cc_recipients == ["cc@example.com"]
        assert bcc_recipients == ["bcc@example.com"]
    
    def test_extract_recipients_no_email_address(self):
        """Test _extract_recipients when recipient has no email address."""
        mock_email = Mock()
        
        # Create mock recipient with only name
        mock_recipient = Mock()
        mock_recipient.Address = ""
        mock_recipient.Name = "Name Only"
        mock_recipient.Type = 1  # To
        
        mock_email.Recipients = [mock_recipient]
        
        recipients, cc_recipients, bcc_recipients = self.adapter._extract_recipients(mock_email)
        
        assert recipients == ["Name Only"]  # Should use name when no email
        assert cc_recipients == []
        assert bcc_recipients == []
    
    def test_extract_recipients_no_recipients(self):
        """Test _extract_recipients when email has no recipients."""
        mock_email = Mock()
        mock_email.Recipients = []
        
        recipients, cc_recipients, bcc_recipients = self.adapter._extract_recipients(mock_email)
        
        assert recipients == []
        assert cc_recipients == []
        assert bcc_recipients == []
    
    def test_extract_recipients_exception(self):
        """Test _extract_recipients when accessing recipients raises exception."""
        mock_email = Mock()
        type(mock_email).Recipients = PropertyMock(side_effect=Exception("Access error"))
        
        recipients, cc_recipients, bcc_recipients = self.adapter._extract_recipients(mock_email)
        
        assert recipients == []
        assert cc_recipients == []
        assert bcc_recipients == []
    
    def test_extract_email_body_both_formats(self):
        """Test _extract_email_body with both plain text and HTML."""
        mock_email = Mock()
        mock_email.Body = "Plain text content"
        mock_email.HTMLBody = "<html><body>HTML content</body></html>"
        
        body, body_html = self.adapter._extract_email_body(mock_email)
        
        assert body == "Plain text content"
        assert body_html == "<html><body>HTML content</body></html>"
    
    def test_extract_email_body_html_only(self):
        """Test _extract_email_body with HTML only."""
        mock_email = Mock()
        mock_email.Body = ""
        mock_email.HTMLBody = "<html><body><p>HTML <b>content</b> only</p></body></html>"
        
        body, body_html = self.adapter._extract_email_body(mock_email)
        
        assert body == "HTML content only"  # Extracted from HTML
        assert body_html == "<html><body><p>HTML <b>content</b> only</p></body></html>"
    
    def test_extract_email_body_text_only(self):
        """Test _extract_email_body with plain text only."""
        mock_email = Mock()
        mock_email.Body = "Plain text only\nWith line breaks"
        mock_email.HTMLBody = ""
        
        body, body_html = self.adapter._extract_email_body(mock_email)
        
        assert body == "Plain text only With line breaks"  # Cleaned
        assert body_html == "<html><body>Plain text only<br>\nWith line breaks</body></html>"
    
    def test_extract_timestamps_success(self):
        """Test _extract_timestamps with valid timestamps."""
        from datetime import datetime
        
        mock_email = Mock()
        mock_email.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email.SentOn = datetime(2023, 1, 15, 10, 25, 0)
        
        received_time, sent_time = self.adapter._extract_timestamps(mock_email)
        
        assert received_time == datetime(2023, 1, 15, 10, 30, 0)
        assert sent_time == datetime(2023, 1, 15, 10, 25, 0)
    
    def test_extract_timestamps_fallback_to_creation_time(self):
        """Test _extract_timestamps fallback to creation time when sent time missing."""
        from datetime import datetime
        
        mock_email = Mock()
        mock_email.ReceivedTime = datetime(2023, 1, 15, 10, 30, 0)
        mock_email.CreationTime = datetime(2023, 1, 15, 10, 20, 0)
        del mock_email.SentOn  # Missing sent time
        
        received_time, sent_time = self.adapter._extract_timestamps(mock_email)
        
        assert received_time == datetime(2023, 1, 15, 10, 30, 0)
        assert sent_time == datetime(2023, 1, 15, 10, 20, 0)  # Uses creation time
    
    def test_extract_timestamps_missing(self):
        """Test _extract_timestamps when timestamps are missing."""
        mock_email = Mock()
        del mock_email.ReceivedTime
        del mock_email.SentOn
        del mock_email.CreationTime
        
        received_time, sent_time = self.adapter._extract_timestamps(mock_email)
        
        assert received_time is None
        assert sent_time is None
    
    def test_extract_attachment_info_with_attachments(self):
        """Test _extract_attachment_info with attachments present."""
        mock_email = Mock()
        mock_attachments = Mock()
        mock_attachments.Count = 3
        mock_email.Attachments = mock_attachments
        
        has_attachments, attachment_count = self.adapter._extract_attachment_info(mock_email)
        
        assert has_attachments is True
        assert attachment_count == 3
    
    def test_extract_attachment_info_no_attachments(self):
        """Test _extract_attachment_info with no attachments."""
        mock_email = Mock()
        mock_attachments = Mock()
        mock_attachments.Count = 0
        mock_email.Attachments = mock_attachments
        
        has_attachments, attachment_count = self.adapter._extract_attachment_info(mock_email)
        
        assert has_attachments is False
        assert attachment_count == 0
    
    def test_extract_attachment_info_missing_property(self):
        """Test _extract_attachment_info when Attachments property is missing."""
        mock_email = Mock()
        del mock_email.Attachments
        
        has_attachments, attachment_count = self.adapter._extract_attachment_info(mock_email)
        
        assert has_attachments is False
        assert attachment_count == 0
    
    def test_extract_importance_levels(self):
        """Test _extract_importance with different importance levels."""
        test_cases = [
            (0, "Low"),
            (1, "Normal"),
            (2, "High"),
            (99, "Normal")  # Unknown value defaults to Normal
        ]
        
        for importance_value, expected in test_cases:
            mock_email = Mock()
            mock_email.Importance = importance_value
            
            result = self.adapter._extract_importance(mock_email)
            assert result == expected
    
    def test_extract_importance_missing(self):
        """Test _extract_importance when Importance property is missing."""
        mock_email = Mock()
        del mock_email.Importance
        
        result = self.adapter._extract_importance(mock_email)
        assert result == "Normal"
    
    def test_extract_folder_name_success(self):
        """Test _extract_folder_name with valid parent folder."""
        mock_email = Mock()
        mock_parent = Mock()
        mock_parent.Name = "Test Folder"
        mock_email.Parent = mock_parent
        
        result = self.adapter._extract_folder_name(mock_email)
        assert result == "Test Folder"
    
    def test_extract_folder_name_no_parent(self):
        """Test _extract_folder_name when parent folder is missing."""
        mock_email = Mock()
        del mock_email.Parent
        
        result = self.adapter._extract_folder_name(mock_email)
        assert result == "Unknown"
    
    def test_validate_sender_info_valid(self):
        """Test _validate_sender_info with valid sender information."""
        sender_name, sender_email = self.adapter._validate_sender_info(
            "John Doe", "john.doe@example.com"
        )
        
        assert sender_name == "John Doe"
        assert sender_email == "john.doe@example.com"
    
    def test_validate_sender_info_invalid_email(self):
        """Test _validate_sender_info with invalid email."""
        sender_name, sender_email = self.adapter._validate_sender_info(
            "John Doe", "invalid-email"
        )
        
        assert sender_name == "John Doe"
        assert sender_email == "JohnDoe@unknown.com"  # Generated from name
    
    def test_validate_sender_info_missing_name(self):
        """Test _validate_sender_info with missing sender name."""
        sender_name, sender_email = self.adapter._validate_sender_info(
            "", "john.doe@example.com"
        )
        
        assert sender_name == "john.doe"  # Extracted from email
        assert sender_email == "john.doe@example.com"
    
    def test_validate_sender_info_both_invalid(self):
        """Test _validate_sender_info with both name and email invalid."""
        sender_name, sender_email = self.adapter._validate_sender_info("", "")
        
        assert sender_name == "unknown"
        assert sender_email == "unknown@unknown.com"
    
    def test_is_valid_email_format_valid(self):
        """Test _is_valid_email_format with valid emails."""
        valid_emails = [
            "test@example.com",
            "user.name@domain.co.uk",
            "user+tag@example.org"
        ]
        
        for email in valid_emails:
            assert self.adapter._is_valid_email_format(email) is True
    
    def test_is_valid_email_format_invalid(self):
        """Test _is_valid_email_format with invalid emails."""
        invalid_emails = [
            "",
            "invalid",
            "@example.com",
            "user@",
            "user@domain",
            None
        ]
        
        for email in invalid_emails:
            assert self.adapter._is_valid_email_format(email) is False
    
    def test_extract_text_from_html_basic(self):
        """Test _extract_text_from_html with basic HTML."""
        html = "<html><body><p>Hello <b>world</b>!</p></body></html>"
        
        result = self.adapter._extract_text_from_html(html)
        assert result == "Hello world!"
    
    def test_extract_text_from_html_with_entities(self):
        """Test _extract_text_from_html with HTML entities."""
        html = "<p>Hello&nbsp;&lt;world&gt;&amp;test</p>"
        
        result = self.adapter._extract_text_from_html(html)
        assert result == "Hello <world>&test"
    
    def test_create_html_from_text_basic(self):
        """Test _create_html_from_text with basic text."""
        text = "Hello world!\nSecond line."
        
        result = self.adapter._create_html_from_text(text)
        assert result == "<html><body>Hello world!<br>\nSecond line.</body></html>"
    
    def test_clean_text_content_whitespace(self):
        """Test _clean_text_content with excessive whitespace."""
        text = "  Hello    world  \n\n  with   spaces  "
        
        result = self.adapter._clean_text_content(text)
        assert result == "Hello world with spaces"
    
    def test_clean_html_content_whitespace(self):
        """Test _clean_html_content with whitespace between tags."""
        html = "<html>  <body>   <p>Content</p>  </body>  </html>"
        
        result = self.adapter._clean_html_content(html)
        assert result == "<html><body><p>Content</p></body></html>"