"""Unit tests for FolderService."""

import pytest
from unittest.mock import Mock, MagicMock, patch
from datetime import datetime

from src.outlook_mcp_server.services.folder_service import FolderService
from src.outlook_mcp_server.models.folder_data import FolderData
from src.outlook_mcp_server.models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    PermissionError,
    ValidationError
)


class TestFolderService:
    """Test cases for FolderService class."""
    
    @pytest.fixture
    def mock_adapter(self):
        """Create a mock Outlook adapter."""
        adapter = Mock()
        adapter.is_connected.return_value = True
        return adapter
    
    @pytest.fixture
    def folder_service(self, mock_adapter):
        """Create a FolderService instance with mocked adapter."""
        return FolderService(mock_adapter)
    
    @pytest.fixture
    def sample_folder_data(self):
        """Create sample FolderData objects for testing."""
        return [
            FolderData(
                id="folder1",
                name="Inbox",
                full_path="Inbox",
                item_count=10,
                unread_count=3,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="folder2",
                name="Sent Items",
                full_path="Sent Items",
                item_count=25,
                unread_count=0,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="folder3",
                name="Subfolder",
                full_path="Inbox/Subfolder",
                item_count=5,
                unread_count=2,
                parent_folder="Inbox",
                folder_type="Mail"
            )
        ]
    
    def test_init(self, mock_adapter):
        """Test FolderService initialization."""
        service = FolderService(mock_adapter)
        assert service.outlook_adapter == mock_adapter
    
    def test_get_folders_success(self, folder_service, mock_adapter, sample_folder_data):
        """Test successful folder retrieval."""
        # Setup mock
        mock_adapter.get_folders.return_value = sample_folder_data
        
        # Execute
        result = folder_service.get_folders()
        
        # Verify
        assert len(result) == 3
        assert all(isinstance(folder, dict) for folder in result)
        
        # Check first folder
        inbox = next(f for f in result if f["name"] == "Inbox")
        assert inbox["id"] == "folder1"
        assert inbox["item_count"] == 10
        assert inbox["unread_count"] == 3
        assert inbox["accessible"] is True
        assert "display_name" in inbox
        
        # Verify adapter was called
        mock_adapter.is_connected.assert_called_once()
        mock_adapter.get_folders.assert_called_once()
    
    def test_get_folders_not_connected(self, folder_service, mock_adapter):
        """Test get_folders when not connected to Outlook."""
        # Setup mock
        mock_adapter.is_connected.return_value = False
        
        # Execute and verify
        with pytest.raises(OutlookConnectionError, match="Not connected to Outlook"):
            folder_service.get_folders()
        
        # Verify adapter methods called
        mock_adapter.is_connected.assert_called_once()
        mock_adapter.get_folders.assert_not_called()
    
    def test_get_folders_permission_error(self, folder_service, mock_adapter):
        """Test get_folders with permission error."""
        # Setup mock
        mock_adapter.get_folders.side_effect = PermissionError("folders", "Access denied")
        
        # Execute and verify
        with pytest.raises(PermissionError):
            folder_service.get_folders()
    
    def test_get_folders_adapter_exception(self, folder_service, mock_adapter):
        """Test get_folders with adapter exception."""
        # Setup mock
        mock_adapter.get_folders.side_effect = Exception("COM error")
        
        # Execute and verify
        with pytest.raises(OutlookConnectionError, match="Failed to retrieve folders"):
            folder_service.get_folders()
    
    def test_get_folders_permission_exception_in_generic_error(self, folder_service, mock_adapter):
        """Test get_folders converts permission-related generic errors."""
        # Setup mock
        mock_adapter.get_folders.side_effect = Exception("Access denied to resource")
        
        # Execute and verify
        with pytest.raises(PermissionError, match="Access denied to folders"):
            folder_service.get_folders()
    
    def test_get_folders_with_invalid_folder_data(self, folder_service, mock_adapter):
        """Test get_folders handles invalid folder data gracefully."""
        # Create valid folder data first
        valid_folder = FolderData(
            id="valid1",
            name="Valid",
            full_path="Valid"
        )
        
        # Create a folder that will cause transformation error
        problematic_folder = FolderData(
            id="problem1",
            name="Problem",
            full_path="Problem"
        )
        
        # Setup mock
        mock_adapter.get_folders.return_value = [problematic_folder, valid_folder]
        
        # Mock the transformation method to simulate error for one folder
        original_transform = folder_service._transform_folder_to_json
        
        def mock_transform(folder_data):
            if folder_data.name == "Problem":
                raise Exception("Transformation error")
            # Call the original method for valid folders
            return original_transform(folder_data)
        
        folder_service._transform_folder_to_json = mock_transform
        
        try:
            # Execute
            result = folder_service.get_folders()
            
            # Verify - should only return valid folder
            assert len(result) == 1
            assert result[0]["name"] == "Valid"
        finally:
            # Restore original method
            folder_service._transform_folder_to_json = original_transform
    
    def test_validate_folder_success(self, folder_service, mock_adapter):
        """Test successful folder validation."""
        # Setup mock
        mock_adapter.validate_folder_access.return_value = True
        
        # Execute
        result = folder_service.validate_folder("Inbox")
        
        # Verify
        assert result is True
        mock_adapter.is_connected.assert_called_once()
        mock_adapter.validate_folder_access.assert_called_once_with("Inbox")
    
    def test_validate_folder_invalid_name(self, folder_service, mock_adapter):
        """Test folder validation with invalid name."""
        # Test empty name
        assert folder_service.validate_folder("") is False
        assert folder_service.validate_folder(None) is False
        
        # Test invalid characters
        assert folder_service.validate_folder("Test<>Folder") is False
        
        # Verify adapter not called for invalid names
        mock_adapter.validate_folder_access.assert_not_called()
    
    def test_validate_folder_not_connected(self, folder_service, mock_adapter):
        """Test folder validation when not connected."""
        # Setup mock
        mock_adapter.is_connected.return_value = False
        
        # Execute
        result = folder_service.validate_folder("Inbox")
        
        # Verify
        assert result is False
        mock_adapter.validate_folder_access.assert_not_called()
    
    def test_validate_folder_adapter_exception(self, folder_service, mock_adapter):
        """Test folder validation with adapter exception."""
        # Setup mock
        mock_adapter.validate_folder_access.side_effect = Exception("Test error")
        
        # Execute
        result = folder_service.validate_folder("Inbox")
        
        # Verify
        assert result is False
    
    def test_get_folder_by_name_success(self, folder_service, mock_adapter):
        """Test successful get folder by name."""
        # Setup mock COM object
        mock_com_folder = Mock()
        mock_com_folder.Name = "Inbox"
        mock_com_folder.EntryID = "folder123"
        
        # Setup folder data
        folder_data = FolderData(
            id="folder123",
            name="Inbox",
            full_path="Inbox",
            item_count=10,
            unread_count=3
        )
        
        # Setup mocks
        mock_adapter.get_folder_by_name.return_value = mock_com_folder
        mock_adapter._transform_folder_to_data.return_value = folder_data
        
        # Execute
        result = folder_service.get_folder_by_name("Inbox")
        
        # Verify
        assert isinstance(result, dict)
        assert result["name"] == "Inbox"
        assert result["id"] == "folder123"
        assert result["accessible"] is True
        
        # Verify adapter calls
        mock_adapter.is_connected.assert_called_once()
        mock_adapter.get_folder_by_name.assert_called_once_with("Inbox")
        mock_adapter._transform_folder_to_data.assert_called_once_with(mock_com_folder, "")
    
    def test_get_folder_by_name_invalid_input(self, folder_service, mock_adapter):
        """Test get folder by name with invalid input."""
        # Test empty name
        with pytest.raises(ValidationError, match="Folder name must be a non-empty string"):
            folder_service.get_folder_by_name("")
        
        # Test None
        with pytest.raises(ValidationError, match="Folder name must be a non-empty string"):
            folder_service.get_folder_by_name(None)
        
        # Test invalid characters
        with pytest.raises(ValidationError, match="Invalid folder name format"):
            folder_service.get_folder_by_name("Test<>Folder")
    
    def test_get_folder_by_name_not_connected(self, folder_service, mock_adapter):
        """Test get folder by name when not connected."""
        # Setup mock
        mock_adapter.is_connected.return_value = False
        
        # Execute and verify
        with pytest.raises(OutlookConnectionError, match="Not connected to Outlook"):
            folder_service.get_folder_by_name("Inbox")
    
    def test_get_folder_by_name_not_found(self, folder_service, mock_adapter):
        """Test get folder by name when folder not found."""
        # Setup mock
        mock_adapter.get_folder_by_name.side_effect = FolderNotFoundError("TestFolder")
        
        # Execute and verify
        with pytest.raises(FolderNotFoundError):
            folder_service.get_folder_by_name("TestFolder")
    
    def test_get_folder_by_name_permission_error(self, folder_service, mock_adapter):
        """Test get folder by name with permission error."""
        # Setup mock
        mock_adapter.get_folder_by_name.side_effect = PermissionError("TestFolder", "Access denied")
        
        # Execute and verify
        with pytest.raises(PermissionError):
            folder_service.get_folder_by_name("TestFolder")
    
    def test_get_folder_by_name_generic_exception_access_denied(self, folder_service, mock_adapter):
        """Test get folder by name converts access denied exceptions."""
        # Setup mock
        mock_adapter.get_folder_by_name.side_effect = Exception("Access denied to resource")
        
        # Execute and verify
        with pytest.raises(PermissionError, match="Access denied to folder"):
            folder_service.get_folder_by_name("TestFolder")
    
    def test_get_folder_by_name_generic_exception_not_found(self, folder_service, mock_adapter):
        """Test get folder by name converts not found exceptions."""
        # Setup mock
        mock_adapter.get_folder_by_name.side_effect = Exception("Folder not found")
        
        # Execute and verify
        with pytest.raises(FolderNotFoundError):
            folder_service.get_folder_by_name("TestFolder")
    
    def test_get_folder_by_name_generic_exception_other(self, folder_service, mock_adapter):
        """Test get folder by name converts other exceptions to connection error."""
        # Setup mock
        mock_adapter.get_folder_by_name.side_effect = Exception("COM error")
        
        # Execute and verify
        with pytest.raises(OutlookConnectionError, match="Failed to retrieve folder"):
            folder_service.get_folder_by_name("TestFolder")
    
    def test_transform_folder_to_json(self, folder_service):
        """Test folder data transformation to JSON."""
        # Create test folder data
        folder_data = FolderData(
            id="test123",
            name="Test Folder",
            full_path="Parent/Test Folder",
            item_count=15,
            unread_count=5,
            parent_folder="Parent",
            folder_type="Mail"
        )
        
        # Execute
        result = folder_service._transform_folder_to_json(folder_data)
        
        # Verify
        assert isinstance(result, dict)
        assert result["id"] == "test123"
        assert result["name"] == "Test Folder"
        assert result["full_path"] == "Parent/Test Folder"
        assert result["item_count"] == 15
        assert result["unread_count"] == 5
        assert result["parent_folder"] == "Parent"
        assert result["folder_type"] == "Mail"
        assert result["accessible"] is True
        assert "has_subfolders" in result
        assert "display_name" in result
    
    def test_transform_folder_to_json_invalid_data(self, folder_service):
        """Test folder transformation with invalid data."""
        # Create folder data and then corrupt it to bypass validation
        folder_data = FolderData(
            id="valid_id",
            name="Test",
            full_path="Test"
        )
        
        # Corrupt the data after creation to bypass __post_init__ validation
        folder_data.id = ""  # Make it invalid
        
        # Execute and verify
        with pytest.raises(ValidationError):
            folder_service._transform_folder_to_json(folder_data)
    
    def test_has_subfolders(self, folder_service):
        """Test subfolder detection logic."""
        # Test common parent folder
        inbox_folder = FolderData(
            id="inbox1",
            name="Inbox",
            full_path="Inbox"
        )
        assert folder_service._has_subfolders(inbox_folder) is True
        
        # Test regular folder
        custom_folder = FolderData(
            id="custom1",
            name="Custom Folder",
            full_path="Custom Folder"
        )
        assert folder_service._has_subfolders(custom_folder) is False
    
    def test_get_display_name(self, folder_service):
        """Test display name generation."""
        # Test root folder
        root_folder = FolderData(
            id="root1",
            name="Inbox",
            full_path="Inbox",
            parent_folder=""
        )
        assert folder_service._get_display_name(root_folder) == "Inbox"
        
        # Test subfolder
        sub_folder = FolderData(
            id="sub1",
            name="Subfolder",
            full_path="Inbox/Subfolder",
            parent_folder="Inbox"
        )
        assert folder_service._get_display_name(sub_folder) == "Inbox > Subfolder"
    
    def test_get_folder_statistics_success(self, folder_service, mock_adapter, sample_folder_data):
        """Test successful folder statistics generation."""
        # Setup mock to return JSON folders
        mock_adapter.get_folders.return_value = sample_folder_data
        
        # Execute
        result = folder_service.get_folder_statistics()
        
        # Verify
        assert isinstance(result, dict)
        assert result["total_folders"] == 3
        assert result["total_items"] == 40  # 10 + 25 + 5
        assert result["total_unread"] == 5   # 3 + 0 + 2
        assert "folder_types" in result
        assert "folders_by_type" in result
        assert result["folder_types"]["Mail"] == 3
        assert len(result["folders_by_type"]["Mail"]) == 3
    
    def test_get_folder_statistics_connection_error(self, folder_service, mock_adapter):
        """Test folder statistics with connection error."""
        # Setup mock
        mock_adapter.is_connected.return_value = False
        
        # Execute and verify
        with pytest.raises(OutlookConnectionError):
            folder_service.get_folder_statistics()
    
    def test_get_folder_statistics_permission_error(self, folder_service, mock_adapter):
        """Test folder statistics with permission error."""
        # Setup mock
        mock_adapter.get_folders.side_effect = PermissionError("folders", "Access denied")
        
        # Execute and verify
        with pytest.raises(PermissionError):
            folder_service.get_folder_statistics()
    
    def test_get_folder_statistics_generic_error(self, folder_service, mock_adapter):
        """Test folder statistics with generic error."""
        # Setup mock
        mock_adapter.get_folders.side_effect = Exception("Generic error")
        
        # Execute and verify - the error will come from get_folders, not get_folder_statistics
        with pytest.raises(OutlookConnectionError, match="Failed to retrieve folders"):
            folder_service.get_folder_statistics()


class TestFolderServiceIntegration:
    """Integration tests for FolderService with more realistic scenarios."""
    
    @pytest.fixture
    def mock_adapter_with_realistic_data(self):
        """Create a mock adapter with realistic folder structure."""
        adapter = Mock()
        adapter.is_connected.return_value = True
        
        # Create realistic folder structure
        folders = [
            FolderData(
                id="inbox_id",
                name="Inbox",
                full_path="Inbox",
                item_count=150,
                unread_count=12,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="sent_id",
                name="Sent Items",
                full_path="Sent Items",
                item_count=89,
                unread_count=0,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="drafts_id",
                name="Drafts",
                full_path="Drafts",
                item_count=3,
                unread_count=3,
                parent_folder="",
                folder_type="Mail"
            ),
            FolderData(
                id="contacts_id",
                name="Contacts",
                full_path="Contacts",
                item_count=45,
                unread_count=0,
                parent_folder="",
                folder_type="Contact"
            )
        ]
        
        adapter.get_folders.return_value = folders
        return adapter
    
    def test_realistic_folder_retrieval(self, mock_adapter_with_realistic_data):
        """Test folder retrieval with realistic data structure."""
        service = FolderService(mock_adapter_with_realistic_data)
        
        # Execute
        folders = service.get_folders()
        
        # Verify structure
        assert len(folders) == 4
        
        # Check specific folders
        inbox = next(f for f in folders if f["name"] == "Inbox")
        assert inbox["item_count"] == 150
        assert inbox["unread_count"] == 12
        assert inbox["folder_type"] == "Mail"
        
        contacts = next(f for f in folders if f["name"] == "Contacts")
        assert contacts["folder_type"] == "Contact"
        assert contacts["unread_count"] == 0
    
    def test_realistic_folder_statistics(self, mock_adapter_with_realistic_data):
        """Test folder statistics with realistic data."""
        service = FolderService(mock_adapter_with_realistic_data)
        
        # Execute
        stats = service.get_folder_statistics()
        
        # Verify statistics
        assert stats["total_folders"] == 4
        assert stats["total_items"] == 287  # 150 + 89 + 3 + 45
        assert stats["total_unread"] == 15  # 12 + 0 + 3 + 0
        assert stats["folder_types"]["Mail"] == 3
        assert stats["folder_types"]["Contact"] == 1
        
        # Verify folder grouping
        mail_folders = stats["folders_by_type"]["Mail"]
        assert len(mail_folders) == 3
        assert any(f["name"] == "Inbox" for f in mail_folders)
        
        contact_folders = stats["folders_by_type"]["Contact"]
        assert len(contact_folders) == 1
        assert contact_folders[0]["name"] == "Contacts"