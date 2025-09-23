"""Folder service layer for handling folder operations."""

import logging
from typing import List, Dict, Any, Optional
from ..adapters.outlook_adapter import OutlookAdapter
from ..models.folder_data import FolderData
from ..models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    PermissionError,
    ValidationError
)


logger = logging.getLogger(__name__)


class FolderService:
    """Service layer for folder management operations."""
    
    def __init__(self, outlook_adapter: OutlookAdapter):
        """
        Initialize the folder service.
        
        Args:
            outlook_adapter: The Outlook adapter instance for COM operations
        """
        self.outlook_adapter = outlook_adapter
        
    def get_folders(self) -> List[Dict[str, Any]]:
        """
        Get all available email folders with proper error handling.
        
        Returns:
            List[Dict[str, Any]]: List of folder data in JSON format
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            PermissionError: If access to folders is denied
        """
        try:
            logger.info("Retrieving all available folders")
            
            # Ensure we're connected to Outlook
            if not self.outlook_adapter.is_connected():
                logger.error("Outlook adapter is not connected")
                raise OutlookConnectionError("Not connected to Outlook")
            
            # Get folders from the adapter
            folder_data_list = self.outlook_adapter.get_folders()
            
            # Transform to JSON format
            json_folders = []
            for folder_data in folder_data_list:
                try:
                    json_folder = self._transform_folder_to_json(folder_data)
                    json_folders.append(json_folder)
                except Exception as e:
                    logger.warning(f"Error transforming folder '{folder_data.name}': {str(e)}")
                    continue
            
            # Sort folders by full path for consistent ordering
            json_folders.sort(key=lambda x: x.get('full_path', ''))
            
            logger.info(f"Successfully retrieved {len(json_folders)} folders")
            return json_folders
            
        except (OutlookConnectionError, PermissionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error retrieving folders: {str(e)}")
            # Check if it's a permission-related error
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError("folders", f"Access denied to folders: {str(e)}")
            # Otherwise, treat as connection error
            raise OutlookConnectionError(f"Failed to retrieve folders: {str(e)}")
    
    def validate_folder(self, folder_name: str) -> bool:
        """
        Validate that a folder exists and is accessible.
        
        Args:
            folder_name: Name of the folder to validate
            
        Returns:
            bool: True if folder is valid and accessible, False otherwise
        """
        try:
            # Basic validation
            if not folder_name or not isinstance(folder_name, str):
                logger.debug("Invalid folder name provided")
                return False
            
            # Validate folder name format
            if not FolderData.validate_folder_name(folder_name):
                logger.debug(f"Folder name format is invalid: {folder_name}")
                return False
            
            # Check if we're connected
            if not self.outlook_adapter.is_connected():
                logger.debug("Outlook adapter is not connected")
                return False
            
            # Use adapter to validate folder access
            is_valid = self.outlook_adapter.validate_folder_access(folder_name)
            
            logger.debug(f"Folder validation for '{folder_name}': {is_valid}")
            return is_valid
            
        except Exception as e:
            logger.debug(f"Error validating folder '{folder_name}': {str(e)}")
            return False
    
    def get_folder_by_name(self, folder_name: str) -> Dict[str, Any]:
        """
        Get a specific folder by name with validation and access control.
        
        Args:
            folder_name: Name of the folder to retrieve
            
        Returns:
            Dict[str, Any]: Folder data in JSON format
            
        Raises:
            ValidationError: If folder name is invalid
            FolderNotFoundError: If folder is not found
            PermissionError: If access to folder is denied
            OutlookConnectionError: If not connected to Outlook
        """
        try:
            logger.debug(f"Retrieving folder by name: {folder_name}")
            
            # Validate input
            if not folder_name or not isinstance(folder_name, str):
                raise ValidationError("Folder name must be a non-empty string", "folder_name")
            
            # Validate folder name format
            if not FolderData.validate_folder_name(folder_name):
                raise ValidationError(f"Invalid folder name format: {folder_name}", "folder_name")
            
            # Ensure we're connected
            if not self.outlook_adapter.is_connected():
                raise OutlookConnectionError("Not connected to Outlook")
            
            # Get the folder from adapter
            folder_com_object = self.outlook_adapter.get_folder_by_name(folder_name)
            
            # Transform COM object to FolderData
            folder_data = self.outlook_adapter._transform_folder_to_data(folder_com_object, "")
            
            # Transform to JSON format
            json_folder = self._transform_folder_to_json(folder_data)
            
            logger.debug(f"Successfully retrieved folder: {folder_name}")
            return json_folder
            
        except (ValidationError, FolderNotFoundError, PermissionError, OutlookConnectionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error retrieving folder '{folder_name}': {str(e)}")
            # Check error type and convert to appropriate exception
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError(folder_name, f"Access denied to folder '{folder_name}': {str(e)}")
            elif any(keyword in str(e).lower() for keyword in ['not found', 'invalid', 'missing']):
                raise FolderNotFoundError(folder_name)
            else:
                raise OutlookConnectionError(f"Failed to retrieve folder '{folder_name}': {str(e)}")
    
    def _transform_folder_to_json(self, folder_data: FolderData) -> Dict[str, Any]:
        """
        Transform FolderData object to JSON-serializable dictionary.
        
        Args:
            folder_data: The FolderData object to transform
            
        Returns:
            Dict[str, Any]: JSON-serializable folder data
            
        Raises:
            ValidationError: If folder data is invalid
        """
        try:
            # Validate the folder data
            folder_data.validate()
            
            # Use the built-in to_dict method
            json_data = folder_data.to_dict()
            
            # Add additional metadata for API response
            json_data.update({
                "accessible": True,  # If we got this far, folder is accessible
                "has_subfolders": self._has_subfolders(folder_data),
                "display_name": self._get_display_name(folder_data)
            })
            
            return json_data
            
        except ValidationError:
            # Re-raise validation errors
            raise
        except Exception as e:
            logger.error(f"Error transforming folder data to JSON: {str(e)}")
            raise ValidationError(f"Failed to transform folder data: {str(e)}")
    
    def _has_subfolders(self, folder_data: FolderData) -> bool:
        """
        Determine if a folder has subfolders based on available information.
        
        Args:
            folder_data: The folder data to check
            
        Returns:
            bool: True if folder likely has subfolders, False otherwise
        """
        try:
            # This is a heuristic since we don't have direct subfolder info in FolderData
            # We could enhance this by checking with the adapter if needed
            
            # Common folders that typically have subfolders
            common_parent_folders = [
                "Inbox", "Sent Items", "Deleted Items", "Drafts", 
                "Outbox", "Junk Email", "Archive", "Personal Folders"
            ]
            
            return folder_data.name in common_parent_folders
            
        except Exception as e:
            logger.debug(f"Error checking subfolders for '{folder_data.name}': {str(e)}")
            return False
    
    def _get_display_name(self, folder_data: FolderData) -> str:
        """
        Get a user-friendly display name for the folder.
        
        Args:
            folder_data: The folder data
            
        Returns:
            str: User-friendly display name
        """
        try:
            # For root folders, use the name as-is
            if not folder_data.parent_folder:
                return folder_data.name
            
            # For subfolders, show the hierarchy
            return folder_data.full_path.replace("/", " > ")
            
        except Exception as e:
            logger.debug(f"Error generating display name for '{folder_data.name}': {str(e)}")
            return folder_data.name
    
    def get_folder_statistics(self) -> Dict[str, Any]:
        """
        Get statistics about all folders.
        
        Returns:
            Dict[str, Any]: Folder statistics including counts and types
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            PermissionError: If access to folders is denied
        """
        try:
            logger.debug("Generating folder statistics")
            
            # Get all folders
            folders = self.get_folders()
            
            # Calculate statistics
            stats = {
                "total_folders": len(folders),
                "total_items": sum(folder.get("item_count", 0) for folder in folders),
                "total_unread": sum(folder.get("unread_count", 0) for folder in folders),
                "folder_types": {},
                "folders_by_type": {}
            }
            
            # Group by folder type
            for folder in folders:
                folder_type = folder.get("folder_type", "Unknown")
                if folder_type not in stats["folder_types"]:
                    stats["folder_types"][folder_type] = 0
                    stats["folders_by_type"][folder_type] = []
                
                stats["folder_types"][folder_type] += 1
                stats["folders_by_type"][folder_type].append({
                    "name": folder.get("name"),
                    "full_path": folder.get("full_path"),
                    "item_count": folder.get("item_count", 0),
                    "unread_count": folder.get("unread_count", 0)
                })
            
            logger.debug(f"Generated statistics for {stats['total_folders']} folders")
            return stats
            
        except (OutlookConnectionError, PermissionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Error generating folder statistics: {str(e)}")
            raise OutlookConnectionError(f"Failed to generate folder statistics: {str(e)}")
    
    def debug_folder_names(self) -> Dict[str, Any]:
        """
        Debug method to get actual folder names for troubleshooting localization issues.
        
        Returns:
            Dict[str, Any]: Debug information about default folder names
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
        """
        try:
            logger.debug("Getting debug information for folder names")
            
            # Ensure we're connected
            if not self.outlook_adapter.is_connected():
                raise OutlookConnectionError("Not connected to Outlook")
            
            # Get default folder names from adapter
            default_folder_names = self.outlook_adapter.get_default_folder_names()
            
            # Get all available folders
            try:
                all_folders = self.get_folders()
                folder_list = [{"name": f.get("name"), "full_path": f.get("full_path"), "type": f.get("folder_type")} for f in all_folders]
            except Exception as e:
                logger.warning(f"Could not get all folders: {e}")
                folder_list = []
            
            debug_info = {
                "default_folders": default_folder_names,
                "all_folders": folder_list,
                "folder_count": len(folder_list),
                "common_inbox_names": [
                    "Inbox", "收件匣", "收件夾", "收件箱", "受信トレイ"
                ],
                "instructions": {
                    "message": "Use the 'actual_name' from default_folders for your search queries",
                    "example": "If you see 'actual_name': '收件匣' for folder ID 6, use '收件匣' in your search_emails request"
                }
            }
            
            logger.debug(f"Generated debug info for {len(folder_list)} folders")
            return debug_info
            
        except OutlookConnectionError:
            raise
        except Exception as e:
            logger.error(f"Error generating folder debug info: {str(e)}")
            raise OutlookConnectionError(f"Failed to generate folder debug info: {str(e)}")