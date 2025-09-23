"""Outlook COM adapter for interfacing with Microsoft Outlook."""

import logging
import re
import time
from datetime import datetime
from typing import Optional, List, Any, Tuple
from unittest.mock import Mock
import win32com.client
import pythoncom
from ..models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    EmailNotFoundError,
    PermissionError,
    ValidationError
)
from ..models.folder_data import FolderData
from ..models.email_data import EmailData


logger = logging.getLogger(__name__)


class OutlookAdapter:
    """Low-level interface with Microsoft Outlook COM objects."""
    
    def __init__(self):
        """Initialize the Outlook adapter."""
        self._outlook_app: Optional[Any] = None
        self._namespace: Optional[Any] = None
        self._connected = False
        
    def connect(self) -> bool:
        """
        Establish connection to Microsoft Outlook.
        
        Returns:
            bool: True if connection successful, False otherwise
            
        Raises:
            OutlookConnectionError: If connection cannot be established
        """
        try:
            logger.info("Attempting to connect to Microsoft Outlook")
            
            # Initialize COM
            pythoncom.CoInitialize()
            
            # Try to get existing Outlook instance first
            try:
                self._outlook_app = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Connected to existing Outlook instance")
            except:
                # If no existing instance, create new one
                self._outlook_app = win32com.client.Dispatch("Outlook.Application")
                logger.info("Created new Outlook instance")
            
            # Get the MAPI namespace
            self._namespace = self._outlook_app.GetNamespace("MAPI")
            
            # Test the connection by trying to access folders
            self._test_connection()
            
            self._connected = True
            logger.info("Successfully connected to Outlook")
            return True
            
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {str(e)}")
            self._cleanup_connection()
            raise OutlookConnectionError(f"Failed to connect to Outlook: {str(e)}")
    
    def _test_connection(self) -> None:
        """
        Test the Outlook connection by accessing basic functionality.
        
        Raises:
            OutlookConnectionError: If connection test fails
        """
        try:
            # Try to access the default folders to verify connection
            if not self._namespace:
                raise OutlookConnectionError("Namespace not available")
                
            # Try to get the default inbox folder
            inbox = self._namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if not inbox:
                raise OutlookConnectionError("Cannot access default inbox folder")
                
            logger.debug("Connection test successful - can access inbox")
            
        except Exception as e:
            logger.error(f"Connection test failed: {str(e)}")
            raise OutlookConnectionError(f"Connection test failed: {str(e)}")
    
    def disconnect(self) -> None:
        """Disconnect from Outlook and cleanup resources."""
        try:
            logger.info("Disconnecting from Outlook")
            self._cleanup_connection()
            logger.info("Successfully disconnected from Outlook")
        except Exception as e:
            logger.error(f"Error during disconnect: {str(e)}")
    
    def _cleanup_connection(self) -> None:
        """Clean up COM objects and connection state."""
        try:
            self._namespace = None
            self._outlook_app = None
            self._connected = False
            
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore errors during COM cleanup
                
        except Exception as e:
            logger.error(f"Error during connection cleanup: {str(e)}")
    
    def is_connected(self) -> bool:
        """
        Check if adapter is connected to Outlook.
        
        Returns:
            bool: True if connected, False otherwise
        """
        logger.debug(f"Checking connection: _connected={self._connected}, _outlook_app={self._outlook_app is not None}, _namespace={self._namespace is not None}")
        
        if not self._connected or not self._outlook_app or not self._namespace:
            logger.debug("Basic connection check failed")
            return False
            
        try:
            # Initialize COM for this thread if needed
            pythoncom.CoInitialize()
            
            # Test connection by accessing a basic property
            # Try multiple approaches to verify connection
            
            # First, try to access the namespace
            if self._namespace is None:
                logger.debug("Namespace is None")
                return False
            
            # Try to get the default folder (inbox)
            logger.debug("Attempting to access inbox folder")
            inbox = self._namespace.GetDefaultFolder(6)  # olFolderInbox = 6
            
            # If we can access the inbox, connection is good
            if inbox is not None:
                logger.debug("Successfully accessed inbox folder")
                return True
            else:
                logger.debug("Inbox folder is None")
                self._connected = False
                return False
                
        except Exception as e:
            # Check if this is a COM threading error
            if "marshaled for a different thread" in str(e) or "-2147417842" in str(e):
                logger.debug("COM threading issue detected - connection exists but not accessible from this thread")
                # Don't mark as disconnected for threading issues, just return the basic connection status
                return self._connected
            else:
                # Log the specific error for debugging
                logger.debug(f"Connection test failed with exception: {str(e)}")
                self._connected = False
                return False
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def get_namespace(self) -> Any:
        """
        Get the MAPI namespace object.
        
        Returns:
            COM object: The MAPI namespace
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        return self._namespace
    
    def get_folder_by_id(self, folder_id: str) -> Any:
        """
        Get folder by ID from Outlook.
        
        Args:
            folder_id: ID of the folder to retrieve
            
        Returns:
            COM object: The folder object
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If folder is not found
            PermissionError: If access to folder is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not folder_id or not isinstance(folder_id, str):
            raise FolderNotFoundError(folder_id or "")
        
        try:
            logger.debug(f"Looking for folder by ID: {folder_id}")
            
            # Use existing namespace to avoid COM threading issues
            # Get main folders and check their IDs
            main_folders = [
                (6, "Inbox"),           # olFolderInbox
                (5, "Sent Items"),      # olFolderSentMail  
                (16, "Drafts"),         # olFolderDrafts
                (3, "Deleted Items"),   # olFolderDeletedItems
                (4, "Outbox"),          # olFolderOutbox
                (9, "Calendar"),        # olFolderCalendar
                (10, "Contacts"),       # olFolderContacts
                (13, "Journal"),        # olFolderJournal
                (12, "Tasks")           # olFolderTasks
            ]
            
            for folder_id_num, folder_name in main_folders:
                try:
                    folder = self._namespace.GetDefaultFolder(folder_id_num)
                    if folder and hasattr(folder, 'EntryID'):
                        if folder.EntryID == folder_id:
                            logger.debug(f"Found folder by ID: {folder_id} -> {folder_name}")
                            return folder
                except Exception as e:
                    logger.debug(f"Error checking folder {folder_name}: {e}")
                    continue
                
            
            # If not found in default folders, we could search all folders but that's expensive
            # For now, just check if it's one of the default folders we missed
                
            logger.warning(f"Folder not found by ID: {folder_id}")
            raise FolderNotFoundError(folder_id)
                
        except FolderNotFoundError:
            raise
        except Exception as e:
            logger.error(f"Error accessing folder by ID '{folder_id}': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_id, f"Access denied to folder ID '{folder_id}'")
            raise FolderNotFoundError(folder_id)
    
    def _search_folder_by_id_recursive(self, folder: Any, target_id: str) -> Optional[Any]:
        """
        Recursively search for a folder by ID.
        
        Args:
            folder: The folder to search in
            target_id: ID of the target folder
            
        Returns:
            COM object or None: The found folder or None if not found
        """
        try:
            # Check if current folder matches
            if hasattr(folder, 'EntryID') and folder.EntryID == target_id:
                return folder
            
            # Search in subfolders
            if hasattr(folder, 'Folders'):
                for subfolder in folder.Folders:
                    result = self._search_folder_by_id_recursive(subfolder, target_id)
                    if result:
                        return result
            
            return None
            
        except Exception as e:
            logger.debug(f"Error searching in folder by ID: {str(e)}")
            return None
    
    def get_folder_by_name_or_id(self, identifier: str) -> Any:
        """
        Get folder by name or ID from Outlook.
        
        Args:
            identifier: Name or ID of the folder to retrieve
            
        Returns:
            COM object: The folder object
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If folder is not found
            PermissionError: If access to folder is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not identifier or not isinstance(identifier, str):
            raise FolderNotFoundError(identifier or "")
        
        # First try as folder ID (if it looks like an EntryID - long hex string)
        if len(identifier) > 50 and all(c in '0123456789ABCDEFabcdef' for c in identifier):
            logger.debug(f"Identifier looks like folder ID, trying ID lookup: {identifier[:20]}...")
            try:
                return self.get_folder_by_id(identifier)
            except FolderNotFoundError:
                logger.debug(f"Not found as ID, trying as name: {identifier[:20]}...")
        
        # Try as folder name
        logger.debug(f"Trying as folder name: {identifier}")
        return self.get_folder_by_name(identifier)
    
    def get_folder_by_name(self, name: str) -> Any:
        """
        Get folder by name from Outlook.
        
        Args:
            name: Name of the folder to retrieve
            
        Returns:
            COM object: The folder object
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If folder is not found
            PermissionError: If access to folder is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not name or not isinstance(name, str):
            raise FolderNotFoundError(name or "")
        
        try:
            logger.debug(f"Looking for folder: {name}")
            
            # Try to get folder by name
            # First check default folders with localized names
            default_folders = {
                # English names
                "Inbox": 6,      # olFolderInbox
                "Outbox": 4,     # olFolderOutbox
                "Sent Items": 5, # olFolderSentMail
                "Deleted Items": 3, # olFolderDeletedItems
                "Drafts": 16,    # olFolderDrafts
                "Junk Email": 23, # olFolderJunk
                # Chinese Traditional names
                "收件匣": 6,      # olFolderInbox
                "收件夾": 6,      # olFolderInbox (alternative)
                "寄件匣": 4,      # olFolderOutbox
                "寄件備份": 5,    # olFolderSentMail
                "已傳送的郵件": 5, # olFolderSentMail (alternative)
                "刪除的郵件": 3,  # olFolderDeletedItems
                "已刪除的郵件": 3, # olFolderDeletedItems (alternative)
                "草稿": 16,      # olFolderDrafts
                "垃圾郵件": 23,   # olFolderJunk
                # Chinese Simplified names
                "收件箱": 6,      # olFolderInbox
                "发件箱": 4,      # olFolderOutbox
                "已发送邮件": 5,  # olFolderSentMail
                "已删除邮件": 3,  # olFolderDeletedItems
                "草稿箱": 16,     # olFolderDrafts
                "垃圾邮件": 23,   # olFolderJunk
                # Japanese names
                "受信トレイ": 6,   # olFolderInbox
                "送信トレイ": 4,   # olFolderOutbox
                "送信済みアイテム": 5, # olFolderSentMail
                "削除済みアイテム": 3, # olFolderDeletedItems
                "下書き": 16,     # olFolderDrafts
                "迷惑メール": 23,  # olFolderJunk
            }
            
            # Try exact match first
            if name in default_folders:
                try:
                    folder = self._namespace.GetDefaultFolder(default_folders[name])
                    if folder:
                        logger.debug(f"Found default folder by exact match: {name}")
                        return folder
                except Exception as e:
                    logger.debug(f"Error accessing default folder {name}: {e}")
            
            # Try to get default folders and match by actual name
            try:
                logger.debug(f"Trying to match folder name '{name}' with actual folder names")
                for folder_id in [6, 5, 4, 3, 16, 23]:  # Common default folders
                    try:
                        folder = self._namespace.GetDefaultFolder(folder_id)
                        if folder and hasattr(folder, 'Name'):
                            actual_name = folder.Name
                            logger.debug(f"Checking folder ID {folder_id}: '{actual_name}' vs '{name}'")
                            if actual_name == name:
                                logger.debug(f"Found default folder by name match: {name} (ID: {folder_id})")
                                return folder
                    except Exception as e:
                        logger.debug(f"Error checking folder ID {folder_id}: {e}")
                        continue
            except Exception as e:
                logger.debug(f"Error during folder name matching: {e}")
            
            # If not a default folder, search through all folders
            folders = self._namespace.Folders
            for folder in folders:
                if self._search_folder_recursive(folder, name):
                    logger.debug(f"Found folder: {name}")
                    return self._search_folder_recursive(folder, name)
            
            logger.warning(f"Folder not found: {name}")
            raise FolderNotFoundError(name)
            
        except FolderNotFoundError:
            raise
        except Exception as e:
            logger.error(f"Error accessing folder '{name}': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(name, f"Access denied to folder '{name}'")
            raise FolderNotFoundError(name)
    
    def _search_folder_recursive(self, folder: Any, target_name: str) -> Optional[Any]:
        """
        Recursively search for a folder by name.
        
        Args:
            folder: The folder to search in
            target_name: Name of the target folder
            
        Returns:
            COM object or None: The found folder or None if not found
        """
        try:
            # Check if current folder matches
            if hasattr(folder, 'Name') and folder.Name == target_name:
                return folder
            
            # Search in subfolders
            if hasattr(folder, 'Folders'):
                for subfolder in folder.Folders:
                    result = self._search_folder_recursive(subfolder, target_name)
                    if result:
                        return result
            
            return None
            
        except Exception as e:
            logger.debug(f"Error searching in folder: {str(e)}")
            return None
    
    def get_folders(self) -> List[FolderData]:
        """
        Get all available Outlook folders.
        
        Returns:
            List[FolderData]: List of all available folders
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            PermissionError: If access to folders is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            logger.debug("Retrieving all Outlook folders")
            
            # For HTTP requests, we need to create a new Outlook connection in this thread
            # because COM objects cannot be shared across threads
            try:
                outlook_app = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook_app.GetNamespace("MAPI")
                
                # Test the connection
                namespace.GetDefaultFolder(6)  # Try to access inbox
                
                logger.debug("Created thread-local Outlook connection")
                
            except Exception as e:
                logger.error(f"Failed to create thread-local Outlook connection: {e}")
                # Fall back to original namespace (might work in some cases)
                namespace = self._namespace
            
            folders = []
            
            # Get only the main default folders (faster approach)
            logger.debug("Getting main Outlook folders only (non-recursive)")
            
            # Define the main folders we want to retrieve
            main_folders = [
                (6, "Inbox"),           # olFolderInbox
                (5, "Sent Items"),      # olFolderSentMail  
                (16, "Drafts"),         # olFolderDrafts
                (3, "Deleted Items"),   # olFolderDeletedItems
                (4, "Outbox"),          # olFolderOutbox
                (9, "Calendar"),        # olFolderCalendar
                (10, "Contacts"),       # olFolderContacts
                (13, "Journal"),        # olFolderJournal
                (12, "Tasks")           # olFolderTasks
            ]
            
            for folder_id, folder_name in main_folders:
                try:
                    logger.debug(f"Retrieving folder: {folder_name} (ID: {folder_id})")
                    folder = namespace.GetDefaultFolder(folder_id)
                    
                    if folder:
                        folder_data = self._transform_folder_to_data_simple(folder, "")
                        if folder_data:
                            folders.append(folder_data)
                            logger.debug(f"Successfully added folder: {folder_data.name}")
                        
                except Exception as e:
                    logger.warning(f"Could not access {folder_name}: {str(e)}")
                    continue
            
            logger.debug(f"Retrieved {len(folders)} folders")
            return folders
            
        except Exception as e:
            logger.error(f"Error retrieving folders: {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError("", f"Access denied to folders: {str(e)}")
            raise OutlookConnectionError(f"Failed to retrieve folders: {str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def _transform_folder_to_data_simple(self, folder: Any, parent_path: str) -> Optional[FolderData]:
        """
        Transform a single Outlook folder to FolderData without recursion.
        
        Args:
            folder: The COM folder object
            parent_path: Parent folder path
            
        Returns:
            FolderData object or None if transformation fails
        """
        try:
            # Get basic folder properties
            folder_name = getattr(folder, 'Name', 'Unknown')
            folder_id = getattr(folder, 'EntryID', '')
            
            # Sanitize folder name for problematic characters
            if not folder_name or folder_name == 'Unknown':
                return None
                
            # Clean the folder name of any problematic characters
            clean_name = ''.join(c for c in folder_name if ord(c) >= 32)  # Remove control chars
            if not clean_name:
                clean_name = "Unknown_Folder"
            
            # Get item counts safely
            try:
                item_count = getattr(folder, 'Items', None)
                item_count = len(item_count) if item_count else 0
            except:
                item_count = 0
                
            try:
                unread_count = getattr(folder, 'UnReadItemCount', 0)
            except:
                unread_count = 0
            
            # Create folder data with safe values
            return FolderData(
                id=folder_id or f"folder_{clean_name}_{id(folder)}",
                name=clean_name,
                full_path=clean_name if not parent_path else f"{parent_path}/{clean_name}",
                item_count=max(0, item_count),
                unread_count=max(0, min(unread_count, item_count)),
                parent_folder=parent_path,
                folder_type="Mail"
            )
            
        except Exception as e:
            logger.warning(f"Could not transform folder: {str(e)}")
            return None

    def _collect_folders_recursive(self, folder: Any, folder_list: List[FolderData], parent_path: str) -> None:
        """
        Recursively collect folders and convert them to FolderData objects.
        
        Args:
            folder: The COM folder object to process
            folder_list: List to append FolderData objects to
            parent_path: Path of the parent folder
        """
        try:
            # Transform current folder to FolderData
            folder_data = self._transform_folder_to_data(folder, parent_path)
            folder_list.append(folder_data)
            
            # Process subfolders
            if hasattr(folder, 'Folders') and folder.Folders:
                current_path = folder_data.full_path
                for subfolder in folder.Folders:
                    try:
                        self._collect_folders_recursive(subfolder, folder_list, current_path)
                    except Exception as e:
                        logger.debug(f"Error processing subfolder: {str(e)}")
                        continue
                        
        except Exception as e:
            logger.debug(f"Error collecting folder data: {str(e)}")
    
    def _transform_folder_to_data(self, folder: Any, parent_path: str) -> FolderData:
        """
        Transform Outlook COM folder object to FolderData.
        
        Args:
            folder: The COM folder object
            parent_path: Path of the parent folder
            
        Returns:
            FolderData: Transformed folder data
        """
        try:
            # Get basic folder properties
            folder_id = getattr(folder, 'EntryID', '')
            folder_name = getattr(folder, 'Name', 'Unknown')
            
            # Ensure we have a valid ID for validation
            if not folder_id:
                folder_id = f"unknown_{folder_name}_{id(folder)}"
            
            # Build full path
            if parent_path:
                full_path = f"{parent_path}/{folder_name}"
            else:
                full_path = folder_name
            
            # Get item counts
            item_count = 0
            unread_count = 0
            
            try:
                if hasattr(folder, 'Items'):
                    item_count = folder.Items.Count
                if hasattr(folder, 'UnReadItemCount'):
                    unread_count = folder.UnReadItemCount
            except Exception as e:
                logger.debug(f"Error getting folder counts: {str(e)}")
            
            # Determine folder type
            folder_type = self._get_folder_type(folder)
            
            return FolderData(
                id=folder_id,
                name=folder_name,
                full_path=full_path,
                item_count=item_count,
                unread_count=unread_count,
                parent_folder=parent_path,
                folder_type=folder_type
            )
            
        except Exception as e:
            logger.error(f"Error transforming folder to data: {str(e)}")
            # Return minimal folder data if transformation fails
            folder_name = getattr(folder, 'Name', 'Unknown')
            folder_id = getattr(folder, 'EntryID', '')
            
            # Ensure we have a valid ID for validation
            if not folder_id:
                folder_id = f"unknown_{folder_name}_{id(folder)}"
            
            return FolderData(
                id=folder_id,
                name=folder_name,
                full_path=folder_name if not parent_path else f"{parent_path}/{folder_name}",
                item_count=0,
                unread_count=0,
                parent_folder=parent_path,
                folder_type="Unknown"
            )
    
    def _get_folder_type(self, folder: Any) -> str:
        """
        Determine the type of folder based on its properties.
        
        Args:
            folder: The COM folder object
            
        Returns:
            str: The folder type
        """
        try:
            # Check if folder has DefaultItemType property
            if hasattr(folder, 'DefaultItemType'):
                item_type = folder.DefaultItemType
                type_mapping = {
                    0: "Mail",      # olMailItem
                    1: "Contact",   # olContactItem
                    2: "Task",      # olTaskItem
                    3: "Journal",   # olJournalItem
                    4: "Note",      # olNoteItem
                    5: "Post"       # olPostItem
                }
                return type_mapping.get(item_type, "Mail")
            
            # Fallback: check folder name for common patterns
            folder_name = getattr(folder, 'Name', '').lower()
            if 'contact' in folder_name:
                return "Contact"
            elif 'calendar' in folder_name:
                return "Calendar"
            elif 'task' in folder_name:
                return "Task"
            elif 'note' in folder_name:
                return "Note"
            else:
                return "Mail"
                
        except Exception as e:
            logger.debug(f"Error determining folder type: {str(e)}")
            return "Mail"
    
    def validate_folder_access(self, folder_name: str) -> bool:
        """
        Validate that a folder exists and is accessible.
        
        Args:
            folder_name: Name of the folder to validate
            
        Returns:
            bool: True if folder is accessible, False otherwise
        """
        if not self.is_connected():
            return False
        
        try:
            folder = self.get_folder_by_name(folder_name)
            return folder is not None
        except (FolderNotFoundError, PermissionError):
            return False
        except Exception as e:
            logger.debug(f"Error validating folder access: {str(e)}")
            return False
    
    def get_default_folder_names(self) -> dict:
        """
        Get the actual names of default Outlook folders for debugging.
        
        Returns:
            dict: Mapping of folder IDs to actual folder names
        """
        if not self.is_connected():
            return {}
        
        folder_names = {}
        default_folder_ids = {
            6: "Inbox",
            5: "Sent Items", 
            4: "Outbox",
            3: "Deleted Items",
            16: "Drafts",
            23: "Junk Email",
            9: "Calendar",
            10: "Contacts",
            13: "Journal",
            12: "Tasks"
        }
        
        for folder_id, english_name in default_folder_ids.items():
            try:
                folder = self._namespace.GetDefaultFolder(folder_id)
                if folder and hasattr(folder, 'Name'):
                    actual_name = folder.Name
                    folder_names[folder_id] = {
                        'english_name': english_name,
                        'actual_name': actual_name,
                        'accessible': True
                    }
                    logger.debug(f"Default folder {folder_id} ({english_name}): '{actual_name}'")
                else:
                    folder_names[folder_id] = {
                        'english_name': english_name,
                        'actual_name': None,
                        'accessible': False
                    }
            except Exception as e:
                folder_names[folder_id] = {
                    'english_name': english_name,
                    'actual_name': None,
                    'accessible': False,
                    'error': str(e)
                }
                logger.debug(f"Error accessing folder {folder_id} ({english_name}): {e}")
        
        return folder_names
    
    def get_email_by_id(self, email_id: str) -> EmailData:
        """
        Get detailed email information by ID from Outlook.
        
        Args:
            email_id: Unique identifier of the email
            
        Returns:
            EmailData: Complete email data with detailed content
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            EmailNotFoundError: If email is not found or invalid ID
            PermissionError: If access to email is denied
        """
        logger.info(f"🔧 DEBUG: *** GET_EMAIL_BY_ID CALLED *** ID: {email_id[:50]}...")
        
        if not self.is_connected():
            logger.info(f"🔧 DEBUG: Not connected to Outlook")
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not email_id or not isinstance(email_id, str):
            logger.info(f"🔧 DEBUG: Invalid email_id: {email_id}")
            raise EmailNotFoundError(email_id or "")
        
        # Validate email ID format
        if not EmailData.validate_email_id(email_id):
            logger.info(f"🔧 DEBUG: Email ID validation failed: {email_id}")
            raise EmailNotFoundError(email_id)
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            logger.info(f"🔧 DEBUG: COM initialized for get_email_by_id")
            
            logger.debug(f"Retrieving detailed email with ID: {email_id}")
            logger.info(f"🔧 DEBUG: About to create thread-local Outlook connection")
            
            # Create thread-local Outlook connection (same as list_inbox_emails)
            try:
                logger.info(f"🔧 DEBUG: Creating thread-local Outlook connection")
                outlook_app = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook_app.GetNamespace("MAPI")
                
                logger.debug("Created thread-local Outlook connection for get_email_by_id")
                logger.info(f"🔧 DEBUG: Thread-local connection created successfully")
                
            except Exception as e:
                logger.error(f"Failed to create thread-local Outlook connection: {e}")
                logger.info(f"🔧 DEBUG: Falling back to original namespace")
                # Fall back to original namespace
                namespace = self._namespace
            
            # CRITICAL FIX: Instead of using GetItemFromID (which doesn't work properly),
            # find the email by searching through the inbox using the same method as list_inbox_emails
            logger.debug("Using list-based approach to find email (more reliable than GetItemFromID)")
            logger.info(f"🔧 DEBUG: Starting email search process")
            
            email_item = None
            
            # Get all emails from inbox using the working method
            try:
                logger.info(f"🔧 DEBUG: Getting namespace folders")
                # Get inbox folder using the same reliable method as list_inbox_emails
                try:
                    folders = namespace.Folders
                    logger.info(f"🔧 DEBUG: Got {folders.Count} folders")
                except Exception as e:
                    logger.info(f"🔧 DEBUG: Error accessing namespace.Folders: {e}")
                    raise e
                inbox_folder = None
                
                for f in folders:
                    try:
                        if hasattr(f, 'Items') and hasattr(f, 'Name'):
                            folder_name = f.Name
                            logger.info(f"🔧 DEBUG: Checking folder: {folder_name}")
                            
                            # Check if this is the inbox folder directly
                            if ("收件" in folder_name or "inbox" in folder_name.lower() or "æ¶ä»¶" in folder_name):
                                inbox_folder = f
                                logger.debug(f"Found inbox folder: {folder_name}")
                                logger.info(f"🔧 DEBUG: Found inbox folder: {folder_name}")
                                break
                            
                            # Check if this is a user mailbox folder that contains subfolders
                            elif "@" in folder_name and not folder_name.startswith("公用"):
                                logger.info(f"🔧 DEBUG: Checking subfolders in mailbox: {folder_name}")
                                try:
                                    subfolders = f.Folders
                                    logger.info(f"🔧 DEBUG: Found {subfolders.Count} subfolders in {folder_name}")
                                    for sf in subfolders:
                                        try:
                                            if hasattr(sf, 'Items') and hasattr(sf, 'Name'):
                                                subfolder_name = sf.Name
                                                logger.info(f"🔧 DEBUG: Checking subfolder: {subfolder_name}")
                                                if ("收件" in subfolder_name or "inbox" in subfolder_name.lower() or "æ¶ä»¶" in subfolder_name):
                                                    inbox_folder = sf
                                                    logger.debug(f"Found inbox subfolder: {subfolder_name}")
                                                    logger.info(f"🔧 DEBUG: Found inbox subfolder: {subfolder_name}")
                                                    break
                                        except Exception as e:
                                            logger.info(f"🔧 DEBUG: Error checking subfolder: {e}")
                                            continue
                                    if inbox_folder:
                                        break
                                except Exception as e:
                                    logger.info(f"🔧 DEBUG: Error accessing subfolders of {folder_name}: {e}")
                                    continue
                    except Exception as e:
                        logger.debug(f"Error checking folder: {e}")
                        logger.info(f"🔧 DEBUG: Error checking folder: {e}")
                        continue
                
                if not inbox_folder:
                    logger.info(f"🔧 DEBUG: No inbox folder found!")
                    raise FolderNotFoundError("Inbox")
                
                logger.info(f"🔧 DEBUG: Inbox folder found, proceeding with search")
                
                # Try Method 1: Use Outlook's GetItemFromID (most reliable)
                try:
                    logger.debug(f"Trying GetItemFromID for: {email_id[:50]}...")
                    email_item = namespace.GetItemFromID(email_id)
                    if email_item and hasattr(email_item, 'Class') and email_item.Class == 43:
                        logger.debug(f"✅ Found email using GetItemFromID")
                    else:
                        logger.debug(f"❌ GetItemFromID returned invalid item")
                        email_item = None
                except Exception as e:
                    logger.debug(f"❌ GetItemFromID failed: {e}")
                    email_item = None
                
                # Method 2: If GetItemFromID fails, search through inbox items
                if not email_item:
                    logger.debug(f"🔍 Falling back to inbox search...")
                    items = inbox_folder.Items
                    logger.debug(f"📧 Searching through {items.Count} inbox items...")
                    
                    checked_count = 0
                    for item in items:
                        try:
                            checked_count += 1
                            if checked_count <= 5:  # Log first 5 for debugging
                                item_id = getattr(item, 'EntryID', 'NO_ID')
                                logger.debug(f"   Item {checked_count}: {item_id[:50]}...")
                            
                            if (hasattr(item, 'EntryID') and 
                                hasattr(item, 'Class') and 
                                item.Class == 43 and  # olMail
                                item.EntryID == email_id):
                                email_item = item
                                logger.debug(f"✅ Found email by ID in inbox search (item {checked_count})")
                                break
                        except Exception as e:
                            logger.debug(f"Error checking item {checked_count}: {e}")
                            continue
                    
                    logger.debug(f"📊 Searched {checked_count} items, found: {'YES' if email_item else 'NO'}")
                        
            except Exception as e:
                logger.debug(f"Error finding email in inbox: {e}")
            
            if not email_item:
                logger.warning(f"Email not found: {email_id}")
                raise EmailNotFoundError(email_id)
            
            # Verify it's a mail item (type 43 = olMail)
            if not hasattr(email_item, 'Class') or email_item.Class != 43:
                logger.warning(f"Item is not a mail item: {email_id}")
                raise EmailNotFoundError(email_id)
            
            # Transform to EmailData using the same method as list_inbox_emails (which works)
            folder_name = "Unknown"
            try:
                # Try to get folder name from the email's parent folder
                parent_folder = getattr(email_item, 'Parent', None)
                if parent_folder:
                    folder_name = getattr(parent_folder, 'Name', 'Unknown')
            except Exception as e:
                logger.debug(f"Could not get folder name: {e}")
            
            email_data = self._transform_email_to_data(email_item, folder_name)
            
            logger.debug(f"Successfully retrieved detailed email: {email_id}")
            return email_data
            
        except EmailNotFoundError:
            raise
        except Exception as e:
            logger.error(f"Error retrieving email '{email_id}': {str(e)}")
            
            # Check for permission-related errors
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError(email_id, f"Access denied to email '{email_id}': {str(e)}")
            
            # Check for invalid ID errors
            if any(keyword in str(e).lower() for keyword in ['invalid', 'malformed', 'corrupt']):
                raise EmailNotFoundError(email_id)
            
            raise EmailNotFoundError(email_id)
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def _find_email_by_id_in_folders(self, email_id: str, namespace) -> Any:
        """
        Find an email by ID by searching through all folders.
        This is a fallback method when GetItemFromID fails.
        
        Args:
            email_id: The email ID to search for
            namespace: The Outlook namespace to search in
            
        Returns:
            Email item if found, None otherwise
        """
        try:
            logger.debug(f"Searching for email ID {email_id[:20]}... in all folders")
            
            # Get all folders
            folders = namespace.Folders
            
            for folder in folders:
                try:
                    if hasattr(folder, 'Items'):
                        items = folder.Items
                        
                        # Search through items in this folder (limit to first 100 for performance)
                        count = 0
                        for item in items:
                            if count >= 100:  # Limit search for performance
                                break
                            try:
                                if (hasattr(item, 'EntryID') and 
                                    hasattr(item, 'Class') and 
                                    item.Class == 43 and  # olMail
                                    item.EntryID == email_id):
                                    logger.debug(f"Found email in folder: {folder.Name}")
                                    return item
                                count += 1
                            except Exception as e:
                                logger.debug(f"Error checking item: {e}")
                                continue
                                
                except Exception as e:
                    logger.debug(f"Error searching folder: {e}")
                    continue
            
            logger.debug(f"Email ID {email_id[:20]}... not found in any folder")
            return None
            
        except Exception as e:
            logger.debug(f"Error in folder search: {e}")
            return None
    
    def search_emails(self, query: str, folder_identifier: str = None, limit: int = 50) -> List[EmailData]:
        """
        Search emails based on query with enhanced functionality.
        
        Args:
            query: Search query string
            folder_identifier: Optional folder name or ID to search in (searches all accessible folders if None)
            limit: Maximum number of results to return (default: 50)
            
        Returns:
            List[EmailData]: List of matching email data objects
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        # Validate input parameters
        if query is None or not isinstance(query, str):
            logger.debug("Empty or invalid query provided")
            return []
        
        if not query.strip():
            logger.debug("Empty or whitespace-only query provided")
            return []
        
        if limit <= 0:
            limit = 50  # Default limit
        
        try:
            logger.debug(f"Searching emails with query: '{query}', folder: {folder_identifier or 'all folders'}, limit: {limit}")
            
            # Process search query to ensure Outlook compatibility
            processed_query = self._process_search_query(query)
            
            if folder_identifier:
                # Search in specific folder (by name or ID)
                return self._search_in_folder(processed_query, folder_identifier, limit)
            else:
                # Search across all accessible folders
                return self._search_all_folders(processed_query, limit)
            
        except (FolderNotFoundError, PermissionError):
            raise
        except Exception as e:
            logger.error(f"Error searching emails: {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_identifier or "folders", f"Access denied during search: {str(e)}")
            return []
    
    def _process_search_query(self, query: str) -> str:
        """
        Process and validate search query for Outlook compatibility.
        
        Args:
            query: Raw search query
            
        Returns:
            Processed query string compatible with Outlook search
        """
        try:
            # Clean up the query
            processed_query = query.strip()
            
            # If query doesn't contain Outlook search syntax, make it search in subject and body
            if not any(keyword in processed_query.lower() for keyword in ['subject:', 'body:', 'from:', 'to:', 'received:']):
                # Search in subject and body by default
                processed_query = f'(subject:"{processed_query}" OR body:"{processed_query}")'
            
            logger.debug(f"Processed search query: {processed_query}")
            return processed_query
            
        except Exception as e:
            logger.debug(f"Error processing search query: {str(e)}")
            return query  # Return original query if processing fails
    
    def _search_in_folder(self, query: str, folder_identifier: str, limit: int) -> List[EmailData]:
        """
        Search emails in a specific folder.
        
        Args:
            query: Processed search query
            folder_identifier: Name or ID of the folder to search in
            limit: Maximum number of results to return
            
        Returns:
            List[EmailData]: List of matching email data objects
        """
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            # Create thread-local Outlook connection
            try:
                outlook_app = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook_app.GetNamespace("MAPI")
                
                logger.debug("Created thread-local Outlook connection for search")
                
            except Exception as e:
                logger.error(f"Failed to create thread-local Outlook connection: {e}")
                # Fall back to original namespace
                namespace = self._namespace
            
            # Get the target folder (by name or ID) using thread-local connection
            if len(folder_identifier) > 50 and all(c in '0123456789ABCDEFabcdef' for c in folder_identifier):
                # This looks like a folder ID
                folder = self._get_folder_by_id_thread_local(folder_identifier, namespace)
            else:
                # This is a folder name
                folder = self._get_folder_by_name_thread_local(folder_identifier, namespace)
            
            if not folder:
                raise FolderNotFoundError(folder_identifier)
            
            # Get the actual folder name for logging
            folder_name = getattr(folder, 'Name', folder_identifier)
            logger.debug(f"Searching in folder: {folder_name} (identifier: {folder_identifier})")
            
            # Perform search in the folder
            results = self._perform_folder_search(folder, query, limit, folder_name)
            
            logger.debug(f"Found {len(results)} emails in folder '{folder_name}'")
            return results
            
        except (FolderNotFoundError, PermissionError):
            raise
        except Exception as e:
            logger.error(f"Error searching in folder '{folder_identifier}': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_identifier, f"Access denied to folder '{folder_identifier}': {str(e)}")
            return []
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def _search_all_folders(self, query: str, limit: int) -> List[EmailData]:
        """
        Search emails across all accessible folders.
        
        Args:
            query: Processed search query
            limit: Maximum number of results to return
            
        Returns:
            List[EmailData]: List of matching email data objects from all folders
        """
        try:
            logger.debug("Searching across all accessible folders")
            
            all_results = []
            folders_searched = 0
            
            # Get all folders
            try:
                all_folders = self.get_folders()
            except Exception as e:
                logger.warning(f"Could not get all folders, searching in default folders: {str(e)}")
                # Fallback to searching in default folders
                return self._search_default_folders(query, limit)
            
            # Search in each folder
            for folder_data in all_folders:
                if len(all_results) >= limit:
                    break
                
                try:
                    # Skip non-mail folders
                    if folder_data.folder_type != "Mail":
                        continue
                    
                    # Get remaining limit for this folder
                    remaining_limit = limit - len(all_results)
                    
                    # Search in this folder
                    folder_results = self._search_in_folder(query, folder_data.name, remaining_limit)
                    all_results.extend(folder_results)
                    folders_searched += 1
                    
                    logger.debug(f"Searched folder '{folder_data.name}': {len(folder_results)} results")
                    
                except Exception as e:
                    logger.debug(f"Error searching folder '{folder_data.name}': {str(e)}")
                    continue
            
            # Sort results by received time (newest first)
            all_results.sort(key=lambda x: x.received_time or datetime.min, reverse=True)
            
            # Apply final limit
            final_results = all_results[:limit]
            
            logger.debug(f"Global search complete: {len(final_results)} results from {folders_searched} folders")
            return final_results
            
        except Exception as e:
            logger.error(f"Error in global search: {str(e)}")
            # Fallback to searching in default folders
            return self._search_default_folders(query, limit)
    
    def _search_default_folders(self, query: str, limit: int) -> List[EmailData]:
        """
        Search in default Outlook folders as fallback.
        
        Args:
            query: Processed search query
            limit: Maximum number of results to return
            
        Returns:
            List[EmailData]: List of matching email data objects from default folders
        """
        try:
            logger.debug("Searching in default folders as fallback")
            
            default_folders = ["Inbox", "Sent Items", "Drafts"]
            all_results = []
            
            for folder_name in default_folders:
                if len(all_results) >= limit:
                    break
                
                try:
                    remaining_limit = limit - len(all_results)
                    folder_results = self._search_in_folder(query, folder_name, remaining_limit)
                    all_results.extend(folder_results)
                    
                except Exception as e:
                    logger.debug(f"Error searching default folder '{folder_name}': {str(e)}")
                    continue
            
            # Sort results by received time (newest first)
            all_results.sort(key=lambda x: x.received_time or datetime.min, reverse=True)
            
            return all_results[:limit]
            
        except Exception as e:
            logger.error(f"Error searching default folders: {str(e)}")
            return []
    
    def _perform_folder_search(self, folder: Any, query: str, limit: int, folder_name: str) -> List[EmailData]:
        """
        Perform the actual search operation in a folder.
        
        Args:
            folder: COM folder object to search in
            query: Search query string
            limit: Maximum number of results to return
            folder_name: Name of the folder (for logging and data)
            
        Returns:
            List[EmailData]: List of matching email data objects
        """
        try:
            # Get folder items
            items = folder.Items
            
            # Sort by received time (newest first) for consistent results
            items.Sort("[ReceivedTime]", True)  # True for descending order
            
            results = []
            count = 0
            
            # Parse the search query to extract search terms and fields
            search_terms = self._parse_search_query(query)
            
            logger.debug(f"Searching folder '{folder_name}' with terms: {search_terms}")
            
            # Manual search through items (more reliable than Outlook's Find method)
            items_checked = 0
            max_items_to_check = min(1000, items.Count)  # Limit for performance
            
            for item in items:
                if count >= limit or items_checked >= max_items_to_check:
                    break
                
                items_checked += 1
                
                try:
                    # Check if it's a mail item (type 43 = olMail)
                    if not hasattr(item, 'Class') or item.Class != 43:
                        continue
                    
                    # Check if item matches search criteria
                    if self._item_matches_search(item, search_terms):
                        # Transform to EmailData
                        email_data = self._transform_email_to_data(item, folder_name)
                        results.append(email_data)
                        count += 1
                    
                except Exception as e:
                    logger.debug(f"Error processing item in search: {str(e)}")
                    continue
            
            logger.debug(f"Search in folder '{folder_name}' checked {items_checked} items, found {len(results)} matches")
            return results
            
        except Exception as e:
            logger.error(f"Error performing search in folder '{folder_name}': {str(e)}")
            return []
    
    def _parse_search_query(self, query: str) -> dict:
        """
        Parse search query into searchable terms and fields.
        
        Args:
            query: Search query string
            
        Returns:
            Dictionary with search terms and target fields
        """
        search_terms = {
            'subject': [],
            'body': [],
            'from': [],
            'to': [],
            'general': []
        }
        
        try:
            # Split query by common operators
            query = query.lower().strip()
            
            # Handle subject: queries
            if 'subject:' in query:
                import re
                subject_matches = re.findall(r'subject:(["\']?)([^"\'\s]+)\1', query)
                for match in subject_matches:
                    search_terms['subject'].append(match[1])
            
            # Handle from: queries
            if 'from:' in query:
                import re
                from_matches = re.findall(r'from:(["\']?)([^"\'\s]+)\1', query)
                for match in from_matches:
                    search_terms['from'].append(match[1])
            
            # Handle body: queries
            if 'body:' in query:
                import re
                body_matches = re.findall(r'body:(["\']?)([^"\'\s]+)\1', query)
                for match in body_matches:
                    search_terms['body'].append(match[1])
            
            # If no specific field queries, treat as general search
            if not any(search_terms[field] for field in ['subject', 'body', 'from']):
                # Remove field prefixes and use as general search
                import re
                clean_query = re.sub(r'(subject:|body:|from:|to:)', '', query).strip()
                if clean_query:
                    search_terms['general'].append(clean_query)
            
            logger.debug(f"Parsed search terms: {search_terms}")
            return search_terms
            
        except Exception as e:
            logger.debug(f"Error parsing search query: {e}")
            # Fallback: treat entire query as general search
            return {'subject': [], 'body': [], 'from': [], 'to': [], 'general': [query.lower()]}
    
    def _item_matches_search(self, item: Any, search_terms: dict) -> bool:
        """
        Check if an email item matches the search criteria.
        
        Args:
            item: Outlook email item
            search_terms: Parsed search terms
            
        Returns:
            True if item matches search criteria
        """
        try:
            # Get item properties safely
            subject = str(getattr(item, 'Subject', '')).lower()
            sender_name = str(getattr(item, 'SenderName', '')).lower()
            sender_email = str(getattr(item, 'SenderEmailAddress', '')).lower()
            
            # For performance, don't load body unless specifically searching for it
            body = ""
            if search_terms['body']:
                body = str(getattr(item, 'Body', '')).lower()
            
            logger.debug(f"Checking item: subject='{subject[:50]}...', sender='{sender_name}'")
            
            # Check subject matches
            for term in search_terms['subject']:
                if term.lower() in subject:
                    logger.debug(f"Subject match found: '{term}' in '{subject[:50]}...'")
                    return True
            
            # Check body matches (only if body terms exist)
            if search_terms['body']:
                for term in search_terms['body']:
                    if term.lower() in body:
                        logger.debug(f"Body match found: '{term}'")
                        return True
            
            # Check from matches
            for term in search_terms['from']:
                if term.lower() in sender_name or term.lower() in sender_email:
                    logger.debug(f"From match found: '{term}' in '{sender_name}' or '{sender_email}'")
                    return True
            
            # Check general matches (search in subject and sender, skip body for performance)
            for term in search_terms['general']:
                if (term.lower() in subject or 
                    term.lower() in sender_name or 
                    term.lower() in sender_email):
                    logger.debug(f"General match found: '{term}' in subject or sender")
                    return True
            
            return False
            
        except Exception as e:
            logger.debug(f"Error checking item match: {e}")
            return False
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def list_inbox_emails(self, unread_only: bool = False, limit: int = 50) -> List[EmailData]:
        """
        List emails from the default inbox folder.
        
        Args:
            unread_only: If True, only return unread emails
            limit: Maximum number of emails to return
            
        Returns:
            List[EmailData]: List of email data objects
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            PermissionError: If access to inbox is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if limit <= 0:
            limit = 50  # Default limit
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            logger.debug(f"Listing emails from inbox, unread_only: {unread_only}, limit: {limit}")
            
            # Create thread-local Outlook connection
            try:
                outlook_app = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook_app.GetNamespace("MAPI")
                
                logger.debug("Created thread-local Outlook connection for list_inbox_emails")
                
            except Exception as e:
                logger.error(f"Failed to create thread-local Outlook connection: {e}")
                # Fall back to original namespace
                namespace = self._namespace
            
            # Get the inbox folder using the EXACT same method as list_emails
            # Find the inbox folder by looking for the folder with Chinese name "收件匣"
            folder = None
            
            # Search through all folders to find the inbox (same as list_emails logic)
            try:
                folders = namespace.Folders
                max_items = 0
                inbox_candidate = None
                
                for f in folders:
                    try:
                        if hasattr(f, 'Items') and hasattr(f, 'Name'):
                            item_count = f.Items.Count
                            folder_name = f.Name
                            
                            logger.debug(f"Checking folder: {folder_name} with {item_count} items")
                            
                            # Look for inbox-like names (Chinese, English, etc.)
                            if ("收件" in folder_name or "inbox" in folder_name.lower() or 
                                "æ¶ä»¶" in folder_name):
                                logger.debug(f"Found inbox by name: {folder_name}")
                                folder = f
                                break
                            
                            # Also track the folder with most items as backup
                            if item_count > max_items:
                                max_items = item_count
                                inbox_candidate = f
                                
                    except Exception as e:
                        logger.debug(f"Error checking folder: {e}")
                        continue
                
                # If we didn't find by name, use the folder with most items
                if not folder and inbox_candidate:
                    folder = inbox_candidate
                    logger.debug(f"Using folder with most items as inbox: {inbox_candidate.Name} ({max_items} items)")
                    
            except Exception as e:
                logger.debug(f"Error searching folders: {e}")
            
            # Final fallback: try GetDefaultFolder
            if not folder:
                logger.debug("Trying GetDefaultFolder as final fallback")
                try:
                    folder = namespace.GetDefaultFolder(6)  # olFolderInbox
                except Exception as e:
                    logger.debug(f"GetDefaultFolder failed: {e}")
            
            if not folder:
                raise FolderNotFoundError("Inbox")
            
            # Get folder items
            items = folder.Items
            
            # Sort by received time (newest first)
            items.Sort("[ReceivedTime]", True)  # True for descending order
            
            emails = []
            count = 0
            
            # Iterate through items and apply filters
            for item in items:
                if count >= limit:
                    break
                
                try:
                    # Check if it's a mail item (type 43 = olMail)
                    if not hasattr(item, 'Class') or item.Class != 43:
                        continue
                    
                    # Apply unread filter if specified
                    if unread_only and getattr(item, 'UnRead', True) is False:
                        continue
                    
                    # Transform email to EmailData
                    folder_name = getattr(folder, 'Name', "Inbox")
                    email_data = self._transform_email_to_data(item, folder_name)
                    emails.append(email_data)
                    count += 1
                    
                except Exception as e:
                    logger.debug(f"Error processing email item: {str(e)}")
                    continue
            
            logger.debug(f"Retrieved {len(emails)} emails from folder")
            return emails
            
        except (FolderNotFoundError, PermissionError):
            raise
        except Exception as e:
            logger.error(f"Error listing emails from inbox: {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError("Inbox", f"Access denied to inbox: {str(e)}")
            raise OutlookConnectionError(f"Failed to list emails: {str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def list_emails(self, folder_id: str, unread_only: bool = False, limit: int = 50) -> List[EmailData]:
        """
        List emails from a specific folder by folder ID.
        
        Args:
            folder_id: ID of the folder to list emails from
            unread_only: If True, only return unread emails
            limit: Maximum number of emails to return
            
        Returns:
            List[EmailData]: List of email data objects
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
            ValidationError: If folder_id is invalid
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not folder_id or not isinstance(folder_id, str):
            raise ValidationError("folder_id must be a non-empty string", "folder_id")
        
        if limit <= 0:
            limit = 50  # Default limit
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            logger.debug(f"Listing emails from folder ID: {folder_id[:20]}..., unread_only: {unread_only}, limit: {limit}")
            
            # Create thread-local Outlook connection
            try:
                outlook_app = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook_app.GetNamespace("MAPI")
                
                logger.debug("Created thread-local Outlook connection for list_emails")
                
            except Exception as e:
                logger.error(f"Failed to create thread-local Outlook connection: {e}")
                # Fall back to original namespace
                namespace = self._namespace
            
            # Get the folder by ID
            folder = self._get_folder_by_id_thread_local(folder_id, namespace)
            
            if not folder:
                raise FolderNotFoundError(folder_id)
            
            # Get folder items
            items = folder.Items
            
            # Sort by received time (newest first)
            items.Sort("[ReceivedTime]", True)  # True for descending order
            
            emails = []
            count = 0
            
            # Iterate through items and apply filters
            for item in items:
                if count >= limit:
                    break
                
                try:
                    # Check if it's a mail item (type 43 = olMail)
                    if not hasattr(item, 'Class') or item.Class != 43:
                        continue
                    
                    # Apply unread filter if specified
                    if unread_only and getattr(item, 'UnRead', True) is False:
                        continue
                    
                    # Transform email to EmailData
                    folder_name = getattr(folder, 'Name', folder_id)
                    email_data = self._transform_email_to_data(item, folder_name)
                    emails.append(email_data)
                    count += 1
                    
                except Exception as e:
                    logger.debug(f"Error processing email item: {str(e)}")
                    continue
            
            logger.debug(f"Retrieved {len(emails)} emails from folder ID: {folder_id[:20]}...")
            return emails
            
        except (FolderNotFoundError, PermissionError, ValidationError):
            raise
        except Exception as e:
            logger.error(f"Error listing emails from folder ID '{folder_id[:20]}...': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_id, f"Access denied to folder: {str(e)}")
            raise OutlookConnectionError(f"Failed to list emails: {str(e)}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def _get_folder_by_id_thread_local(self, folder_id: str, namespace: Any) -> Any:
        """
        Get folder by ID using thread-local namespace.
        
        Args:
            folder_id: ID of the folder to retrieve
            namespace: Thread-local namespace object
            
        Returns:
            COM object: The folder object
        """
        try:
            logger.debug(f"Looking for folder by ID (thread-local): {folder_id[:20]}...")
            
            # Check main default folders
            main_folders = [
                (6, "Inbox"),           # olFolderInbox
                (5, "Sent Items"),      # olFolderSentMail  
                (16, "Drafts"),         # olFolderDrafts
                (3, "Deleted Items"),   # olFolderDeletedItems
                (4, "Outbox"),          # olFolderOutbox
                (9, "Calendar"),        # olFolderCalendar
                (10, "Contacts"),       # olFolderContacts
                (13, "Journal"),        # olFolderJournal
                (12, "Tasks")           # olFolderTasks
            ]
            
            for folder_id_num, folder_name in main_folders:
                try:
                    folder = namespace.GetDefaultFolder(folder_id_num)
                    if folder and hasattr(folder, 'EntryID'):
                        if folder.EntryID == folder_id:
                            logger.debug(f"Found folder by ID: {folder_id[:20]}... -> {folder_name}")
                            return folder
                except Exception as e:
                    logger.debug(f"Error checking folder {folder_name}: {e}")
                    continue
            
            logger.warning(f"Folder not found by ID: {folder_id[:20]}...")
            raise FolderNotFoundError(folder_id)
            
        except Exception as e:
            logger.error(f"Error in thread-local folder ID lookup: {e}")
            raise FolderNotFoundError(folder_id)
    
    def _get_folder_by_name_thread_local(self, folder_name: str, namespace: Any) -> Any:
        """
        Get folder by name using thread-local namespace.
        
        Args:
            folder_name: Name of the folder to retrieve
            namespace: Thread-local namespace object
            
        Returns:
            COM object: The folder object
        """
        try:
            logger.debug(f"Looking for folder by name (thread-local): {folder_name}")
            
            # Try default folders first
            default_folders = {
                "Inbox": 6, "收件匣": 6,
                "Sent Items": 5, "寄件備份": 5, "已傳送的郵件": 5,
                "Drafts": 16, "草稿": 16,
                "Deleted Items": 3, "已刪除的項目": 3,
                "Outbox": 4, "寄件匣": 4
            }
            
            if folder_name in default_folders:
                folder_id = default_folders[folder_name]
                folder = namespace.GetDefaultFolder(folder_id)
                if folder:
                    logger.debug(f"Found folder by name: {folder_name}")
                    return folder
            
            logger.warning(f"Folder not found by name: {folder_name}")
            raise FolderNotFoundError(folder_name)
            
        except Exception as e:
            logger.error(f"Error in thread-local folder name lookup: {e}")
            raise FolderNotFoundError(folder_name)
    
    def _transform_email_to_data(self, email_item: Any, folder_name: str) -> EmailData:
        """
        Transform Outlook COM email object to EmailData.
        
        Args:
            email_item: The COM email object
            folder_name: Name of the folder containing the email
            
        Returns:
            EmailData: Transformed email data
        """
        try:
            logger.info(f"🔧 DEBUG: *** TRANSFORM EMAIL TO DATA CALLED *** folder: {folder_name}")
            logger.info(f"🔧 DEBUG: Starting email transformation for folder: {folder_name}")
            
            # Force COM object to load all properties by accessing them in a specific way
            # This is critical to ensure all email properties are accessible
            try:
                # Force the COM object to fully initialize by accessing key properties
                _ = email_item.Class  # Force COM object initialization
                _ = email_item.MessageClass  # Force message type loading
                logger.info(f"🔧 DEBUG: COM object initialized successfully")
            except Exception as com_e:
                logger.info(f"🔧 DEBUG: COM object initialization failed: {com_e}")
            
            # Get basic email properties first
            email_id = self._get_email_property(email_item, 'EntryID', f"unknown_{id(email_item)}")
            subject = self._get_email_property(email_item, 'Subject', '')
            logger.info(f"🔧 DEBUG: Basic properties - ID: {email_id[:20]}..., Subject: '{subject[:50]}'")
            
            # Get sender information
            sender_name = self._get_email_property(email_item, 'SenderName', '')
            sender_email = self._get_email_property(email_item, 'SenderEmailAddress', '')
            logger.info(f"🔧 DEBUG: Raw sender - Name: '{sender_name}', Email: '{sender_email}'")
            logger.info(f"🔧 DEBUG: About to continue with property loading...")
            
            # Force property loading by trying to access them
            try:
                _ = email_item.Subject
                _ = email_item.Body
                _ = email_item.SenderName
            except Exception as e:
                logger.debug(f"Error forcing COM property access: {e}")
            
            logger.info(f"🔧 DEBUG: Property loading completed, continuing with email processing...")
            
            # Get basic email properties with FORCED property access
            sender_name = ''
            
            try:
                # Use direct property access with error handling
                email_id = str(email_item.EntryID) if hasattr(email_item, 'EntryID') else ''
            except Exception as e:
                logger.debug(f"Error getting EntryID: {e}")
                email_id = f"unknown_{id(email_item)}"
            
            try:
                subject = str(email_item.Subject) if hasattr(email_item, 'Subject') else ''
            except Exception as e:
                logger.debug(f"Error getting Subject: {e}")
                subject = '(No Subject)'
            
            try:
                sender_name = str(email_item.SenderName) if hasattr(email_item, 'SenderName') else ''
            except Exception as e:
                logger.debug(f"Error getting SenderName: {e}")
                sender_name = 'Unknown Sender'
            
            # Try multiple ways to get sender email address with FORCED access
            sender_email = ''
            logger.info(f"🔧 DEBUG: Starting sender email extraction methods...")
            try:
                # Method 1: Direct property with forced access
                try:
                    if hasattr(email_item, 'SenderEmailAddress'):
                        sender_email = str(email_item.SenderEmailAddress)
                        logger.info(f"🔧 DEBUG: Method 1 - SenderEmailAddress: '{sender_email}'")
                except Exception as e:
                    logger.info(f"🔧 DEBUG: Method 1 failed: {e}")
                
                # Method 2: If empty or invalid, try Sender property
                if not sender_email or '@' not in sender_email:
                    logger.info(f"🔧 DEBUG: Trying Method 2 - Sender.Address...")
                    try:
                        if hasattr(email_item, 'Sender'):
                            sender_obj = email_item.Sender
                            if sender_obj and hasattr(sender_obj, 'Address'):
                                sender_email = str(sender_obj.Address)
                                logger.info(f"🔧 DEBUG: Method 2 - Sender.Address: '{sender_email}'")
                    except Exception as e:
                        logger.debug(f"Method 2 failed: {e}")
                
                # Method 3: Try Reply Recipients
                if not sender_email or '@' not in sender_email:
                    try:
                        if hasattr(email_item, 'ReplyRecipients'):
                            reply_recipients = email_item.ReplyRecipients
                            if reply_recipients and reply_recipients.Count > 0:
                                sender_email = str(reply_recipients.Item(1).Address)
                                logger.debug(f"Method 3 - ReplyRecipients: '{sender_email}'")
                    except Exception as e:
                        logger.debug(f"Method 3 failed: {e}")
                        
            except Exception as e:
                logger.debug(f"Error getting sender email: {e}")
                sender_email = ''
            
            # Ensure we have a valid ID
            if not email_id:
                email_id = f"unknown_{id(email_item)}"
            
            # Get recipient information
            recipients = []
            cc_recipients = []
            bcc_recipients = []
            
            logger.info(f"🔧 DEBUG: Starting recipient extraction (duplicate section)...")
            try:
                if hasattr(email_item, 'Recipients'):
                    logger.info(f"🔧 DEBUG: Found {email_item.Recipients.Count} recipients in duplicate section")
                    for recipient in email_item.Recipients:
                        recipient_email = getattr(recipient, 'Address', '')
                        recipient_name = getattr(recipient, 'Name', '')
                        recipient_type = getattr(recipient, 'Type', 1)  # 1=To, 2=CC, 3=BCC
                        
                        logger.info(f"🔧 DEBUG: Processing recipient (duplicate) - Email: '{recipient_email}', Name: '{recipient_name}', Type: {recipient_type}")
                        
                        # Handle Exchange internal addresses
                        if recipient_email and (recipient_email.startswith('/o=') or recipient_email.startswith('/O=')):
                            logger.info(f"🔧 DEBUG: Found Exchange recipient address (duplicate): '{recipient_email}' - converting...")
                            # Convert Exchange address to valid format
                            if recipient_name and recipient_name != recipient_email:
                                clean_name = ''.join(c for c in recipient_name if c.isalnum() or c in '._-')
                                recipient_email = f"{clean_name}@internal.exchange" if clean_name else "unknown@internal.exchange"
                                logger.info(f"🔧 DEBUG: Converted recipient using name: '{recipient_email}'")
                            else:
                                # Extract CN from Exchange address
                                if 'cn=' in recipient_email.lower():
                                    try:
                                        cn_part = recipient_email.lower().split('cn=')[-1].split('/')[0].split('-')[0]
                                        recipient_email = f"{cn_part}@internal.exchange"
                                        logger.info(f"🔧 DEBUG: Converted recipient using CN: '{recipient_email}'")
                                    except:
                                        recipient_email = "unknown@internal.exchange"
                                        logger.info(f"🔧 DEBUG: Recipient conversion failed, using default: '{recipient_email}'")
                                else:
                                    recipient_email = "unknown@internal.exchange"
                                    logger.info(f"🔧 DEBUG: No CN found in recipient, using default: '{recipient_email}'")
                        
                        # Only add if we have a valid email format
                        if recipient_email and '@' in recipient_email:
                            if recipient_type == 1:  # To
                                recipients.append(recipient_email)
                            elif recipient_type == 2:  # CC
                                cc_recipients.append(recipient_email)
                            elif recipient_type == 3:  # BCC
                                bcc_recipients.append(recipient_email)
                        else:
                            logger.info(f"🔧 DEBUG: Skipping invalid recipient email: '{recipient_email}'")
            except Exception as e:
                logger.debug(f"Error processing recipients: {str(e)}")
            
            # Get email body with SIMPLIFIED and RELIABLE extraction
            body = ''
            body_html = ''
            
            try:
                logger.debug("Starting body extraction...")
                
                # Method 1: Simple direct access to Body property
                try:
                    if hasattr(email_item, 'Body'):
                        body_raw = email_item.Body
                        if body_raw is not None:
                            body = str(body_raw).strip()
                            if body:
                                logger.debug(f"Body extraction SUCCESS: {len(body)} chars")
                            else:
                                logger.debug("Body property exists but is empty")
                        else:
                            logger.debug("Body property is None")
                    else:
                        logger.debug("Email item has no Body property")
                except Exception as e:
                    logger.debug(f"Body access failed: {e}")
                
                # Method 2: Simple direct access to HTMLBody property
                try:
                    if hasattr(email_item, 'HTMLBody'):
                        html_body_raw = email_item.HTMLBody
                        if html_body_raw is not None:
                            body_html = str(html_body_raw).strip()
                            if body_html:
                                logger.debug(f"HTMLBody extraction SUCCESS: {len(body_html)} chars")
                                
                                # If we have HTML but no plain text, extract it now
                                if not body:
                                    try:
                                        import re
                                        # Extract text from HTML
                                        text_from_html = re.sub(r'<[^>]+>', '', body_html)
                                        text_from_html = text_from_html.replace('&nbsp;', ' ')
                                        text_from_html = text_from_html.replace('&lt;', '<')
                                        text_from_html = text_from_html.replace('&gt;', '>')
                                        text_from_html = text_from_html.replace('&amp;', '&')
                                        text_from_html = text_from_html.strip()
                                        
                                        if text_from_html:
                                            body = text_from_html
                                            logger.debug(f"Extracted text from HTML: {len(body)} chars")
                                    except Exception as e:
                                        logger.debug(f"Failed to extract text from HTML: {e}")
                            else:
                                logger.debug("HTMLBody property exists but is empty")
                        else:
                            logger.debug("HTMLBody property is None")
                    else:
                        logger.debug("Email item has no HTMLBody property")
                except Exception as e:
                    logger.debug(f"HTMLBody access failed: {e}")
                
                # Method 3: If both are still empty, try simple COM object refresh
                if not body and not body_html:
                    logger.debug("Both body properties empty, trying simple refresh...")
                    try:
                        # Simple refresh by accessing Size property
                        _ = email_item.Size
                        
                        # Try body again after refresh
                        if hasattr(email_item, 'Body'):
                            body_raw = email_item.Body
                            if body_raw:
                                body = str(body_raw).strip()
                                logger.debug(f"Post-refresh Body: {len(body)} chars")
                        
                        if hasattr(email_item, 'HTMLBody'):
                            html_body_raw = email_item.HTMLBody
                            if html_body_raw:
                                body_html = str(html_body_raw).strip()
                                logger.debug(f"Post-refresh HTMLBody: {len(body_html)} chars")
                                
                    except Exception as e:
                        logger.debug(f"Simple refresh failed: {e}")
                

                
                # Final check: Log if body extraction failed
                if not body and not body_html:
                    try:
                        message_class = str(email_item.MessageClass) if hasattr(email_item, 'MessageClass') else 'Unknown'
                        email_size = getattr(email_item, 'Size', 0)
                        
                        logger.warning(f"Body extraction failed for: '{subject[:50]}...', MessageClass: {message_class}, Size: {email_size}")
                        
                        # Don't set placeholder text - leave body empty so we can identify the real issue
                        
                    except Exception as e:
                        logger.debug(f"Final check failed: {e}")
                
            except Exception as e:
                logger.error(f"CRITICAL ERROR in body extraction: {e}")
                body = "[CRITICAL BODY EXTRACTION ERROR]"
                body_html = ""
            
            # Get timestamps
            received_time = None
            sent_time = None
            
            try:
                if hasattr(email_item, 'ReceivedTime'):
                    received_time = email_item.ReceivedTime
                if hasattr(email_item, 'SentOn'):
                    sent_time = email_item.SentOn
            except Exception as e:
                logger.debug(f"Error processing timestamps: {str(e)}")
            
            # Get other properties
            is_read = not getattr(email_item, 'UnRead', True)
            has_attachments = getattr(email_item, 'Attachments', Mock()).Count > 0 if hasattr(email_item, 'Attachments') else False
            
            # Get importance
            importance_value = getattr(email_item, 'Importance', 1)  # 0=Low, 1=Normal, 2=High
            importance_map = {0: "Low", 1: "Normal", 2: "High"}
            importance = importance_map.get(importance_value, "Normal")
            
            # Get size
            size = getattr(email_item, 'Size', 0)
            
            # Ensure we have a subject
            if not subject:
                subject = "(No Subject)"
            
            # Only use fallback email if we truly can't find the real one
            if not sender_email or '@' not in sender_email:
                logger.debug(f"Could not find valid sender email for: {subject[:50]}...")
                if sender_name:
                    # Create a more recognizable placeholder
                    clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                    sender_email = f"{clean_name}@email-not-available.com" if clean_name else "unknown@email-not-available.com"
                else:
                    sender_email = "unknown@email-not-available.com"
            
            # Ensure we have a sender name
            if not sender_name:
                sender_name = sender_email.split('@')[0] if '@' in sender_email else "Unknown Sender"
            
            return EmailData(
                id=email_id,
                subject=subject,
                sender=sender_name,
                sender_email=sender_email,
                recipients=recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body=body,
                body_html=body_html,
                received_time=received_time,
                sent_time=sent_time,
                is_read=is_read,
                has_attachments=has_attachments,
                importance=importance,
                folder_name=folder_name,
                size=size
            )
            
        except Exception as e:
            logger.error(f"Error transforming email to data: {str(e)}")
            logger.info(f"🔧 DEBUG: Exception handler - creating minimal email data")
            
            # Return minimal email data if transformation fails
            email_id = getattr(email_item, 'EntryID', f"unknown_{id(email_item)}")
            subject = getattr(email_item, 'Subject', '(No Subject)')
            sender_name = getattr(email_item, 'SenderName', 'Unknown Sender')
            sender_email = getattr(email_item, 'SenderEmailAddress', 'unknown@unknown.com')
            
            logger.info(f"🔧 DEBUG: Exception handler - Raw sender: Name='{sender_name}', Email='{sender_email}'")
            
            # Handle Exchange addresses in error handler too
            if sender_email and (sender_email.startswith('/o=') or sender_email.startswith('/O=')):
                logger.info(f"🔧 DEBUG: Exception handler - Found Exchange address, converting...")
                if sender_name and sender_name != sender_email:
                    clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                    sender_email = f"{clean_name}@internal.exchange" if clean_name else "unknown@internal.exchange"
                else:
                    sender_email = "unknown@internal.exchange"
                logger.info(f"🔧 DEBUG: Exception handler - Converted to: '{sender_email}'")
            
            # Ensure valid email format
            elif not sender_email or '@' not in sender_email:
                if sender_name and sender_name != 'Unknown Sender':
                    # Remove any invalid characters from sender name for email
                    clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                    sender_email = f"{clean_name}@unknown.com" if clean_name else "unknown@unknown.com"
                else:
                    sender_email = "unknown@unknown.com"
            
            return EmailData(
                id=email_id,
                subject=subject,
                sender=sender_name,
                sender_email=sender_email,
                recipients=[],
                cc_recipients=[],
                bcc_recipients=[],
                body="",
                body_html="",
                received_time=None,
                sent_time=None,
                is_read=False,
                has_attachments=False,
                importance="Normal",
                folder_name=folder_name,
                size=0
            )

    def _transform_email_to_detailed_data(self, email_item: Any) -> EmailData:
        """
        Transform Outlook COM email object to detailed EmailData with enhanced content extraction.
        
        Args:
            email_item: The COM email object
            
        Returns:
            EmailData: Detailed email data with full content and metadata
        """
        try:
            logger.debug("Transforming email to detailed data")
            
            # Get basic email properties with enhanced error handling
            email_id = self._get_email_property(email_item, 'EntryID', '')
            subject = self._get_email_property(email_item, 'Subject', '')
            sender_name = self._get_email_property(email_item, 'SenderName', '')
            sender_email = self._get_email_property(email_item, 'SenderEmailAddress', '')
            
            # Ensure we have a subject
            if not subject:
                subject = "(No Subject)"
            
            # Ensure we have a valid ID
            if not email_id:
                email_id = f"unknown_{id(email_item)}"
            
            # Get enhanced recipient information
            logger.info(f"🔧 DEBUG: About to extract recipients...")
            try:
                recipients, cc_recipients, bcc_recipients = self._extract_recipients(email_item)
                logger.info(f"🔧 DEBUG: Recipients extracted - To: {len(recipients)}, CC: {len(cc_recipients)}, BCC: {len(bcc_recipients)}")
            except Exception as recipient_error:
                logger.info(f"🔧 DEBUG: Recipient extraction failed: {recipient_error}")
                recipients, cc_recipients, bcc_recipients = [], [], []
            
            # Get email body with enhanced content extraction
            logger.info(f"🔧 DEBUG: About to extract body...")
            body, body_html = self._extract_email_body(email_item)
            logger.info(f"🔧 DEBUG: Body extracted - Text: {len(body)}, HTML: {len(body_html)}")
            
            # Additional fallback: if we have HTML but no plain text, extract text from HTML
            if body_html and not body:
                try:
                    body = self._extract_text_from_html(body_html)
                    logger.debug(f"Extracted text from HTML: {len(body)} chars")
                except Exception as e:
                    logger.debug(f"Failed to extract text from HTML: {e}")
            
            # Get timestamps with proper handling
            received_time, sent_time = self._extract_timestamps(email_item)
            
            # Get email status and properties
            is_read = not self._get_email_property(email_item, 'UnRead', True)
            
            # Get attachment information with detailed handling
            has_attachments, attachment_count = self._extract_attachment_info(email_item)
            
            # Get importance with proper mapping
            importance = self._extract_importance(email_item)
            
            # Get email size
            size = self._get_email_property(email_item, 'Size', 0)
            
            # Determine folder name from parent folder
            folder_name = self._extract_folder_name(email_item)
            
            # Validate and fix sender information
            logger.info(f"🔧 DEBUG: About to validate sender info...")
            sender_name, sender_email = self._validate_sender_info(sender_name, sender_email)
            logger.info(f"🔧 DEBUG: Sender validated - Name: '{sender_name}', Email: '{sender_email}'")
            
            # Create EmailData with all extracted information
            logger.info(f"🔧 DEBUG: About to create EmailData object...")
            logger.info(f"🔧 DEBUG: Recipients for EmailData - To: {recipients}, CC: {cc_recipients}, BCC: {bcc_recipients}")
            
            email_data = EmailData(
                id=email_id,
                subject=subject,
                sender=sender_name,
                sender_email=sender_email,
                recipients=recipients,
                cc_recipients=cc_recipients,
                bcc_recipients=bcc_recipients,
                body=body,
                body_html=body_html,
                received_time=received_time,
                sent_time=sent_time,
                is_read=is_read,
                has_attachments=has_attachments,
                importance=importance,
                folder_name=folder_name,
                size=size
            )
            
            logger.debug(f"Successfully transformed email to detailed data: {email_id}")
            return email_data
            
        except Exception as e:
            logger.error(f"Error transforming email to detailed data: {str(e)}")
            
            # Return minimal email data if detailed transformation fails
            email_id = self._get_email_property(email_item, 'EntryID', f"unknown_{id(email_item)}")
            subject = self._get_email_property(email_item, 'Subject', '(No Subject)')
            sender_name = self._get_email_property(email_item, 'SenderName', 'Unknown Sender')
            sender_email = self._get_email_property(email_item, 'SenderEmailAddress', 'unknown@unknown.com')
            
            # Ensure valid email format
            sender_name, sender_email = self._validate_sender_info(sender_name, sender_email)
            
            return EmailData(
                id=email_id,
                subject=subject,
                sender=sender_name,
                sender_email=sender_email,
                recipients=[],
                cc_recipients=[],
                bcc_recipients=[],
                body="",
                body_html="",
                received_time=None,
                sent_time=None,
                is_read=False,
                has_attachments=False,
                importance="Normal",
                folder_name="Unknown",
                size=0
            )
    
    def _get_email_property(self, email_item: Any, property_name: str, default_value: Any) -> Any:
        """
        Safely get property from email item with FORCED COM access and error handling.
        
        Args:
            email_item: The COM email object
            property_name: Name of the property to retrieve
            default_value: Default value if property is not available
            
        Returns:
            Property value or default value
        """
        try:
            # Method 1: Standard attribute access
            if hasattr(email_item, property_name):
                value = getattr(email_item, property_name)
                if value is not None:
                    logger.info(f"🔧 DEBUG: Got {property_name}={value} via standard access")
                    return value
            
            # Method 2: Try case variations
            property_variations = [
                property_name,
                property_name.lower(),
                property_name.upper(),
                property_name.capitalize()
            ]
            
            for prop_var in property_variations:
                try:
                    if hasattr(email_item, prop_var):
                        value = getattr(email_item, prop_var)
                        if value is not None:
                            logger.info(f"🔧 DEBUG: Got {property_name}={value} via case variation {prop_var}")
                            return value
                except:
                    continue
            
            # Method 3: Force COM object property access
            try:
                import pythoncom
                if hasattr(email_item, '_oleobj_'):
                    # Get the dispatch ID for the property
                    disp_id = email_item._oleobj_.GetIDsOfNames(property_name)[0]
                    # Invoke the property getter
                    value = email_item._oleobj_.Invoke(disp_id, 0, pythoncom.DISPATCH_PROPERTYGET, 1)
                    if value is not None:
                        logger.info(f"🔧 DEBUG: Got {property_name}={value} via COM invoke")
                        return value
            except Exception as com_e:
                logger.info(f"🔧 DEBUG: COM invoke failed for {property_name}: {com_e}")
            
            logger.info(f"🔧 DEBUG: All methods failed for {property_name}, using default: {default_value}")
            return default_value
            
        except Exception as e:
            logger.info(f"🔧 DEBUG: Error getting property '{property_name}': {str(e)}")
            logger.debug(f"Error getting property '{property_name}': {str(e)}")
            return default_value
    
    def _extract_recipients(self, email_item: Any) -> tuple[List[str], List[str], List[str]]:
        """
        Extract recipient information from email item.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Tuple of (recipients, cc_recipients, bcc_recipients)
        """
        recipients = []
        cc_recipients = []
        bcc_recipients = []
        
        try:
            logger.info(f"🔧 DEBUG: Starting recipient extraction...")
            if hasattr(email_item, 'Recipients') and email_item.Recipients:
                logger.info(f"🔧 DEBUG: Found {email_item.Recipients.Count} recipients")
                for recipient in email_item.Recipients:
                    try:
                        recipient_email = self._get_email_property(recipient, 'Address', '')
                        recipient_name = self._get_email_property(recipient, 'Name', '')
                        recipient_type = self._get_email_property(recipient, 'Type', 1)  # 1=To, 2=CC, 3=BCC
                        
                        logger.info(f"🔧 DEBUG: Processing recipient - Email: '{recipient_email}', Name: '{recipient_name}', Type: {recipient_type}")
                        
                        # Use email address if available, otherwise use name
                        recipient_address = recipient_email if recipient_email else recipient_name
                        
                        # Handle Exchange internal addresses and validate email format
                        if recipient_address:
                            # Convert Exchange internal addresses to a readable format
                            if recipient_address.startswith('/o=') or recipient_address.startswith('/O='):
                                logger.info(f"🔧 DEBUG: Found Exchange address: '{recipient_address}' - converting...")
                                # This is an Exchange internal address - convert to name@internal.exchange
                                if recipient_name and recipient_name != recipient_address:
                                    # Use the display name if available
                                    clean_name = ''.join(c for c in recipient_name if c.isalnum() or c in '._-')
                                    recipient_address = f"{clean_name}@internal.exchange" if clean_name else "unknown@internal.exchange"
                                    logger.info(f"🔧 DEBUG: Converted using name: '{recipient_address}'")
                                else:
                                    # Extract some identifier from the Exchange address
                                    if 'cn=' in recipient_address.lower():
                                        try:
                                            # Extract the CN (Common Name) part
                                            cn_part = recipient_address.lower().split('cn=')[-1].split('/')[0].split('-')[0]
                                            recipient_address = f"{cn_part}@internal.exchange"
                                            logger.info(f"🔧 DEBUG: Converted using CN: '{recipient_address}'")
                                        except:
                                            recipient_address = "unknown@internal.exchange"
                                            logger.info(f"🔧 DEBUG: Conversion failed, using default: '{recipient_address}'")
                                    else:
                                        recipient_address = "unknown@internal.exchange"
                                        logger.info(f"🔧 DEBUG: No CN found, using default: '{recipient_address}'")
                            
                            # Now validate the cleaned email format
                            if self._is_valid_email_format(recipient_address):
                                if recipient_type == 1:  # To
                                    recipients.append(recipient_address)
                                elif recipient_type == 2:  # CC
                                    cc_recipients.append(recipient_address)
                                elif recipient_type == 3:  # BCC
                                    bcc_recipients.append(recipient_address)
                            else:
                                # Skip invalid addresses to prevent EmailData validation errors
                                logger.debug(f"Skipping invalid recipient address: {recipient_address}")
                                continue
                                
                    except Exception as e:
                        logger.debug(f"Error processing recipient: {str(e)}")
                        continue
                        
        except Exception as e:
            logger.debug(f"Error extracting recipients: {str(e)}")
        
        return recipients, cc_recipients, bcc_recipients
    
    def _extract_email_body(self, email_item: Any) -> tuple[str, str]:
        """
        Extract email body content with enhanced formatting handling.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Tuple of (plain_text_body, html_body)
        """
        body = ""
        body_html = ""
        
        try:
            logger.info(f"🔧 DEBUG: Starting body extraction...")
            
            # Get plain text body
            body = self._get_email_property(email_item, 'Body', '')
            logger.info(f"🔧 DEBUG: Plain text body length: {len(body)}")
            
            # Get HTML body
            body_html = self._get_email_property(email_item, 'HTMLBody', '')
            logger.info(f"🔧 DEBUG: HTML body length: {len(body_html)}")
            
            # If we have HTML but no plain text, try to extract text from HTML
            if body_html and not body:
                logger.info(f"🔧 DEBUG: Extracting text from HTML...")
                body = self._extract_text_from_html(body_html)
                logger.info(f"🔧 DEBUG: Extracted text length: {len(body)}")
            
            # If we have plain text but no HTML, create basic HTML
            elif body and not body_html:
                logger.info(f"🔧 DEBUG: Creating HTML from text...")
                body_html = self._create_html_from_text(body)
            
            # Clean up the content
            body = self._clean_text_content(body)
            body_html = self._clean_html_content(body_html)
            
            logger.info(f"🔧 DEBUG: Final body lengths - Text: {len(body)}, HTML: {len(body_html)}")
            
        except Exception as e:
            logger.info(f"🔧 DEBUG: Error extracting email body: {str(e)}")
            logger.debug(f"Error extracting email body: {str(e)}")
        
        return body, body_html
    
    def _extract_timestamps(self, email_item: Any) -> tuple[Optional[datetime], Optional[datetime]]:
        """
        Extract timestamp information from email item.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Tuple of (received_time, sent_time)
        """
        received_time = None
        sent_time = None
        
        try:
            # Get received time
            if hasattr(email_item, 'ReceivedTime'):
                received_time = email_item.ReceivedTime
                
            # Get sent time
            if hasattr(email_item, 'SentOn'):
                sent_time = email_item.SentOn
            elif hasattr(email_item, 'CreationTime'):
                # Fallback to creation time if sent time not available
                sent_time = email_item.CreationTime
                
        except Exception as e:
            logger.debug(f"Error extracting timestamps: {str(e)}")
        
        return received_time, sent_time
    
    def _extract_attachment_info(self, email_item: Any) -> tuple[bool, int]:
        """
        Extract attachment information with detailed handling.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Tuple of (has_attachments, attachment_count)
        """
        has_attachments = False
        attachment_count = 0
        
        try:
            if hasattr(email_item, 'Attachments') and email_item.Attachments:
                attachment_count = email_item.Attachments.Count
                has_attachments = attachment_count > 0
                
                if has_attachments:
                    logger.debug(f"Email has {attachment_count} attachments")
                    
        except Exception as e:
            logger.debug(f"Error extracting attachment info: {str(e)}")
        
        return has_attachments, attachment_count
    
    def _extract_importance(self, email_item: Any) -> str:
        """
        Extract and map email importance level.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Importance level as string
        """
        try:
            importance_value = self._get_email_property(email_item, 'Importance', 1)
            importance_map = {0: "Low", 1: "Normal", 2: "High"}
            return importance_map.get(importance_value, "Normal")
        except Exception as e:
            logger.debug(f"Error extracting importance: {str(e)}")
            return "Normal"
    
    def _extract_folder_name(self, email_item: Any) -> str:
        """
        Extract folder name from email item's parent folder.
        
        Args:
            email_item: The COM email object
            
        Returns:
            Folder name or "Unknown" if not available
        """
        try:
            if hasattr(email_item, 'Parent') and email_item.Parent:
                parent_folder = email_item.Parent
                if hasattr(parent_folder, 'Name'):
                    return parent_folder.Name
        except Exception as e:
            logger.debug(f"Error extracting folder name: {str(e)}")
        
        return "Unknown"
    
    def send_email(self, 
                   to_recipients: List[str], 
                   subject: str, 
                   body: str, 
                   cc_recipients: List[str] = None,
                   bcc_recipients: List[str] = None,
                   body_format: str = "html",
                   importance: str = "normal",
                   attachments: List[str] = None,
                   save_to_sent_items: bool = True) -> str:
        """
        Send an email through Outlook.
        
        Args:
            to_recipients: List of recipient email addresses
            subject: Email subject
            body: Email body content
            cc_recipients: Optional list of CC recipients
            bcc_recipients: Optional list of BCC recipients
            body_format: Body format - "html", "text", or "rtf" (default: "html")
            importance: Email importance - "low", "normal", or "high" (default: "normal")
            attachments: Optional list of file paths to attach
            save_to_sent_items: Whether to save to Sent Items folder (default: True)
            
        Returns:
            str: The EntryID of the sent email
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            ValidationError: If parameters are invalid
            PermissionError: If sending is not allowed
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        # Validate required parameters
        if not to_recipients or not isinstance(to_recipients, list) or len(to_recipients) == 0:
            raise ValidationError("At least one recipient is required", "to_recipients")
        
        if not subject or not isinstance(subject, str):
            raise ValidationError("Subject is required and must be a string", "subject")
        
        if not body or not isinstance(body, str):
            raise ValidationError("Body is required and must be a string", "body")
        
        # Validate email addresses
        for email in to_recipients:
            if not self._validate_email_address(email):
                raise ValidationError(f"Invalid email address: {email}", "to_recipients")
        
        if cc_recipients:
            for email in cc_recipients:
                if not self._validate_email_address(email):
                    raise ValidationError(f"Invalid CC email address: {email}", "cc_recipients")
        
        if bcc_recipients:
            for email in bcc_recipients:
                if not self._validate_email_address(email):
                    raise ValidationError(f"Invalid BCC email address: {email}", "bcc_recipients")
        
        # Validate body format
        valid_formats = {"html": 2, "text": 1, "rtf": 3}  # Outlook constants
        if body_format.lower() not in valid_formats:
            raise ValidationError(f"Invalid body format. Must be one of: {list(valid_formats.keys())}", "body_format")
        
        # Validate importance
        importance_map = {"low": 0, "normal": 1, "high": 2}  # Outlook constants
        if importance.lower() not in importance_map:
            raise ValidationError(f"Invalid importance. Must be one of: {list(importance_map.keys())}", "importance")
        
        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            
            logger.info(f"Sending email to {len(to_recipients)} recipients: {', '.join(to_recipients)}")
            
            # Create a new Outlook application instance for this thread
            # This is necessary because COM objects can't be shared across threads
            try:
                outlook_app = win32com.client.GetActiveObject("Outlook.Application")
                logger.info("Using existing Outlook instance for sending email")
            except Exception as e:
                logger.info(f"Could not get existing Outlook instance: {e}")
                try:
                    outlook_app = win32com.client.Dispatch("Outlook.Application")
                    logger.info("Created new Outlook instance for sending email")
                except Exception as e2:
                    logger.error(f"Failed to create Outlook instance: {e2}")
                    raise OutlookConnectionError(f"Cannot create Outlook application: {e2}")
            
            # Create a new mail item
            try:
                mail_item = outlook_app.CreateItem(0)  # 0 = olMailItem
                logger.info("Successfully created mail item")
            except Exception as e:
                logger.error(f"Failed to create mail item: {e}")
                raise OutlookConnectionError(f"Cannot create mail item: {e}")
            
            # Set recipients
            for recipient in to_recipients:
                mail_item.Recipients.Add(recipient)
            
            if cc_recipients:
                for cc_recipient in cc_recipients:
                    cc_recip = mail_item.Recipients.Add(cc_recipient)
                    cc_recip.Type = 2  # olCC
            
            if bcc_recipients:
                for bcc_recipient in bcc_recipients:
                    bcc_recip = mail_item.Recipients.Add(bcc_recipient)
                    bcc_recip.Type = 3  # olBCC
            
            # Resolve all recipients
            mail_item.Recipients.ResolveAll()
            
            # Set email properties
            mail_item.Subject = subject
            mail_item.Body = body if body_format.lower() == "text" else ""
            
            if body_format.lower() == "html":
                mail_item.HTMLBody = body
            elif body_format.lower() == "rtf":
                mail_item.RTFBody = body
            
            # Set importance
            mail_item.Importance = importance_map[importance.lower()]
            
            # Add attachments if provided
            if attachments:
                for attachment_path in attachments:
                    if not isinstance(attachment_path, str):
                        logger.warning(f"Skipping invalid attachment path: {attachment_path}")
                        continue
                    
                    try:
                        import os
                        if os.path.exists(attachment_path):
                            mail_item.Attachments.Add(attachment_path)
                            logger.debug(f"Added attachment: {attachment_path}")
                        else:
                            logger.warning(f"Attachment file not found: {attachment_path}")
                    except Exception as e:
                        logger.warning(f"Failed to add attachment {attachment_path}: {str(e)}")
            
            # Get a unique identifier before sending (EntryID is not available until saved)
            # We'll use a timestamp-based ID since EntryID is not reliable for unsent items
            email_id = f"sent_{int(time.time() * 1000)}_{hash(subject + str(to_recipients))}"
            
            # Send the email - wrap in try/catch to handle Outlook COM quirks
            try:
                logger.info(f"Attempting to send email to {', '.join(to_recipients)}")
                mail_item.Send()
                logger.info(f"Email sent successfully to {', '.join(to_recipients)}")
                
                # Email was sent successfully, return immediately
                return email_id
                
            except Exception as send_error:
                error_msg = str(send_error)
                logger.warning(f"Exception during send operation: {error_msg}")
                
                # Check if this is the "item moved or deleted" error which often occurs
                # even when the email is sent successfully due to Outlook COM behavior
                if ("moved or deleted" in error_msg.lower() or 
                    "移動或刪除" in error_msg or 
                    "-2147221238" in error_msg or
                    "-2147352567" in error_msg or  # Another common Outlook COM error after successful send
                    "項目已經移動或刪除" in error_msg):
                    
                    logger.info(f"Outlook COM quirk detected - email sent successfully despite COM exception")
                    logger.info(f"This is normal behavior - Outlook moves the mail item after sending")
                    # Treat as success since this is a known Outlook COM issue
                    return email_id
                else:
                    # This is likely a real send failure
                    logger.error(f"Actual email send failure: {send_error}")
                    raise OutlookConnectionError(f"Email send failed: {send_error}")
            
        except Exception as e:
            logger.error(f"Failed to send email: {str(e)}")
            
            # Check for permission-related errors
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized', 'policy']):
                raise PermissionError("send_email", f"Permission denied to send email: {str(e)}")
            
            # Check for validation errors
            if any(keyword in str(e).lower() for keyword in ['invalid', 'malformed', 'resolve', 'recipient']):
                raise ValidationError(f"Email validation failed: {str(e)}")
            
            raise OutlookConnectionError(f"Failed to send email: {str(e)}")
        
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore cleanup errors
    
    def _validate_email_address(self, email: str) -> bool:
        """
        Validate email address format.
        
        Args:
            email: Email address to validate
            
        Returns:
            bool: True if valid, False otherwise
        """
        if not email or not isinstance(email, str):
            return False
        
        # Basic email validation regex
        import re
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email.strip()))

    def _validate_sender_info(self, sender_name: str, sender_email: str) -> tuple[str, str]:
        """
        Validate and fix sender name and email information.
        
        Args:
            sender_name: Original sender name
            sender_email: Original sender email
            
        Returns:
            Tuple of (validated_sender_name, validated_sender_email)
        """
        logger.info(f"🔧 DEBUG: Validating sender - Name: '{sender_name}', Email: '{sender_email}'")
        
        # Handle Exchange internal addresses first
        if sender_email and (sender_email.startswith('/o=') or sender_email.startswith('/O=')):
            logger.info(f"🔧 DEBUG: Found Exchange sender address: '{sender_email}' - converting...")
            # This is an Exchange internal address - convert to name@internal.exchange
            if sender_name and sender_name != sender_email:
                # Use the display name if available
                clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                sender_email = f"{clean_name}@internal.exchange" if clean_name else "unknown@internal.exchange"
                logger.info(f"🔧 DEBUG: Converted sender using name: '{sender_email}'")
            else:
                # Extract some identifier from the Exchange address
                if 'cn=' in sender_email.lower():
                    try:
                        # Extract the CN (Common Name) part
                        cn_part = sender_email.lower().split('cn=')[-1].split('/')[0].split('-')[0]
                        sender_email = f"{cn_part}@internal.exchange"
                        logger.info(f"🔧 DEBUG: Converted sender using CN: '{sender_email}'")
                    except:
                        sender_email = "unknown@internal.exchange"
                        logger.info(f"🔧 DEBUG: Sender conversion failed, using default: '{sender_email}'")
                else:
                    sender_email = "unknown@internal.exchange"
                    logger.info(f"🔧 DEBUG: No CN found in sender, using default: '{sender_email}'")
        
        # Fix sender email if invalid
        elif not sender_email or not self._is_valid_email_format(sender_email):
            logger.info(f"🔧 DEBUG: Invalid sender email format, creating from name...")
            if sender_name:
                # Create email from sender name
                clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                sender_email = f"{clean_name}@unknown.com" if clean_name else "unknown@unknown.com"
            else:
                sender_email = "unknown@unknown.com"
            logger.info(f"🔧 DEBUG: Created sender email: '{sender_email}'")
        
        # Fix sender name if missing
        if not sender_name:
            sender_name = sender_email.split('@')[0] if '@' in sender_email else "Unknown Sender"
            logger.info(f"🔧 DEBUG: Created sender name: '{sender_name}'")
        
        logger.info(f"🔧 DEBUG: Final sender - Name: '{sender_name}', Email: '{sender_email}'")
        return sender_name, sender_email
    
    def _is_valid_email_format(self, email: str) -> bool:
        """
        Check if email has valid format (basic validation).
        
        Args:
            email: Email address to validate
            
        Returns:
            True if email format is valid
        """
        if not email or not isinstance(email, str):
            return False
        
        # Must contain @ and have content before and after @
        if '@' not in email:
            return False
        
        parts = email.split('@')
        if len(parts) != 2:
            return False
        
        local_part, domain_part = parts
        
        # Local part (before @) must not be empty
        if not local_part:
            return False
        
        # Domain part must contain at least one dot and have content before and after
        if '.' not in domain_part:
            return False
        
        domain_parts = domain_part.split('.')
        # Must have at least 2 parts and none can be empty
        if len(domain_parts) < 2 or any(not part for part in domain_parts):
            return False
        
        return True
    
    def _extract_text_from_html(self, html_content: str) -> str:
        """
        Extract plain text from HTML content (basic implementation).
        
        Args:
            html_content: HTML content
            
        Returns:
            Plain text content
        """
        try:
            # Basic HTML tag removal (for simple cases)
            import re
            text = re.sub(r'<[^>]+>', '', html_content)
            text = text.replace('&nbsp;', ' ')
            text = text.replace('&lt;', '<')
            text = text.replace('&gt;', '>')
            text = text.replace('&amp;', '&')
            return text.strip()
        except Exception as e:
            logger.debug(f"Error extracting text from HTML: {str(e)}")
            return html_content
    
    def _create_html_from_text(self, text_content: str) -> str:
        """
        Create basic HTML from plain text content.
        
        Args:
            text_content: Plain text content
            
        Returns:
            Basic HTML content
        """
        try:
            # Convert line breaks to HTML
            html_content = text_content.replace('\n', '<br>\n')
            return f"<html><body>{html_content}</body></html>"
        except Exception as e:
            logger.debug(f"Error creating HTML from text: {str(e)}")
            return text_content
    
    def _clean_text_content(self, content: str) -> str:
        """
        Clean and normalize text content.
        
        Args:
            content: Text content to clean
            
        Returns:
            Cleaned text content
        """
        if not content:
            return ""
        
        try:
            # Remove excessive whitespace
            content = re.sub(r'\s+', ' ', content)
            content = content.strip()
            return content
        except Exception as e:
            logger.debug(f"Error cleaning text content: {str(e)}")
            return content
    
    def _clean_html_content(self, content: str) -> str:
        """
        Clean and normalize HTML content.
        
        Args:
            content: HTML content to clean
            
        Returns:
            Cleaned HTML content
        """
        if not content:
            return ""
        
        try:
            # Basic HTML cleanup - remove excessive whitespace between tags
            content = re.sub(r'>\s+<', '><', content)
            content = content.strip()
            return content
        except Exception as e:
            logger.debug(f"Error cleaning HTML content: {str(e)}")
            return content

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.disconnect()