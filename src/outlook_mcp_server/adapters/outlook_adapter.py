"""Outlook COM adapter for interfacing with Microsoft Outlook."""

import logging
import re
from datetime import datetime
from typing import Optional, List, Any, Tuple
from unittest.mock import Mock
import win32com.client
import pythoncom
from ..models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    EmailNotFoundError,
    PermissionError
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
            # First check default folders
            default_folders = {
                "Inbox": 6,      # olFolderInbox
                "Outbox": 4,     # olFolderOutbox
                "Sent Items": 5, # olFolderSentMail
                "Deleted Items": 3, # olFolderDeletedItems
                "Drafts": 16,    # olFolderDrafts
                "Junk Email": 23 # olFolderJunk
            }
            
            if name in default_folders:
                folder = self._namespace.GetDefaultFolder(default_folders[name])
                if folder:
                    logger.debug(f"Found default folder: {name}")
                    return folder
            
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
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if not email_id or not isinstance(email_id, str):
            raise EmailNotFoundError(email_id or "")
        
        # Validate email ID format
        if not EmailData.validate_email_id(email_id):
            raise EmailNotFoundError(email_id)
        
        try:
            logger.debug(f"Retrieving detailed email with ID: {email_id}")
            
            # Try to get the email item by EntryID
            email_item = self._namespace.GetItemFromID(email_id)
            
            if not email_item:
                logger.warning(f"Email not found: {email_id}")
                raise EmailNotFoundError(email_id)
            
            # Verify it's a mail item (type 43 = olMail)
            if not hasattr(email_item, 'Class') or email_item.Class != 43:
                logger.warning(f"Item is not a mail item: {email_id}")
                raise EmailNotFoundError(email_id)
            
            # Transform to detailed EmailData with enhanced content extraction
            email_data = self._transform_email_to_detailed_data(email_item)
            
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
    
    def search_emails(self, query: str, folder_name: str = None, limit: int = 50) -> List[EmailData]:
        """
        Search emails based on query with enhanced functionality.
        
        Args:
            query: Search query string
            folder_name: Optional folder name to search in (searches all accessible folders if None)
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
            logger.debug(f"Searching emails with query: '{query}', folder: {folder_name or 'all folders'}, limit: {limit}")
            
            # Process search query to ensure Outlook compatibility
            processed_query = self._process_search_query(query)
            
            if folder_name:
                # Search in specific folder
                return self._search_in_folder(processed_query, folder_name, limit)
            else:
                # Search across all accessible folders
                return self._search_all_folders(processed_query, limit)
            
        except (FolderNotFoundError, PermissionError):
            raise
        except Exception as e:
            logger.error(f"Error searching emails: {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_name or "folders", f"Access denied during search: {str(e)}")
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
    
    def _search_in_folder(self, query: str, folder_name: str, limit: int) -> List[EmailData]:
        """
        Search emails in a specific folder.
        
        Args:
            query: Processed search query
            folder_name: Name of the folder to search in
            limit: Maximum number of results to return
            
        Returns:
            List[EmailData]: List of matching email data objects
        """
        try:
            # Get the target folder
            folder = self.get_folder_by_name(folder_name)
            
            if not folder:
                raise FolderNotFoundError(folder_name)
            
            logger.debug(f"Searching in folder: {folder_name}")
            
            # Perform search in the folder
            results = self._perform_folder_search(folder, query, limit, folder_name)
            
            logger.debug(f"Found {len(results)} emails in folder '{folder_name}'")
            return results
            
        except (FolderNotFoundError, PermissionError):
            raise
        except Exception as e:
            logger.error(f"Error searching in folder '{folder_name}': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_name, f"Access denied to folder '{folder_name}': {str(e)}")
            return []
    
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
            
            # Use Outlook's Find method to search
            found_item = items.Find(query)
            
            results = []
            count = 0
            
            # Collect search results with safety counter to prevent infinite loops
            max_iterations = limit * 10  # Safety limit to prevent infinite loops
            iteration_count = 0
            
            while found_item and count < limit and iteration_count < max_iterations:
                iteration_count += 1
                
                try:
                    # Check if it's a mail item (type 43 = olMail)
                    if hasattr(found_item, 'Class') and found_item.Class == 43:
                        # Transform to EmailData
                        email_data = self._transform_email_to_data(found_item, folder_name)
                        results.append(email_data)
                        count += 1
                    
                    # Get next result
                    found_item = items.FindNext()
                    
                except Exception as e:
                    logger.debug(f"Error processing search result: {str(e)}")
                    # Try to get next result
                    try:
                        found_item = items.FindNext()
                    except:
                        break
                    continue
            
            if iteration_count >= max_iterations:
                logger.warning(f"Search in folder '{folder_name}' hit iteration limit, may have more results")
            
            logger.debug(f"Search in folder '{folder_name}' returned {len(results)} results")
            return results
            
        except Exception as e:
            logger.error(f"Error performing search in folder '{folder_name}': {str(e)}")
            return []
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def list_emails(self, folder_name: str = None, unread_only: bool = False, limit: int = 50) -> List[EmailData]:
        """
        List emails from specified folder with filtering options.
        
        Args:
            folder_name: Name of the folder to list emails from (None for Inbox)
            unread_only: If True, only return unread emails
            limit: Maximum number of emails to return
            
        Returns:
            List[EmailData]: List of email data objects
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
        """
        if not self.is_connected():
            raise OutlookConnectionError("Not connected to Outlook")
        
        if limit <= 0:
            limit = 50  # Default limit
        
        try:
            logger.debug(f"Listing emails from folder: {folder_name or 'Inbox'}, unread_only: {unread_only}, limit: {limit}")
            
            # Get the target folder
            if folder_name:
                folder = self.get_folder_by_name(folder_name)
            else:
                # Default to Inbox
                folder = self._namespace.GetDefaultFolder(6)  # olFolderInbox
            
            if not folder:
                raise FolderNotFoundError(folder_name or "Inbox")
            
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
                    email_data = self._transform_email_to_data(item, folder_name or "Inbox")
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
            logger.error(f"Error listing emails from folder '{folder_name}': {str(e)}")
            if "access" in str(e).lower() or "permission" in str(e).lower():
                raise PermissionError(folder_name or "Inbox", f"Access denied to folder: {str(e)}")
            raise OutlookConnectionError(f"Failed to list emails: {str(e)}")
    
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
            # Get basic email properties
            email_id = getattr(email_item, 'EntryID', '')
            subject = getattr(email_item, 'Subject', '')
            sender_name = getattr(email_item, 'SenderName', '')
            sender_email = getattr(email_item, 'SenderEmailAddress', '')
            
            # Ensure we have a valid ID
            if not email_id:
                email_id = f"unknown_{id(email_item)}"
            
            # Get recipient information
            recipients = []
            cc_recipients = []
            bcc_recipients = []
            
            try:
                if hasattr(email_item, 'Recipients'):
                    for recipient in email_item.Recipients:
                        recipient_email = getattr(recipient, 'Address', '')
                        recipient_type = getattr(recipient, 'Type', 1)  # 1=To, 2=CC, 3=BCC
                        
                        if recipient_type == 1:  # To
                            recipients.append(recipient_email)
                        elif recipient_type == 2:  # CC
                            cc_recipients.append(recipient_email)
                        elif recipient_type == 3:  # BCC
                            bcc_recipients.append(recipient_email)
            except Exception as e:
                logger.debug(f"Error processing recipients: {str(e)}")
            
            # Get email body
            body = getattr(email_item, 'Body', '')
            body_html = getattr(email_item, 'HTMLBody', '')
            
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
            
            # Validate and fix sender email - use sender name if email is invalid
            if not sender_email or '@' not in sender_email:
                if sender_name:
                    # Remove any invalid characters from sender name for email
                    clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                    sender_email = f"{clean_name}@unknown.com" if clean_name else "unknown@unknown.com"
                else:
                    sender_email = "unknown@unknown.com"
            
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
            # Return minimal email data if transformation fails
            email_id = getattr(email_item, 'EntryID', f"unknown_{id(email_item)}")
            subject = getattr(email_item, 'Subject', '(No Subject)')
            sender_name = getattr(email_item, 'SenderName', 'Unknown Sender')
            sender_email = getattr(email_item, 'SenderEmailAddress', 'unknown@unknown.com')
            
            # Ensure valid email format
            if not sender_email or '@' not in sender_email:
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
            recipients, cc_recipients, bcc_recipients = self._extract_recipients(email_item)
            
            # Get email body with enhanced content extraction
            body, body_html = self._extract_email_body(email_item)
            
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
            sender_name, sender_email = self._validate_sender_info(sender_name, sender_email)
            
            # Create EmailData with all extracted information
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
        Safely get property from email item with error handling.
        
        Args:
            email_item: The COM email object
            property_name: Name of the property to retrieve
            default_value: Default value if property is not available
            
        Returns:
            Property value or default value
        """
        try:
            if hasattr(email_item, property_name):
                value = getattr(email_item, property_name)
                return value if value is not None else default_value
            return default_value
        except Exception as e:
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
            if hasattr(email_item, 'Recipients') and email_item.Recipients:
                for recipient in email_item.Recipients:
                    try:
                        recipient_email = self._get_email_property(recipient, 'Address', '')
                        recipient_name = self._get_email_property(recipient, 'Name', '')
                        recipient_type = self._get_email_property(recipient, 'Type', 1)  # 1=To, 2=CC, 3=BCC
                        
                        # Use email address if available, otherwise use name
                        recipient_address = recipient_email if recipient_email else recipient_name
                        
                        # Validate email format
                        if recipient_address and self._is_valid_email_format(recipient_address):
                            if recipient_type == 1:  # To
                                recipients.append(recipient_address)
                            elif recipient_type == 2:  # CC
                                cc_recipients.append(recipient_address)
                            elif recipient_type == 3:  # BCC
                                bcc_recipients.append(recipient_address)
                        elif recipient_address:  # Add even if not valid email format
                            if recipient_type == 1:
                                recipients.append(recipient_address)
                            elif recipient_type == 2:
                                cc_recipients.append(recipient_address)
                            elif recipient_type == 3:
                                bcc_recipients.append(recipient_address)
                                
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
            # Get plain text body
            body = self._get_email_property(email_item, 'Body', '')
            
            # Get HTML body
            body_html = self._get_email_property(email_item, 'HTMLBody', '')
            
            # If we have HTML but no plain text, try to extract text from HTML
            if body_html and not body:
                body = self._extract_text_from_html(body_html)
            
            # If we have plain text but no HTML, create basic HTML
            elif body and not body_html:
                body_html = self._create_html_from_text(body)
            
            # Clean up the content
            body = self._clean_text_content(body)
            body_html = self._clean_html_content(body_html)
            
        except Exception as e:
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
    
    def _validate_sender_info(self, sender_name: str, sender_email: str) -> tuple[str, str]:
        """
        Validate and fix sender name and email information.
        
        Args:
            sender_name: Original sender name
            sender_email: Original sender email
            
        Returns:
            Tuple of (validated_sender_name, validated_sender_email)
        """
        # Fix sender email if invalid
        if not sender_email or not self._is_valid_email_format(sender_email):
            if sender_name:
                # Create email from sender name
                clean_name = ''.join(c for c in sender_name if c.isalnum() or c in '._-')
                sender_email = f"{clean_name}@unknown.com" if clean_name else "unknown@unknown.com"
            else:
                sender_email = "unknown@unknown.com"
        
        # Fix sender name if missing
        if not sender_name:
            sender_name = sender_email.split('@')[0] if '@' in sender_email else "Unknown Sender"
        
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