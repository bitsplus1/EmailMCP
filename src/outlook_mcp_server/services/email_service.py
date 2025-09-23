"""Email service layer for handling email operations."""

import logging
from typing import List, Dict, Any, Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from ..performance.memory_manager import MemoryManager
    from ..performance.lazy_loader import LazyEmailLoader
    from ..performance.rate_limiter import RateLimiter
from ..adapters.outlook_adapter import OutlookAdapter
from ..adapters.connection_pool import OutlookConnectionPool
from ..models.email_data import EmailData
from ..models.exceptions import (
    OutlookConnectionError,
    EmailNotFoundError,
    FolderNotFoundError,
    PermissionError,
    ValidationError,
    SearchError
)
# Performance imports - temporarily commented to avoid circular imports
# from ..performance.memory_manager import MemoryManager, MemoryConfig
# from ..performance.lazy_loader import LazyEmailLoader, LazyLoadConfig
# from ..performance.rate_limiter import RateLimiter, RateLimitConfig


logger = logging.getLogger(__name__)


class EmailService:
    """Service layer for email management operations."""
    
    def __init__(self, 
                 outlook_adapter: OutlookAdapter,
                 connection_pool: Optional[OutlookConnectionPool] = None,
                 memory_manager: Optional['MemoryManager'] = None,
                 lazy_loader: Optional['LazyEmailLoader'] = None,
                 rate_limiter: Optional['RateLimiter'] = None):
        """
        Initialize the email service.
        
        Args:
            outlook_adapter: The Outlook adapter instance for COM operations
            connection_pool: Optional connection pool for performance
            memory_manager: Optional memory manager for caching
            lazy_loader: Optional lazy loader for content
            rate_limiter: Optional rate limiter for request throttling
        """
        self.outlook_adapter = outlook_adapter
        self.connection_pool = connection_pool
        self.memory_manager = memory_manager
        self.lazy_loader = lazy_loader
        self.rate_limiter = rate_limiter
        
        # Performance statistics
        self._stats = {
            "requests_processed": 0,
            "cache_hits": 0,
            "cache_misses": 0,
            "rate_limited_requests": 0
        }
        
    async def list_inbox_emails(self, unread_only: bool = False, limit: int = 50, client_id: str = "default") -> List[Dict[str, Any]]:
        """
        List emails from the default inbox folder.
        
        Args:
            unread_only: Whether to show only unread emails
            limit: Maximum number of emails to return
            
        Returns:
            List[Dict[str, Any]]: List of email data in JSON format
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            PermissionError: If access to inbox is denied
            ValidationError: If parameters are invalid
        """
        try:
            # Apply rate limiting
            if self.rate_limiter:
                await self.rate_limiter.acquire(client_id, "list_emails")
            
            logger.info(f"Listing emails from inbox - unread_only: {unread_only}, limit: {limit}")
            
            # Validate parameters
            if limit <= 0:
                limit = 50
            elif limit > 1000:
                limit = 1000
            
            # Check cache first if memory manager is available
            cache_key = f"list_inbox_emails:{unread_only}:{limit}"
            if self.memory_manager:
                cached_result = self.memory_manager.folder_cache.get(cache_key)
                if cached_result:
                    self._stats["cache_hits"] += 1
                    logger.debug(f"Cache hit for inbox email list: {cache_key}")
                    return cached_result
                else:
                    self._stats["cache_misses"] += 1
            
            # Use connection pool if available
            if self.connection_pool:
                with self.connection_pool.get_connection() as connection:
                    # Create temporary adapter with pooled connection
                    temp_adapter = OutlookAdapter()
                    temp_adapter._outlook_app = connection.outlook_app
                    temp_adapter._namespace = connection.namespace
                    temp_adapter._connected = True
                    
                    email_data_list = temp_adapter.list_inbox_emails(unread_only, limit)
            else:
                # Ensure we're connected to Outlook
                if not self.outlook_adapter.is_connected():
                    logger.error("Outlook adapter is not connected")
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Get emails from the adapter
                email_data_list = self.outlook_adapter.list_inbox_emails(unread_only, limit)
            
            # Transform to JSON format
            json_emails = []
            for email_data in email_data_list:
                try:
                    json_email = self._transform_email_to_json(email_data)
                    json_emails.append(json_email)
                    
                    # Cache individual emails if memory manager is available
                    if self.memory_manager:
                        self.memory_manager.cache_email(email_data.id, email_data)
                        
                except Exception as e:
                    logger.warning(f"Error transforming email '{email_data.id}': {str(e)}")
                    continue
            
            # Cache the result if memory manager is available
            if self.memory_manager:
                self.memory_manager.folder_cache.put(cache_key, json_emails, len(str(json_emails)))
            
            # Preload email content if lazy loader is available
            if self.lazy_loader and len(json_emails) > 0:
                email_ids = [email["id"] for email in json_emails[:10]]  # Preload first 10
                self.lazy_loader.preload_emails(email_ids)
            
            self._stats["requests_processed"] += 1
            logger.info(f"Successfully listed {len(json_emails)} emails")
            return json_emails
            
        except (OutlookConnectionError, FolderNotFoundError, PermissionError, ValidationError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error listing emails: {str(e)}")
            # Check if it's a permission-related error
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError("inbox", f"Access denied: {str(e)}")
            # Otherwise, treat as connection error
            raise OutlookConnectionError(f"Failed to list emails: {str(e)}")
    
    async def list_emails(self, folder_id: str, unread_only: bool = False, limit: int = 50, client_id: str = "default") -> List[Dict[str, Any]]:
        """
        List emails from a specific folder by folder ID.
        
        Args:
            folder_id: ID of the folder to list emails from
            unread_only: Whether to show only unread emails
            limit: Maximum number of emails to return
            
        Returns:
            List[Dict[str, Any]]: List of email data in JSON format
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
            ValidationError: If parameters are invalid
        """
        try:
            # Apply rate limiting
            if self.rate_limiter:
                await self.rate_limiter.acquire(client_id, "list_emails")
            
            logger.info(f"Listing emails from folder ID: {folder_id[:20] if len(folder_id) > 20 else folder_id}..., unread_only: {unread_only}, limit: {limit}")
            
            # Validate parameters
            if not folder_id or not isinstance(folder_id, str):
                raise ValidationError("folder_id must be a non-empty string", "folder_id")
            
            if limit <= 0:
                limit = 50
            elif limit > 1000:
                limit = 1000
            
            # Check cache first if memory manager is available
            cache_key = f"list_emails:{folder_id}:{unread_only}:{limit}"
            if self.memory_manager:
                cached_result = self.memory_manager.folder_cache.get(cache_key)
                if cached_result:
                    self._stats["cache_hits"] += 1
                    logger.debug(f"Cache hit for folder email list: {cache_key}")
                    return cached_result
                else:
                    self._stats["cache_misses"] += 1
            
            # Use connection pool if available
            if self.connection_pool:
                with self.connection_pool.get_connection() as connection:
                    # Create temporary adapter with pooled connection
                    temp_adapter = OutlookAdapter()
                    temp_adapter._outlook_app = connection.outlook_app
                    temp_adapter._namespace = connection.namespace
                    temp_adapter._connected = True
                    
                    email_data_list = temp_adapter.list_emails(folder_id, unread_only, limit)
            else:
                # Ensure we're connected to Outlook
                if not self.outlook_adapter.is_connected():
                    logger.error("Outlook adapter is not connected")
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Get emails from the adapter
                email_data_list = self.outlook_adapter.list_emails(folder_id, unread_only, limit)
            
            # Transform to JSON format
            json_emails = []
            for email_data in email_data_list:
                try:
                    json_email = self._transform_email_to_json(email_data)
                    json_emails.append(json_email)
                    
                    # Cache individual emails if memory manager is available
                    if self.memory_manager:
                        self.memory_manager.cache_email(email_data.id, email_data)
                        
                except Exception as e:
                    logger.warning(f"Error transforming email '{email_data.id}': {str(e)}")
                    continue
            
            # Cache results if memory manager is available
            if self.memory_manager and len(json_emails) > 0:
                self.memory_manager.folder_cache.put(cache_key, json_emails, len(str(json_emails)))
            
            # Preload email content if lazy loader is available
            if self.lazy_loader and len(json_emails) > 0:
                email_ids = [email["id"] for email in json_emails[:10]]  # Preload first 10
                self.lazy_loader.preload_emails(email_ids)
            
            self._stats["requests_processed"] += 1
            logger.info(f"Listed {len(json_emails)} emails from folder ID: {folder_id[:20]}...")
            return json_emails
            
        except (OutlookConnectionError, FolderNotFoundError, PermissionError, ValidationError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error listing emails from folder ID: {str(e)}")
            # Check if it's a permission-related error
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError(folder_id, f"Access denied: {str(e)}")
            # Otherwise, treat as connection error
            raise OutlookConnectionError(f"Failed to list emails: {str(e)}")
    
    async def get_email(self, email_id: str, client_id: str = "default") -> Dict[str, Any]:
        """
        Get detailed information for a specific email by ID.
        
        Args:
            email_id: Unique identifier of the email to retrieve
            
        Returns:
            Dict[str, Any]: Complete email data in JSON format
            
        Raises:
            ValidationError: If email ID is invalid
            EmailNotFoundError: If email is not found
            PermissionError: If access to email is denied
            OutlookConnectionError: If not connected to Outlook
        """
        try:
            # Apply rate limiting
            if self.rate_limiter:
                await self.rate_limiter.acquire(client_id, "get_email")
            
            logger.debug(f"Retrieving email with ID: {email_id}")
            
            # Validate input
            if not email_id or not isinstance(email_id, str):
                raise ValidationError("Email ID must be a non-empty string", "email_id")
            
            # Validate email ID format
            if not EmailData.validate_email_id(email_id):
                raise ValidationError(f"Invalid email ID format: {email_id}", "email_id")
            
            # Check cache first if memory manager is available
            if self.memory_manager:
                cached_email = self.memory_manager.get_cached_email(email_id)
                if cached_email:
                    self._stats["cache_hits"] += 1
                    logger.debug(f"Cache hit for email: {email_id}")
                    return self._transform_email_to_json(cached_email)
                else:
                    self._stats["cache_misses"] += 1
            
            # Use lazy loader if available
            if self.lazy_loader:
                lazy_content = self.lazy_loader.get_lazy_email(email_id)
                email_data = lazy_content.get_content()
            else:
                # Use connection pool if available
                if self.connection_pool:
                    with self.connection_pool.get_connection() as connection:
                        # Create temporary adapter with pooled connection
                        temp_adapter = OutlookAdapter()
                        temp_adapter._outlook_app = connection.outlook_app
                        temp_adapter._namespace = connection.namespace
                        temp_adapter._connected = True
                        
                        email_data = temp_adapter.get_email_by_id(email_id)
                else:
                    # Ensure we're connected
                    if not self.outlook_adapter.is_connected():
                        raise OutlookConnectionError("Not connected to Outlook")
                    
                    # Get the email from adapter
                    email_data = self.outlook_adapter.get_email_by_id(email_id)
            
            # Cache the email if memory manager is available
            if self.memory_manager:
                self.memory_manager.cache_email(email_id, email_data)
            
            # Transform to JSON format
            json_email = self._transform_email_to_json(email_data)
            
            self._stats["requests_processed"] += 1
            logger.debug(f"Successfully retrieved email: {email_id}")
            return json_email
            
        except (ValidationError, EmailNotFoundError, PermissionError, OutlookConnectionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error retrieving email '{email_id}': {str(e)}")
            # Check error type and convert to appropriate exception
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError(email_id, f"Access denied to email '{email_id}': {str(e)}")
            elif any(keyword in str(e).lower() for keyword in ['not found', 'invalid', 'missing']):
                raise EmailNotFoundError(email_id)
            else:
                raise OutlookConnectionError(f"Failed to retrieve email '{email_id}': {str(e)}")
    
    async def search_emails(self, query: str, folder_id: str = None, limit: int = 50, client_id: str = "default") -> List[Dict[str, Any]]:
        """
        Search emails based on user-defined queries.
        
        Args:
            query: Search query string
            folder_id: Optional folder ID to search in
            limit: Maximum number of results to return
            
        Returns:
            List[Dict[str, Any]]: List of matching email data in JSON format
            
        Raises:
            ValidationError: If query is invalid
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
            OutlookConnectionError: If not connected to Outlook
            SearchError: If search operation fails
        """
        try:
            # Apply rate limiting
            if self.rate_limiter:
                await self.rate_limiter.acquire(client_id, "search_emails")
            
            logger.info(f"Searching emails - query: '{query}', folder_id: {folder_id}, limit: {limit}")
            
            # Validate input
            if not query or not isinstance(query, str):
                raise ValidationError("Search query must be a non-empty string", "query")
            
            query = query.strip()
            if not query:
                raise ValidationError("Search query cannot be empty", "query")
            
            if len(query) > 1000:
                raise ValidationError("Search query is too long (max 1000 characters)", "query")
            
            # Validate limit
            if limit <= 0:
                limit = 50
            elif limit > 1000:
                limit = 1000
            
            # Check cache for search results if memory manager is available
            cache_key = f"search:{hash(query)}:{folder_id}:{limit}"
            if self.memory_manager:
                cached_result = self.memory_manager.folder_cache.get(cache_key)
                if cached_result:
                    self._stats["cache_hits"] += 1
                    logger.debug(f"Cache hit for search: {query}")
                    return cached_result
                else:
                    self._stats["cache_misses"] += 1
            
            # Use connection pool if available
            if self.connection_pool:
                with self.connection_pool.get_connection() as connection:
                    # Create temporary adapter with pooled connection
                    temp_adapter = OutlookAdapter()
                    temp_adapter._outlook_app = connection.outlook_app
                    temp_adapter._namespace = connection.namespace
                    temp_adapter._connected = True
                    
                    email_data_list = temp_adapter.search_emails(query, folder_identifier=folder_id, limit=limit)
            else:
                # Ensure we're connected
                if not self.outlook_adapter.is_connected():
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Perform search using adapter
                email_data_list = self.outlook_adapter.search_emails(query, folder_identifier=folder_id, limit=limit)
            
            # Transform to JSON format
            json_emails = []
            for email_data in email_data_list:
                try:
                    json_email = self._transform_email_to_json(email_data)
                    json_emails.append(json_email)
                    
                    # Cache individual emails if memory manager is available
                    if self.memory_manager:
                        self.memory_manager.cache_email(email_data.id, email_data)
                        
                except Exception as e:
                    logger.warning(f"Error transforming search result email '{getattr(email_data, 'id', 'unknown')}': {str(e)}")
                    continue
            
            # Sort results by received time (newest first)
            json_emails.sort(key=lambda x: x.get('received_time', ''), reverse=True)
            
            # Cache search results if memory manager is available
            if self.memory_manager and len(json_emails) > 0:
                # Cache for shorter time since search results can change
                self.memory_manager.folder_cache.put(cache_key, json_emails, len(str(json_emails)))
            
            # Preload email content if lazy loader is available
            if self.lazy_loader and len(json_emails) > 0:
                email_ids = [email["id"] for email in json_emails[:5]]  # Preload first 5 search results
                self.lazy_loader.preload_emails(email_ids)
            
            self._stats["requests_processed"] += 1
            logger.info(f"Search completed: {len(json_emails)} results found")
            return json_emails
            
        except (ValidationError, FolderNotFoundError, PermissionError, OutlookConnectionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error searching emails: {str(e)}")
            # Check error type and convert to appropriate exception
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized']):
                raise PermissionError(folder_id or "emails", f"Access denied during search: {str(e)}")
            elif any(keyword in str(e).lower() for keyword in ['not found', 'invalid', 'missing']) and folder_id:
                raise FolderNotFoundError(folder_id)
            else:
                raise SearchError(query, f"Search operation failed: {str(e)}")
    
    def _transform_email_to_json(self, email_data: EmailData) -> Dict[str, Any]:
        """
        Transform EmailData object to JSON-serializable dictionary.
        
        Args:
            email_data: The EmailData object to transform
            
        Returns:
            Dict[str, Any]: JSON-serializable email data
            
        Raises:
            ValidationError: If email data is invalid
        """
        try:
            # Validate the email data
            email_data.validate()
            
            # Use the built-in to_dict method
            json_data = email_data.to_dict()
            
            # Add additional metadata for API response
            json_data.update({
                "accessible": True,  # If we got this far, email is accessible
                "has_body": bool(email_data.body or email_data.body_html),
                "body_preview": self._get_body_preview(email_data),
                "attachment_count": len(email_data.attachments) if hasattr(email_data, 'attachments') else 0
            })
            
            return json_data
            
        except ValidationError:
            # Re-raise validation errors
            raise
        except Exception as e:
            logger.error(f"Error transforming email data to JSON: {str(e)}")
            raise ValidationError(f"Failed to transform email data: {str(e)}")
    
    def _get_body_preview(self, email_data: EmailData, max_length: int = 200) -> str:
        """
        Get a preview of the email body for display purposes.
        
        Args:
            email_data: The email data
            max_length: Maximum length of the preview
            
        Returns:
            str: Body preview text
        """
        try:
            # Prefer plain text body over HTML
            body_text = email_data.body or email_data.body_html or ""
            
            # Remove HTML tags if present
            if email_data.body_html and not email_data.body:
                import re
                body_text = re.sub(r'<[^>]+>', '', body_text)
            
            # Clean up whitespace
            body_text = ' '.join(body_text.split())
            
            # Truncate if too long
            if len(body_text) > max_length:
                body_text = body_text[:max_length] + "..."
            
            return body_text
            
        except Exception as e:
            logger.debug(f"Error generating body preview: {str(e)}")
            return ""
    
    def get_email_statistics(self, folder: str = None) -> Dict[str, Any]:
        """
        Get statistics about emails in a folder or all folders.
        
        Args:
            folder: Optional folder name to get statistics for
            
        Returns:
            Dict[str, Any]: Email statistics
            
        Raises:
            OutlookConnectionError: If not connected to Outlook
            FolderNotFoundError: If specified folder is not found
            PermissionError: If access to folder is denied
        """
        try:
            logger.debug(f"Generating email statistics for folder: {folder or 'all folders'}")
            
            # Ensure we're connected
            if not self.outlook_adapter.is_connected():
                raise OutlookConnectionError("Not connected to Outlook")
            
            # Get a sample of emails to analyze
            sample_emails = self.list_emails(folder, unread_only=False, limit=100)
            
            # Calculate statistics
            stats = {
                "total_sampled": len(sample_emails),
                "unread_count": sum(1 for email in sample_emails if not email.get("is_read", True)),
                "has_attachments_count": sum(1 for email in sample_emails if email.get("has_attachments", False)),
                "senders": {},
                "recent_activity": {}
            }
            
            # Analyze senders
            for email in sample_emails:
                sender = email.get("sender", "Unknown")
                stats["senders"][sender] = stats["senders"].get(sender, 0) + 1
            
            # Get top senders
            top_senders = sorted(stats["senders"].items(), key=lambda x: x[1], reverse=True)[:10]
            stats["top_senders"] = [{"sender": sender, "count": count} for sender, count in top_senders]
            
            # Analyze recent activity (last 7 days)
            from datetime import datetime, timedelta
            week_ago = datetime.now() - timedelta(days=7)
            
            recent_emails = [
                email for email in sample_emails 
                if email.get("received_time") and 
                datetime.fromisoformat(email["received_time"].replace('Z', '+00:00')) > week_ago
            ]
            
            stats["recent_activity"] = {
                "last_7_days": len(recent_emails),
                "daily_average": len(recent_emails) / 7 if recent_emails else 0
            }
            
            logger.debug(f"Generated statistics: {stats['total_sampled']} emails analyzed")
            return stats
            
        except (OutlookConnectionError, FolderNotFoundError, PermissionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Error generating email statistics: {str(e)}")
            raise OutlookConnectionError(f"Failed to generate email statistics: {str(e)}")
    
    def get_performance_stats(self) -> Dict[str, Any]:
        """Get performance statistics for the email service."""
        stats = self._stats.copy()
        
        # Add component stats if available
        if self.memory_manager:
            stats["memory"] = self.memory_manager.get_stats()
        
        if self.lazy_loader:
            stats["lazy_loader"] = self.lazy_loader.get_stats()
        
        if self.rate_limiter:
            stats["rate_limiter"] = self.rate_limiter.get_stats()
        
        if self.connection_pool:
            stats["connection_pool"] = self.connection_pool.get_stats()
        
        return stats
    
    async def send_email(self, 
                        to_recipients: List[str], 
                        subject: str, 
                        body: str, 
                        cc_recipients: List[str] = None,
                        bcc_recipients: List[str] = None,
                        body_format: str = "html",
                        importance: str = "normal",
                        attachments: List[str] = None,
                        save_to_sent_items: bool = True,
                        client_id: str = "default") -> Dict[str, Any]:
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
            client_id: Client identifier for rate limiting
            
        Returns:
            Dict[str, Any]: Send result with email ID and status
            
        Raises:
            ValidationError: If parameters are invalid
            PermissionError: If sending is not allowed
            OutlookConnectionError: If not connected to Outlook
        """
        try:
            # Apply rate limiting
            if self.rate_limiter:
                await self.rate_limiter.acquire(client_id, "send_email")
            
            logger.info(f"Sending email to {len(to_recipients)} recipients: {', '.join(to_recipients[:3])}{'...' if len(to_recipients) > 3 else ''}")
            
            # Validate parameters
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
            
            # Use connection pool if available
            if self.connection_pool:
                with self.connection_pool.get_connection() as connection:
                    # Create temporary adapter with pooled connection
                    temp_adapter = OutlookAdapter()
                    temp_adapter._outlook_app = connection.outlook_app
                    temp_adapter._namespace = connection.namespace
                    temp_adapter._connected = True
                    
                    email_id = temp_adapter.send_email(
                        to_recipients=to_recipients,
                        subject=subject,
                        body=body,
                        cc_recipients=cc_recipients,
                        bcc_recipients=bcc_recipients,
                        body_format=body_format,
                        importance=importance,
                        attachments=attachments,
                        save_to_sent_items=save_to_sent_items
                    )
            else:
                # Ensure we're connected to Outlook
                if not self.outlook_adapter.is_connected():
                    logger.error("Outlook adapter is not connected")
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Send email using adapter
                email_id = self.outlook_adapter.send_email(
                    to_recipients=to_recipients,
                    subject=subject,
                    body=body,
                    cc_recipients=cc_recipients,
                    bcc_recipients=bcc_recipients,
                    body_format=body_format,
                    importance=importance,
                    attachments=attachments,
                    save_to_sent_items=save_to_sent_items
                )
            
            # Verify email was actually sent by searching for it
            verification_result = await self._verify_email_sent(subject, to_recipients[0], email_id)
            
            # Prepare response
            result = {
                "email_id": email_id,
                "status": "sent",
                "recipients": {
                    "to": to_recipients,
                    "cc": cc_recipients or [],
                    "bcc": bcc_recipients or []
                },
                "subject": subject,
                "body_format": body_format,
                "importance": importance,
                "attachments_count": len(attachments) if attachments else 0,
                "sent_time": self._get_current_timestamp(),
                "verification": verification_result
            }
            
            self._stats["requests_processed"] += 1
            logger.info(f"Email sent successfully with ID: {email_id}, verification: {verification_result['status']}")
            return result
            
        except (ValidationError, PermissionError, OutlookConnectionError):
            # Re-raise known exceptions
            raise
        except Exception as e:
            logger.error(f"Unexpected error sending email: {str(e)}")
            # Check if it's a permission-related error
            if any(keyword in str(e).lower() for keyword in ['access', 'permission', 'denied', 'unauthorized', 'policy']):
                raise PermissionError("send_email", f"Permission denied to send email: {str(e)}")
            # Otherwise, treat as connection error
            raise OutlookConnectionError(f"Failed to send email: {str(e)}")
    
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
    
    def _get_current_timestamp(self) -> str:
        """Get current timestamp in ISO format."""
        from datetime import datetime
        return datetime.utcnow().isoformat() + "Z"
    
    async def _verify_email_sent(self, subject: str, recipient: str, email_id: str) -> Dict[str, Any]:
        """
        Verify that an email was actually sent by searching for it in Sent Items.
        
        Args:
            subject: Email subject to search for
            recipient: First recipient to search for
            email_id: Email ID returned from send operation
            
        Returns:
            Dict[str, Any]: Verification result with status and details
        """
        try:
            import asyncio
            import time
            
            # Wait a moment for the email to appear in Sent Items
            await asyncio.sleep(2)
            
            logger.debug(f"Verifying email sent: subject='{subject}', recipient='{recipient}', id='{email_id}'")
            
            # Search for the email in Sent Items using subject and recipient
            search_query = f'subject:"{subject}" AND to:"{recipient}"'
            
            # Try different folder names for Sent Items (localized)
            sent_folders = ["Sent Items", "已傳送的郵件", "寄件備份", "已发送邮件", "送信済みアイテム"]
            
            for folder_name in sent_folders:
                try:
                    logger.debug(f"Searching in folder: {folder_name}")
                    search_results = await self.search_emails(
                        query=search_query,
                        folder=folder_name,
                        limit=10,
                        client_id="verification"
                    )
                    
                    if search_results:
                        # Check if any of the found emails match our criteria
                        for email in search_results:
                            email_subject = email.get("subject", "")
                            email_recipients = email.get("recipients", [])
                            
                            # Check if subject matches and recipient is in the list
                            if (email_subject == subject and 
                                any(recipient.lower() in r.lower() for r in email_recipients)):
                                
                                logger.info(f"Email verification successful: found email in {folder_name}")
                                return {
                                    "status": "verified",
                                    "method": "search_verification",
                                    "found_in_folder": folder_name,
                                    "search_query": search_query,
                                    "found_email_id": email.get("id"),
                                    "verification_time": self._get_current_timestamp()
                                }
                    
                except Exception as e:
                    logger.debug(f"Error searching in folder {folder_name}: {e}")
                    continue
            
            # If not found in any folder, try searching without folder restriction
            try:
                logger.debug("Searching across all folders for verification")
                search_results = await self.search_emails(
                    query=search_query,
                    folder=None,  # Search all folders
                    limit=10,
                    client_id="verification"
                )
                
                if search_results:
                    for email in search_results:
                        email_subject = email.get("subject", "")
                        email_recipients = email.get("recipients", [])
                        
                        if (email_subject == subject and 
                            any(recipient.lower() in r.lower() for r in email_recipients)):
                            
                            logger.info("Email verification successful: found email in general search")
                            return {
                                "status": "verified",
                                "method": "general_search_verification",
                                "found_in_folder": email.get("folder", "unknown"),
                                "search_query": search_query,
                                "found_email_id": email.get("id"),
                                "verification_time": self._get_current_timestamp()
                            }
                
            except Exception as e:
                logger.warning(f"Error in general search verification: {e}")
            
            # If still not found, return unverified status
            logger.warning(f"Email verification failed: could not find sent email with subject '{subject}' to '{recipient}'")
            return {
                "status": "unverified",
                "method": "search_verification",
                "reason": "Email not found in search results",
                "search_query": search_query,
                "searched_folders": sent_folders,
                "verification_time": self._get_current_timestamp(),
                "note": "Email may have been sent but not yet indexed for search, or may be in a different folder"
            }
            
        except Exception as e:
            logger.error(f"Error during email verification: {e}")
            return {
                "status": "verification_error",
                "method": "search_verification",
                "error": str(e),
                "verification_time": self._get_current_timestamp()
            }

    def optimize_performance(self) -> Dict[str, Any]:
        """Perform performance optimization operations."""
        results = {}
        
        try:
            # Clean up memory manager cache
            if self.memory_manager:
                initial_memory = self.memory_manager.get_memory_usage()
                self.memory_manager.clear_cache("email")  # Clear only email cache, keep folders
                final_memory = self.memory_manager.get_memory_usage()
                
                results["memory_cleanup"] = {
                    "initial_memory_mb": initial_memory.get("rss_mb", 0),
                    "final_memory_mb": final_memory.get("rss_mb", 0),
                    "memory_saved_mb": initial_memory.get("rss_mb", 0) - final_memory.get("rss_mb", 0)
                }
            
            # Clean up lazy loader cache
            if self.lazy_loader:
                initial_cached = len(self.lazy_loader._lazy_emails)
                self.lazy_loader.cleanup_cache()
                final_cached = len(self.lazy_loader._lazy_emails)
                
                results["lazy_loader_cleanup"] = {
                    "initial_cached_emails": initial_cached,
                    "final_cached_emails": final_cached,
                    "emails_cleaned": initial_cached - final_cached
                }
            
            logger.info(f"Performance optimization completed: {results}")
            return results
            
        except Exception as e:
            logger.error(f"Error during performance optimization: {str(e)}")
            return {"error": str(e)}