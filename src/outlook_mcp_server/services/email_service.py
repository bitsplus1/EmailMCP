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
        
    async def list_emails(self, folder: str = None, unread_only: bool = False, limit: int = 50, client_id: str = "default") -> List[Dict[str, Any]]:
        """
        List emails with filtering and pagination.
        
        Args:
            folder: Optional folder name to list emails from
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
            
            logger.info(f"Listing emails - folder: {folder}, unread_only: {unread_only}, limit: {limit}")
            
            # Validate parameters
            if limit <= 0:
                limit = 50
            elif limit > 1000:
                limit = 1000
            
            # Check cache first if memory manager is available
            cache_key = f"list_emails:{folder}:{unread_only}:{limit}"
            if self.memory_manager:
                cached_result = self.memory_manager.folder_cache.get(cache_key)
                if cached_result:
                    self._stats["cache_hits"] += 1
                    logger.debug(f"Cache hit for email list: {cache_key}")
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
                    
                    email_data_list = temp_adapter.list_emails(folder, unread_only, limit)
            else:
                # Ensure we're connected to Outlook
                if not self.outlook_adapter.is_connected():
                    logger.error("Outlook adapter is not connected")
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Get emails from the adapter
                email_data_list = self.outlook_adapter.list_emails(folder, unread_only, limit)
            
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
                raise PermissionError(folder or "emails", f"Access denied: {str(e)}")
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
    
    async def search_emails(self, query: str, folder: str = None, limit: int = 50, client_id: str = "default") -> List[Dict[str, Any]]:
        """
        Search emails based on user-defined queries.
        
        Args:
            query: Search query string
            folder: Optional folder name to search in
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
            
            logger.info(f"Searching emails - query: '{query}', folder: {folder}, limit: {limit}")
            
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
            cache_key = f"search:{hash(query)}:{folder}:{limit}"
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
                    
                    email_data_list = temp_adapter.search_emails(query, folder, limit)
            else:
                # Ensure we're connected
                if not self.outlook_adapter.is_connected():
                    raise OutlookConnectionError("Not connected to Outlook")
                
                # Perform search using adapter
                email_data_list = self.outlook_adapter.search_emails(query, folder, limit)
            
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
                    logger.warning(f"Error transforming search result email '{email_data.id}': {str(e)}")
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
                raise PermissionError(folder or "emails", f"Access denied during search: {str(e)}")
            elif any(keyword in str(e).lower() for keyword in ['not found', 'invalid', 'missing']) and folder:
                raise FolderNotFoundError(folder)
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