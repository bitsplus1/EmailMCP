"""Lazy loading for email content and attachments."""

import logging
import threading
import time
import weakref
from typing import Optional, Dict, Any, List, Callable, Union
from dataclasses import dataclass, field
from concurrent.futures import ThreadPoolExecutor, Future
import asyncio
from ..models.email_data import EmailData
from ..models.exceptions import EmailNotFoundError, OutlookConnectionError


logger = logging.getLogger(__name__)


@dataclass
class LazyLoadConfig:
    """Configuration for lazy loading."""
    max_workers: int = 3  # Maximum number of background loading threads
    cache_timeout: int = 300  # Cache timeout in seconds
    preload_threshold: int = 5  # Preload when this many items requested
    batch_size: int = 10  # Batch size for bulk operations
    enable_prefetch: bool = True  # Enable predictive prefetching


class LazyEmailContent:
    """Lazy-loaded email content wrapper."""
    
    def __init__(self, email_id: str, loader_func: Callable[[], EmailData]):
        """Initialize lazy email content."""
        self.email_id = email_id
        self._loader_func = loader_func
        self._content: Optional[EmailData] = None
        self._loading = False
        self._loaded_at: Optional[float] = None
        self._error: Optional[Exception] = None
        self._lock = threading.Lock()
    
    @property
    def is_loaded(self) -> bool:
        """Check if content is loaded."""
        return self._content is not None
    
    @property
    def is_loading(self) -> bool:
        """Check if content is currently loading."""
        return self._loading
    
    @property
    def has_error(self) -> bool:
        """Check if loading failed."""
        return self._error is not None
    
    def get_content(self, timeout: float = 10.0) -> EmailData:
        """
        Get email content, loading if necessary.
        
        Args:
            timeout: Timeout for loading operation
            
        Returns:
            EmailData: The loaded email content
            
        Raises:
            EmailNotFoundError: If email cannot be loaded
            OutlookConnectionError: If connection fails
        """
        with self._lock:
            # Return cached content if available and fresh
            if self._content and self._is_cache_valid():
                return self._content
            
            # Check for previous errors
            if self._error:
                raise self._error
            
            # Load content if not already loading
            if not self._loading:
                self._loading = True
                try:
                    logger.debug(f"Loading email content for {self.email_id}")
                    self._content = self._loader_func()
                    self._loaded_at = time.time()
                    self._error = None
                    logger.debug(f"Successfully loaded email content for {self.email_id}")
                except Exception as e:
                    self._error = e
                    logger.error(f"Failed to load email content for {self.email_id}: {str(e)}")
                    raise
                finally:
                    self._loading = False
            
            if self._content:
                return self._content
            elif self._error:
                raise self._error
            else:
                raise EmailNotFoundError(self.email_id)
    
    def _is_cache_valid(self) -> bool:
        """Check if cached content is still valid."""
        if not self._loaded_at:
            return False
        
        # Content is valid for 5 minutes
        return time.time() - self._loaded_at < 300
    
    def invalidate(self) -> None:
        """Invalidate cached content."""
        with self._lock:
            self._content = None
            self._loaded_at = None
            self._error = None
            self._loading = False


class LazyAttachmentContent:
    """Lazy-loaded attachment content wrapper."""
    
    def __init__(self, attachment_id: str, loader_func: Callable[[], bytes]):
        """Initialize lazy attachment content."""
        self.attachment_id = attachment_id
        self._loader_func = loader_func
        self._content: Optional[bytes] = None
        self._loading = False
        self._loaded_at: Optional[float] = None
        self._error: Optional[Exception] = None
        self._lock = threading.Lock()
    
    @property
    def is_loaded(self) -> bool:
        """Check if content is loaded."""
        return self._content is not None
    
    @property
    def is_loading(self) -> bool:
        """Check if content is currently loading."""
        return self._loading
    
    @property
    def size(self) -> Optional[int]:
        """Get attachment size if loaded."""
        return len(self._content) if self._content else None
    
    def get_content(self, timeout: float = 30.0) -> bytes:
        """
        Get attachment content, loading if necessary.
        
        Args:
            timeout: Timeout for loading operation
            
        Returns:
            bytes: The loaded attachment content
        """
        with self._lock:
            # Return cached content if available and fresh
            if self._content and self._is_cache_valid():
                return self._content
            
            # Check for previous errors
            if self._error:
                raise self._error
            
            # Load content if not already loading
            if not self._loading:
                self._loading = True
                try:
                    logger.debug(f"Loading attachment content for {self.attachment_id}")
                    self._content = self._loader_func()
                    self._loaded_at = time.time()
                    self._error = None
                    logger.debug(f"Successfully loaded attachment {self.attachment_id} ({len(self._content)} bytes)")
                except Exception as e:
                    self._error = e
                    logger.error(f"Failed to load attachment {self.attachment_id}: {str(e)}")
                    raise
                finally:
                    self._loading = False
            
            if self._content:
                return self._content
            elif self._error:
                raise self._error
            else:
                raise Exception(f"Failed to load attachment {self.attachment_id}")
    
    def _is_cache_valid(self) -> bool:
        """Check if cached content is still valid."""
        if not self._loaded_at:
            return False
        
        # Attachment content is valid for 10 minutes
        return time.time() - self._loaded_at < 600
    
    def invalidate(self) -> None:
        """Invalidate cached content."""
        with self._lock:
            self._content = None
            self._loaded_at = None
            self._error = None
            self._loading = False


class LazyEmailLoader:
    """Manages lazy loading of email content."""
    
    def __init__(self, config: LazyLoadConfig, outlook_adapter):
        """Initialize lazy email loader."""
        self.config = config
        self.outlook_adapter = outlook_adapter
        
        # Thread pool for background loading
        self._executor = ThreadPoolExecutor(
            max_workers=config.max_workers,
            thread_name_prefix="lazy-email-loader"
        )
        
        # Cache of lazy content objects
        self._lazy_emails: Dict[str, LazyEmailContent] = {}
        self._access_patterns: Dict[str, List[float]] = {}
        self._lock = threading.RLock()
        
        # Statistics
        self._stats = {
            "emails_loaded": 0,
            "cache_hits": 0,
            "cache_misses": 0,
            "background_loads": 0,
            "prefetch_loads": 0
        }
        
        logger.info(f"Lazy email loader initialized: {config}")
    
    def get_lazy_email(self, email_id: str) -> LazyEmailContent:
        """
        Get lazy email content wrapper.
        
        Args:
            email_id: Email identifier
            
        Returns:
            LazyEmailContent: Lazy content wrapper
        """
        with self._lock:
            # Record access pattern
            self._record_access(email_id)
            
            # Return existing lazy content if available
            if email_id in self._lazy_emails:
                self._stats["cache_hits"] += 1
                return self._lazy_emails[email_id]
            
            # Create new lazy content
            self._stats["cache_misses"] += 1
            
            def loader():
                return self.outlook_adapter.get_email_by_id(email_id)
            
            lazy_content = LazyEmailContent(email_id, loader)
            self._lazy_emails[email_id] = lazy_content
            
            # Check if we should prefetch related emails
            if self.config.enable_prefetch:
                self._consider_prefetch(email_id)
            
            return lazy_content
    
    def preload_emails(self, email_ids: List[str]) -> Dict[str, Future]:
        """
        Preload multiple emails in background.
        
        Args:
            email_ids: List of email IDs to preload
            
        Returns:
            Dict[str, Future]: Mapping of email IDs to futures
        """
        futures = {}
        
        with self._lock:
            for email_id in email_ids:
                if email_id not in self._lazy_emails:
                    # Create lazy content
                    def loader():
                        return self.outlook_adapter.get_email_by_id(email_id)
                    
                    lazy_content = LazyEmailContent(email_id, loader)
                    self._lazy_emails[email_id] = lazy_content
                
                # Submit background loading task
                if not self._lazy_emails[email_id].is_loaded:
                    future = self._executor.submit(self._background_load_email, email_id)
                    futures[email_id] = future
                    self._stats["background_loads"] += 1
        
        logger.debug(f"Started preloading {len(futures)} emails")
        return futures
    
    def _background_load_email(self, email_id: str) -> None:
        """Load email in background thread."""
        try:
            lazy_content = self._lazy_emails.get(email_id)
            if lazy_content and not lazy_content.is_loaded:
                lazy_content.get_content()
                self._stats["emails_loaded"] += 1
                logger.debug(f"Background loaded email {email_id}")
        except Exception as e:
            logger.error(f"Background loading failed for email {email_id}: {str(e)}")
    
    def _record_access(self, email_id: str) -> None:
        """Record email access for pattern analysis."""
        now = time.time()
        
        if email_id not in self._access_patterns:
            self._access_patterns[email_id] = []
        
        self._access_patterns[email_id].append(now)
        
        # Keep only recent accesses (last hour)
        cutoff = now - 3600
        self._access_patterns[email_id] = [
            t for t in self._access_patterns[email_id] if t > cutoff
        ]
    
    def _consider_prefetch(self, email_id: str) -> None:
        """Consider prefetching related emails based on access patterns."""
        try:
            # Simple heuristic: if this email is accessed frequently,
            # prefetch emails from the same folder or thread
            
            access_count = len(self._access_patterns.get(email_id, []))
            if access_count >= self.config.preload_threshold:
                # This is a frequently accessed email, consider prefetching
                self._prefetch_related_emails(email_id)
        
        except Exception as e:
            logger.debug(f"Error considering prefetch for {email_id}: {str(e)}")
    
    def _prefetch_related_emails(self, email_id: str) -> None:
        """Prefetch emails related to the given email."""
        try:
            # Get the email to analyze
            lazy_content = self._lazy_emails.get(email_id)
            if not lazy_content or not lazy_content.is_loaded:
                return
            
            email_data = lazy_content.get_content()
            
            # Simple strategy: prefetch recent emails from same folder
            if hasattr(email_data, 'folder_name') and email_data.folder_name:
                # This would require integration with the email service
                # For now, just log the intent
                logger.debug(f"Would prefetch emails from folder: {email_data.folder_name}")
                self._stats["prefetch_loads"] += 1
        
        except Exception as e:
            logger.debug(f"Error prefetching related emails for {email_id}: {str(e)}")
    
    def cleanup_cache(self, max_age: int = 600) -> None:
        """Clean up old cached content."""
        with self._lock:
            now = time.time()
            expired_emails = []
            
            for email_id, lazy_content in self._lazy_emails.items():
                if (lazy_content._loaded_at and 
                    now - lazy_content._loaded_at > max_age):
                    expired_emails.append(email_id)
            
            for email_id in expired_emails:
                self._lazy_emails[email_id].invalidate()
                del self._lazy_emails[email_id]
            
            # Clean up access patterns
            cutoff = now - 3600  # Keep 1 hour of history
            for email_id in list(self._access_patterns.keys()):
                self._access_patterns[email_id] = [
                    t for t in self._access_patterns[email_id] if t > cutoff
                ]
                if not self._access_patterns[email_id]:
                    del self._access_patterns[email_id]
            
            logger.debug(f"Cleaned up {len(expired_emails)} expired email cache entries")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get loader statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update({
                "cached_emails": len(self._lazy_emails),
                "loaded_emails": sum(1 for lc in self._lazy_emails.values() if lc.is_loaded),
                "loading_emails": sum(1 for lc in self._lazy_emails.values() if lc.is_loading),
                "access_patterns": len(self._access_patterns)
            })
            return stats
    
    def shutdown(self) -> None:
        """Shutdown the lazy loader."""
        logger.info("Shutting down lazy email loader")
        
        self._executor.shutdown(wait=True, timeout=10.0)
        
        with self._lock:
            self._lazy_emails.clear()
            self._access_patterns.clear()
        
        logger.info("Lazy email loader shutdown complete")


class LazyAttachmentLoader:
    """Manages lazy loading of email attachments."""
    
    def __init__(self, config: LazyLoadConfig, outlook_adapter):
        """Initialize lazy attachment loader."""
        self.config = config
        self.outlook_adapter = outlook_adapter
        
        # Thread pool for background loading
        self._executor = ThreadPoolExecutor(
            max_workers=max(1, config.max_workers // 2),  # Fewer threads for attachments
            thread_name_prefix="lazy-attachment-loader"
        )
        
        # Cache of lazy attachment objects
        self._lazy_attachments: Dict[str, LazyAttachmentContent] = {}
        self._lock = threading.RLock()
        
        # Statistics
        self._stats = {
            "attachments_loaded": 0,
            "cache_hits": 0,
            "cache_misses": 0,
            "total_bytes_loaded": 0
        }
        
        logger.info(f"Lazy attachment loader initialized")
    
    def get_lazy_attachment(self, attachment_id: str, email_id: str, attachment_name: str) -> LazyAttachmentContent:
        """
        Get lazy attachment content wrapper.
        
        Args:
            attachment_id: Attachment identifier
            email_id: Parent email identifier
            attachment_name: Attachment filename
            
        Returns:
            LazyAttachmentContent: Lazy content wrapper
        """
        with self._lock:
            # Return existing lazy content if available
            if attachment_id in self._lazy_attachments:
                self._stats["cache_hits"] += 1
                return self._lazy_attachments[attachment_id]
            
            # Create new lazy content
            self._stats["cache_misses"] += 1
            
            def loader():
                # This would need to be implemented in the outlook adapter
                # For now, return empty bytes as placeholder
                logger.warning(f"Attachment loading not yet implemented: {attachment_name}")
                return b""
            
            lazy_content = LazyAttachmentContent(attachment_id, loader)
            self._lazy_attachments[attachment_id] = lazy_content
            
            return lazy_content
    
    def preload_attachments(self, attachment_specs: List[Dict[str, str]]) -> Dict[str, Future]:
        """
        Preload multiple attachments in background.
        
        Args:
            attachment_specs: List of attachment specifications
            
        Returns:
            Dict[str, Future]: Mapping of attachment IDs to futures
        """
        futures = {}
        
        with self._lock:
            for spec in attachment_specs:
                attachment_id = spec.get("attachment_id")
                email_id = spec.get("email_id")
                attachment_name = spec.get("name", "unknown")
                
                if not attachment_id:
                    continue
                
                if attachment_id not in self._lazy_attachments:
                    # Create lazy content
                    def loader():
                        logger.warning(f"Attachment loading not yet implemented: {attachment_name}")
                        return b""
                    
                    lazy_content = LazyAttachmentContent(attachment_id, loader)
                    self._lazy_attachments[attachment_id] = lazy_content
                
                # Submit background loading task
                if not self._lazy_attachments[attachment_id].is_loaded:
                    future = self._executor.submit(self._background_load_attachment, attachment_id)
                    futures[attachment_id] = future
        
        logger.debug(f"Started preloading {len(futures)} attachments")
        return futures
    
    def _background_load_attachment(self, attachment_id: str) -> None:
        """Load attachment in background thread."""
        try:
            lazy_content = self._lazy_attachments.get(attachment_id)
            if lazy_content and not lazy_content.is_loaded:
                content = lazy_content.get_content()
                self._stats["attachments_loaded"] += 1
                self._stats["total_bytes_loaded"] += len(content)
                logger.debug(f"Background loaded attachment {attachment_id} ({len(content)} bytes)")
        except Exception as e:
            logger.error(f"Background loading failed for attachment {attachment_id}: {str(e)}")
    
    def cleanup_cache(self, max_age: int = 1200) -> None:  # 20 minutes for attachments
        """Clean up old cached attachments."""
        with self._lock:
            now = time.time()
            expired_attachments = []
            
            for attachment_id, lazy_content in self._lazy_attachments.items():
                if (lazy_content._loaded_at and 
                    now - lazy_content._loaded_at > max_age):
                    expired_attachments.append(attachment_id)
            
            for attachment_id in expired_attachments:
                self._lazy_attachments[attachment_id].invalidate()
                del self._lazy_attachments[attachment_id]
            
            logger.debug(f"Cleaned up {len(expired_attachments)} expired attachment cache entries")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get attachment loader statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update({
                "cached_attachments": len(self._lazy_attachments),
                "loaded_attachments": sum(1 for lc in self._lazy_attachments.values() if lc.is_loaded),
                "loading_attachments": sum(1 for lc in self._lazy_attachments.values() if lc.is_loading)
            })
            return stats
    
    def shutdown(self) -> None:
        """Shutdown the lazy attachment loader."""
        logger.info("Shutting down lazy attachment loader")
        
        self._executor.shutdown(wait=True, timeout=10.0)
        
        with self._lock:
            self._lazy_attachments.clear()
        
        logger.info("Lazy attachment loader shutdown complete")