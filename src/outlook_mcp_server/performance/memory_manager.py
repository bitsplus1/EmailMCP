"""Memory management for large email datasets."""

import gc
import logging
import threading
import time
from typing import Dict, Any, Optional, List, TypeVar, Generic
from dataclasses import dataclass
from collections import OrderedDict

try:
    import psutil
except ImportError:
    psutil = None

# Avoid circular import by using TYPE_CHECKING
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from ..models.email_data import EmailData


logger = logging.getLogger(__name__)
T = TypeVar('T')


@dataclass
class MemoryConfig:
    """Configuration for memory management."""
    max_memory_mb: int = 512
    cache_size_limit: int = 1000
    gc_threshold: float = 0.8
    cleanup_interval: int = 60
    enable_compression: bool = True
    max_item_size_mb: int = 10


class LRUCache(Generic[T]):
    """Thread-safe LRU cache with memory management."""
    
    def __init__(self, max_size: int, max_memory_mb: int):
        """Initialize LRU cache."""
        self.max_size = max_size
        self.max_memory_bytes = max_memory_mb * 1024 * 1024
        self._cache: OrderedDict[str, T] = OrderedDict()
        self._memory_usage = 0
        self._lock = threading.RLock()
        self._access_count = 0
        self._hit_count = 0
    
    def get(self, key: str) -> Optional[T]:
        """Get item from cache."""
        with self._lock:
            self._access_count += 1
            
            if key in self._cache:
                value = self._cache.pop(key)
                self._cache[key] = value
                self._hit_count += 1
                return value
            
            return None
    
    def put(self, key: str, value: T, size_bytes: int = 0) -> None:
        """Put item in cache."""
        with self._lock:
            if key in self._cache:
                self._cache.pop(key)
            
            if size_bytes > 0:
                self._memory_usage += size_bytes
            
            self._cache[key] = value
            self._evict_if_needed()
    
    def remove(self, key: str) -> bool:
        """Remove item from cache."""
        with self._lock:
            if key in self._cache:
                self._cache.pop(key)
                return True
            return False
    
    def clear(self) -> None:
        """Clear all items from cache."""
        with self._lock:
            self._cache.clear()
            self._memory_usage = 0
    
    def _evict_if_needed(self) -> None:
        """Evict items if cache exceeds limits."""
        while len(self._cache) > self.max_size:
            self._cache.popitem(last=False)
        
        while self._memory_usage > self.max_memory_bytes and self._cache:
            self._cache.popitem(last=False)
            self._memory_usage *= 0.9
    
    def get_stats(self) -> Dict[str, Any]:
        """Get cache statistics."""
        with self._lock:
            hit_rate = (self._hit_count / self._access_count) if self._access_count > 0 else 0
            return {
                "size": len(self._cache),
                "max_size": self.max_size,
                "memory_usage_mb": self._memory_usage / (1024 * 1024),
                "max_memory_mb": self.max_memory_bytes / (1024 * 1024),
                "hit_rate": hit_rate,
                "access_count": self._access_count,
                "hit_count": self._hit_count
            }


class MemoryManager:
    """Manages memory usage for large email datasets."""
    
    def __init__(self, config: MemoryConfig):
        """Initialize memory manager."""
        self.config = config
        
        self.email_cache: LRUCache['EmailData'] = LRUCache(
            max_size=config.cache_size_limit,
            max_memory_mb=config.max_memory_mb // 2
        )
        
        self.attachment_cache: LRUCache[bytes] = LRUCache(
            max_size=config.cache_size_limit // 10,
            max_memory_mb=config.max_memory_mb // 4
        )
        
        self.folder_cache: LRUCache[Any] = LRUCache(
            max_size=100,
            max_memory_mb=config.max_memory_mb // 8
        )
        
        if psutil:
            self._process = psutil.Process()
        else:
            self._process = None
            
        self._lock = threading.RLock()
        self._shutdown = False
        
        self._stats = {
            "memory_cleanups": 0,
            "gc_collections": 0,
            "cache_evictions": 0,
            "compression_saves": 0
        }
        
        self._monitor_thread = threading.Thread(
            target=self._memory_monitor,
            daemon=True,
            name="memory-monitor"
        )
        self._monitor_thread.start()
        
        logger.info(f"Memory manager initialized: {config}")
    
    def cache_email(self, email_id: str, email_data: 'EmailData') -> None:
        """Cache email data with memory management."""
        try:
            size_bytes = self._estimate_email_size(email_data)
            
            if size_bytes > self.config.max_item_size_mb * 1024 * 1024:
                logger.warning(f"Email {email_id} too large to cache: {size_bytes / (1024*1024):.1f}MB")
                return
            
            if self.config.enable_compression and size_bytes > 1024:
                email_data = self._compress_email_data(email_data)
                self._stats["compression_saves"] += 1
            
            self.email_cache.put(email_id, email_data, size_bytes)
            logger.debug(f"Cached email {email_id} ({size_bytes} bytes)")
            
        except Exception as e:
            logger.error(f"Error caching email {email_id}: {str(e)}")
    
    def get_cached_email(self, email_id: str) -> Optional['EmailData']:
        """Get cached email data."""
        try:
            email_data = self.email_cache.get(email_id)
            if email_data:
                logger.debug(f"Cache hit for email {email_id}")
                if hasattr(email_data, '_compressed'):
                    email_data = self._decompress_email_data(email_data)
            return email_data
        except Exception as e:
            logger.error(f"Error retrieving cached email {email_id}: {str(e)}")
            return None
    
    def cache_attachment(self, attachment_id: str, data: bytes) -> None:
        """Cache attachment data."""
        try:
            size_bytes = len(data)
            
            if size_bytes > self.config.max_item_size_mb * 1024 * 1024:
                logger.warning(f"Attachment {attachment_id} too large to cache")
                return
            
            if self.config.enable_compression and size_bytes > 10240:
                data = self._compress_data(data)
                self._stats["compression_saves"] += 1
            
            self.attachment_cache.put(attachment_id, data, size_bytes)
            logger.debug(f"Cached attachment {attachment_id} ({size_bytes} bytes)")
            
        except Exception as e:
            logger.error(f"Error caching attachment {attachment_id}: {str(e)}")
    
    def get_cached_attachment(self, attachment_id: str) -> Optional[bytes]:
        """Get cached attachment data."""
        try:
            data = self.attachment_cache.get(attachment_id)
            if data and hasattr(data, '_compressed'):
                data = self._decompress_data(data)
            return data
        except Exception as e:
            logger.error(f"Error retrieving cached attachment {attachment_id}: {str(e)}")
            return None
    
    def clear_cache(self, cache_type: str = "all") -> None:
        """Clear specified cache."""
        with self._lock:
            if cache_type in ("all", "email"):
                self.email_cache.clear()
                logger.info("Cleared email cache")
            
            if cache_type in ("all", "attachment"):
                self.attachment_cache.clear()
                logger.info("Cleared attachment cache")
            
            if cache_type in ("all", "folder"):
                self.folder_cache.clear()
                logger.info("Cleared folder cache")
    
    def get_memory_usage(self) -> Dict[str, Any]:
        """Get current memory usage information."""
        try:
            if self._process:
                memory_info = self._process.memory_info()
                memory_percent = self._process.memory_percent()
                available_mb = psutil.virtual_memory().available / (1024 * 1024)
            else:
                memory_info = type('obj', (object,), {'rss': 0, 'vms': 0})()
                memory_percent = 0.0
                available_mb = 1024.0
            
            return {
                "rss_mb": memory_info.rss / (1024 * 1024),
                "vms_mb": memory_info.vms / (1024 * 1024),
                "percent": memory_percent,
                "available_mb": available_mb,
                "email_cache": self.email_cache.get_stats(),
                "attachment_cache": self.attachment_cache.get_stats(),
                "folder_cache": self.folder_cache.get_stats()
            }
        except Exception as e:
            logger.error(f"Error getting memory usage: {str(e)}")
            return {}
    
    def _estimate_email_size(self, email_data: 'EmailData') -> int:
        """Estimate memory size of email data."""
        try:
            size = 0
            size += len(email_data.subject or "") * 2
            size += len(email_data.sender or "") * 2
            size += len(email_data.body or "") * 2
            size += len(email_data.body_html or "") * 2
            
            if email_data.recipients:
                size += sum(len(r) * 2 for r in email_data.recipients)
            
            size += 1024
            return size
            
        except Exception as e:
            logger.debug(f"Error estimating email size: {str(e)}")
            return 1024
    
    def _compress_email_data(self, email_data: 'EmailData') -> 'EmailData':
        """Compress email data if beneficial."""
        try:
            import zlib
            
            if email_data.body and len(email_data.body) > 1024:
                compressed_body = zlib.compress(email_data.body.encode('utf-8'))
                if len(compressed_body) < len(email_data.body):
                    email_data.body = compressed_body
                    email_data._body_compressed = True
            
            if email_data.body_html and len(email_data.body_html) > 1024:
                compressed_html = zlib.compress(email_data.body_html.encode('utf-8'))
                if len(compressed_html) < len(email_data.body_html):
                    email_data.body_html = compressed_html
                    email_data._html_compressed = True
            
            email_data._compressed = True
            return email_data
            
        except Exception as e:
            logger.debug(f"Error compressing email data: {str(e)}")
            return email_data
    
    def _decompress_email_data(self, email_data: 'EmailData') -> 'EmailData':
        """Decompress email data."""
        try:
            import zlib
            
            if hasattr(email_data, '_body_compressed') and email_data._body_compressed:
                email_data.body = zlib.decompress(email_data.body).decode('utf-8')
                delattr(email_data, '_body_compressed')
            
            if hasattr(email_data, '_html_compressed') and email_data._html_compressed:
                email_data.body_html = zlib.decompress(email_data.body_html).decode('utf-8')
                delattr(email_data, '_html_compressed')
            
            if hasattr(email_data, '_compressed'):
                delattr(email_data, '_compressed')
            
            return email_data
            
        except Exception as e:
            logger.error(f"Error decompressing email data: {str(e)}")
            return email_data
    
    def _compress_data(self, data: bytes) -> bytes:
        """Compress binary data."""
        try:
            import zlib
            compressed = zlib.compress(data)
            if len(compressed) < len(data):
                compressed._compressed = True
                return compressed
            return data
        except Exception as e:
            logger.debug(f"Error compressing data: {str(e)}")
            return data
    
    def _decompress_data(self, data: bytes) -> bytes:
        """Decompress binary data."""
        try:
            import zlib
            if hasattr(data, '_compressed'):
                return zlib.decompress(data)
            return data
        except Exception as e:
            logger.error(f"Error decompressing data: {str(e)}")
            return data
    
    def _memory_monitor(self) -> None:
        """Background memory monitoring and cleanup."""
        logger.debug("Starting memory monitor")
        
        while not self._shutdown:
            try:
                time.sleep(self.config.cleanup_interval)
                
                if self._shutdown:
                    break
                
                self._check_memory_usage()
                
            except Exception as e:
                logger.error(f"Error in memory monitor: {str(e)}")
    
    def _check_memory_usage(self) -> None:
        """Check memory usage and perform cleanup if needed."""
        try:
            memory_info = self.get_memory_usage()
            memory_mb = memory_info.get("rss_mb", 0)
            
            if memory_mb > self.config.max_memory_mb * self.config.gc_threshold:
                logger.warning(f"Memory usage high: {memory_mb:.1f}MB")
                self._perform_cleanup()
            
        except Exception as e:
            logger.error(f"Error checking memory usage: {str(e)}")
    
    def _perform_cleanup(self) -> None:
        """Perform memory cleanup operations."""
        with self._lock:
            logger.info("Performing memory cleanup")
            
            initial_email_size = len(self.email_cache._cache)
            initial_attachment_size = len(self.attachment_cache._cache)
            
            target_email_size = int(initial_email_size * 0.75)
            target_attachment_size = int(initial_attachment_size * 0.75)
            
            while len(self.email_cache._cache) > target_email_size:
                self.email_cache._cache.popitem(last=False)
            
            while len(self.attachment_cache._cache) > target_attachment_size:
                self.attachment_cache._cache.popitem(last=False)
            
            collected = gc.collect()
            
            self._stats["memory_cleanups"] += 1
            self._stats["gc_collections"] += collected
            
            logger.info(f"Memory cleanup complete. GC collected {collected} objects")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get memory manager statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update(self.get_memory_usage())
            return stats
    
    def shutdown(self) -> None:
        """Shutdown memory manager."""
        logger.info("Shutting down memory manager")
        
        self._shutdown = True
        
        if self._monitor_thread.is_alive():
            self._monitor_thread.join(timeout=5.0)
        
        self.clear_cache()
        
        logger.info("Memory manager shutdown complete")