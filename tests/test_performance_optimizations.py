"""Tests for performance optimizations and resource management."""

import pytest
import asyncio
import time
import threading
from unittest.mock import Mock, patch, MagicMock
from src.outlook_mcp_server.performance.memory_manager import MemoryManager, MemoryConfig, LRUCache
from src.outlook_mcp_server.performance.rate_limiter import RateLimiter, RateLimitConfig, TokenBucket
from src.outlook_mcp_server.performance.lazy_loader import LazyEmailLoader, LazyAttachmentLoader, LazyLoadConfig
from src.outlook_mcp_server.adapters.connection_pool import OutlookConnectionPool, OutlookConnection
from src.outlook_mcp_server.models.email_data import EmailData
from src.outlook_mcp_server.models.exceptions import ValidationError


class TestMemoryManager:
    """Test memory management functionality."""
    
    def test_memory_config_creation(self):
        """Test memory configuration creation."""
        config = MemoryConfig(
            max_memory_mb=256,
            cache_size_limit=500,
            gc_threshold=0.7
        )
        
        assert config.max_memory_mb == 256
        assert config.cache_size_limit == 500
        assert config.gc_threshold == 0.7
    
    def test_lru_cache_basic_operations(self):
        """Test basic LRU cache operations."""
        cache = LRUCache[str](max_size=3, max_memory_mb=1)
        
        # Test put and get
        cache.put("key1", "value1", 100)
        cache.put("key2", "value2", 100)
        cache.put("key3", "value3", 100)
        
        assert cache.get("key1") == "value1"
        assert cache.get("key2") == "value2"
        assert cache.get("key3") == "value3"
        
        # Test LRU eviction
        cache.put("key4", "value4", 100)  # Should evict key1 (least recently used)
        
        assert cache.get("key1") is None
        assert cache.get("key4") == "value4"
    
    def test_lru_cache_memory_eviction(self):
        """Test memory-based eviction in LRU cache."""
        cache = LRUCache[str](max_size=100, max_memory_mb=1)  # 1MB limit
        
        # Add items that exceed memory limit
        large_value = "x" * (512 * 1024)  # 512KB
        cache.put("key1", large_value, len(large_value))
        cache.put("key2", large_value, len(large_value))
        cache.put("key3", large_value, len(large_value))  # Should trigger eviction
        
        # Should have evicted some items due to memory pressure
        assert len(cache._cache) < 3
    
    def test_memory_manager_initialization(self):
        """Test memory manager initialization."""
        config = MemoryConfig(max_memory_mb=128)
        
        with patch('psutil.Process'):
            manager = MemoryManager(config)
            
            assert manager.config == config
            assert manager.email_cache is not None
            assert manager.attachment_cache is not None
            assert manager.folder_cache is not None
    
    def test_email_caching(self):
        """Test email caching functionality."""
        config = MemoryConfig(max_memory_mb=128)
        
        with patch('psutil.Process'):
            manager = MemoryManager(config)
            
            # Create test email data
            email_data = EmailData(
                id="test-email-1",
                subject="Test Subject",
                sender="test@example.com",
                body="Test body content"
            )
            
            # Cache the email
            manager.cache_email("test-email-1", email_data)
            
            # Retrieve from cache
            cached_email = manager.get_cached_email("test-email-1")
            
            assert cached_email is not None
            assert cached_email.id == "test-email-1"
            assert cached_email.subject == "Test Subject"
    
    def test_attachment_caching(self):
        """Test attachment caching functionality."""
        config = MemoryConfig(max_memory_mb=128)
        
        with patch('psutil.Process'):
            manager = MemoryManager(config)
            
            # Test attachment data
            attachment_data = b"test attachment content"
            
            # Cache the attachment
            manager.cache_attachment("attachment-1", attachment_data)
            
            # Retrieve from cache
            cached_attachment = manager.get_cached_attachment("attachment-1")
            
            assert cached_attachment == attachment_data
    
    def test_cache_size_limits(self):
        """Test cache size limit enforcement."""
        config = MemoryConfig(
            max_memory_mb=1,  # Very small limit
            max_item_size_mb=1
        )
        
        with patch('psutil.Process'):
            manager = MemoryManager(config)
            
            # Try to cache a large item
            large_data = b"x" * (2 * 1024 * 1024)  # 2MB
            
            # Should not cache due to size limit
            manager.cache_attachment("large-attachment", large_data)
            cached = manager.get_cached_attachment("large-attachment")
            
            assert cached is None


class TestRateLimiter:
    """Test rate limiting functionality."""
    
    def test_rate_limit_config_creation(self):
        """Test rate limit configuration creation."""
        config = RateLimitConfig(
            requests_per_second=5.0,
            requests_per_minute=100,
            burst_size=10
        )
        
        assert config.requests_per_second == 5.0
        assert config.requests_per_minute == 100
        assert config.burst_size == 10
    
    def test_token_bucket_basic_operations(self):
        """Test token bucket basic operations."""
        bucket = TokenBucket(capacity=10, refill_rate=2.0)
        
        # Should be able to consume tokens initially
        assert bucket.consume(5) is True
        assert bucket.consume(5) is True
        
        # Should not be able to consume more tokens
        assert bucket.consume(1) is False
        
        # Wait for refill and try again
        time.sleep(1.0)
        assert bucket.consume(1) is True
    
    def test_token_bucket_wait_time(self):
        """Test token bucket wait time calculation."""
        bucket = TokenBucket(capacity=10, refill_rate=2.0)
        
        # Consume all tokens
        bucket.consume(10)
        
        # Should need to wait for tokens
        wait_time = bucket.get_wait_time(2)
        assert wait_time > 0
        assert wait_time <= 1.0  # Should be around 1 second for 2 tokens at 2/sec
    
    @pytest.mark.asyncio
    async def test_rate_limiter_basic_acquire(self):
        """Test basic rate limiter acquire functionality."""
        config = RateLimitConfig(
            requests_per_second=10.0,
            requests_per_minute=100,
            burst_size=5
        )
        
        rate_limiter = RateLimiter(config)
        
        try:
            # Should be able to acquire initially
            result = await rate_limiter.acquire("test-client", "test-method")
            assert result is True
            
            # Should be able to acquire multiple times within burst
            for _ in range(4):
                result = await rate_limiter.acquire("test-client", "test-method")
                assert result is True
        
        finally:
            rate_limiter.shutdown()
    
    @pytest.mark.asyncio
    async def test_rate_limiter_burst_limit(self):
        """Test rate limiter burst limit enforcement."""
        config = RateLimitConfig(
            requests_per_second=1.0,  # Very low rate
            requests_per_minute=60,
            burst_size=2  # Small burst
        )
        
        rate_limiter = RateLimiter(config)
        
        try:
            # Should be able to acquire up to burst size
            await rate_limiter.acquire("test-client", "test-method")
            await rate_limiter.acquire("test-client", "test-method")
            
            # Next acquire should be delayed
            start_time = time.time()
            await rate_limiter.acquire("test-client", "test-method", timeout=2.0)
            elapsed = time.time() - start_time
            
            # Should have waited for token refill
            assert elapsed > 0.5  # Should wait at least 0.5 seconds
        
        finally:
            rate_limiter.shutdown()
    
    @pytest.mark.asyncio
    async def test_rate_limiter_timeout(self):
        """Test rate limiter timeout handling."""
        config = RateLimitConfig(
            requests_per_second=0.1,  # Very slow rate
            burst_size=1
        )
        
        rate_limiter = RateLimiter(config)
        
        try:
            # Consume the burst
            await rate_limiter.acquire("test-client", "test-method")
            
            # Next acquire should timeout
            with pytest.raises(ValidationError):
                await rate_limiter.acquire("test-client", "test-method", timeout=0.5)
        
        finally:
            rate_limiter.shutdown()


class TestLazyLoader:
    """Test lazy loading functionality."""
    
    def test_lazy_load_config_creation(self):
        """Test lazy load configuration creation."""
        config = LazyLoadConfig(
            max_workers=2,
            cache_timeout=600,
            batch_size=5
        )
        
        assert config.max_workers == 2
        assert config.cache_timeout == 600
        assert config.batch_size == 5
    
    def test_lazy_email_content_creation(self):
        """Test lazy email content wrapper creation."""
        def mock_loader():
            return EmailData(
                id="test-email",
                subject="Test Subject",
                sender="test@example.com"
            )
        
        lazy_content = LazyEmailLoader.LazyEmailContent("test-email", mock_loader)
        
        assert lazy_content.email_id == "test-email"
        assert not lazy_content.is_loaded
        assert not lazy_content.is_loading
    
    def test_lazy_email_content_loading(self):
        """Test lazy email content loading."""
        test_email = EmailData(
            id="test-email",
            subject="Test Subject",
            sender="test@example.com"
        )
        
        def mock_loader():
            return test_email
        
        from src.outlook_mcp_server.performance.lazy_loader import LazyEmailContent
        lazy_content = LazyEmailContent("test-email", mock_loader)
        
        # Load content
        loaded_email = lazy_content.get_content()
        
        assert lazy_content.is_loaded
        assert loaded_email.id == "test-email"
        assert loaded_email.subject == "Test Subject"
    
    def test_lazy_email_loader_initialization(self):
        """Test lazy email loader initialization."""
        config = LazyLoadConfig(max_workers=2)
        mock_adapter = Mock()
        
        loader = LazyEmailLoader(config, mock_adapter)
        
        assert loader.config == config
        assert loader.outlook_adapter == mock_adapter
        
        # Cleanup
        loader.shutdown()
    
    def test_lazy_email_loader_get_lazy_email(self):
        """Test getting lazy email from loader."""
        config = LazyLoadConfig(max_workers=2)
        mock_adapter = Mock()
        
        loader = LazyEmailLoader(config, mock_adapter)
        
        try:
            # Get lazy email
            lazy_email = loader.get_lazy_email("test-email-1")
            
            assert lazy_email.email_id == "test-email-1"
            assert not lazy_email.is_loaded
            
            # Getting same email should return cached instance
            lazy_email2 = loader.get_lazy_email("test-email-1")
            assert lazy_email is lazy_email2
        
        finally:
            loader.shutdown()


class TestConnectionPool:
    """Test connection pool functionality."""
    
    def test_outlook_connection_creation(self):
        """Test Outlook connection creation."""
        connection = OutlookConnection("test-conn-1")
        
        assert connection.connection_id == "test-conn-1"
        assert not connection.is_active
        assert connection.use_count == 0
    
    def test_connection_pool_initialization(self):
        """Test connection pool initialization."""
        pool = OutlookConnectionPool(
            min_connections=1,
            max_connections=3,
            max_idle_time=300
        )
        
        assert pool.min_connections == 1
        assert pool.max_connections == 3
        assert pool.max_idle_time == 300
        
        # Cleanup
        pool.shutdown()
    
    @patch('win32com.client.GetActiveObject')
    @patch('win32com.client.Dispatch')
    @patch('pythoncom.CoInitialize')
    @patch('pythoncom.CoUninitialize')
    def test_connection_pool_get_connection(self, mock_uninit, mock_init, mock_dispatch, mock_get_active):
        """Test getting connection from pool."""
        # Mock Outlook COM objects
        mock_outlook = Mock()
        mock_namespace = Mock()
        mock_inbox = Mock()
        
        mock_outlook.GetNamespace.return_value = mock_namespace
        mock_namespace.GetDefaultFolder.return_value = mock_inbox
        mock_get_active.side_effect = Exception("No active instance")
        mock_dispatch.return_value = mock_outlook
        
        pool = OutlookConnectionPool(min_connections=1, max_connections=2)
        
        try:
            # Initialize pool
            pool.initialize()
            
            # Get connection
            with pool.get_connection() as connection:
                assert connection is not None
                assert connection.is_active
        
        finally:
            pool.shutdown()


class TestPerformanceIntegration:
    """Test integration of performance components."""
    
    @pytest.mark.asyncio
    async def test_email_service_with_performance_components(self):
        """Test email service with all performance components."""
        from src.outlook_mcp_server.services.email_service import EmailService
        from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
        
        # Create mock components
        mock_adapter = Mock(spec=OutlookAdapter)
        mock_adapter.is_connected.return_value = True
        
        # Create performance components
        memory_config = MemoryConfig(max_memory_mb=64)
        rate_config = RateLimitConfig(requests_per_second=10.0)
        lazy_config = LazyLoadConfig(max_workers=1)
        
        with patch('psutil.Process'):
            memory_manager = MemoryManager(memory_config)
            rate_limiter = RateLimiter(rate_config)
            lazy_loader = LazyEmailLoader(lazy_config, mock_adapter)
            
            try:
                # Create email service with performance components
                email_service = EmailService(
                    outlook_adapter=mock_adapter,
                    memory_manager=memory_manager,
                    lazy_loader=lazy_loader,
                    rate_limiter=rate_limiter
                )
                
                # Test that service has performance components
                assert email_service.memory_manager is not None
                assert email_service.lazy_loader is not None
                assert email_service.rate_limiter is not None
                
                # Test performance stats
                stats = email_service.get_performance_stats()
                assert "memory" in stats
                assert "lazy_loader" in stats
                assert "rate_limiter" in stats
            
            finally:
                rate_limiter.shutdown()
                lazy_loader.shutdown()
                memory_manager.shutdown()
    
    def test_performance_optimization_execution(self):
        """Test performance optimization execution."""
        from src.outlook_mcp_server.services.email_service import EmailService
        from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
        
        # Create mock components
        mock_adapter = Mock(spec=OutlookAdapter)
        
        memory_config = MemoryConfig(max_memory_mb=64)
        lazy_config = LazyLoadConfig(max_workers=1)
        
        with patch('psutil.Process'):
            memory_manager = MemoryManager(memory_config)
            lazy_loader = LazyEmailLoader(lazy_config, mock_adapter)
            
            try:
                # Create email service
                email_service = EmailService(
                    outlook_adapter=mock_adapter,
                    memory_manager=memory_manager,
                    lazy_loader=lazy_loader
                )
                
                # Run performance optimization
                results = email_service.optimize_performance()
                
                assert "memory_cleanup" in results
                assert "lazy_loader_cleanup" in results
            
            finally:
                lazy_loader.shutdown()
                memory_manager.shutdown()


class TestPerformanceMetrics:
    """Test performance metrics and monitoring."""
    
    def test_memory_usage_tracking(self):
        """Test memory usage tracking."""
        config = MemoryConfig(max_memory_mb=128)
        
        with patch('psutil.Process') as mock_process:
            # Mock memory info
            mock_memory_info = Mock()
            mock_memory_info.rss = 100 * 1024 * 1024  # 100MB
            mock_memory_info.vms = 200 * 1024 * 1024  # 200MB
            
            mock_process_instance = Mock()
            mock_process_instance.memory_info.return_value = mock_memory_info
            mock_process_instance.memory_percent.return_value = 5.0
            mock_process.return_value = mock_process_instance
            
            # Mock virtual memory
            with patch('psutil.virtual_memory') as mock_vm:
                mock_vm.return_value.available = 1024 * 1024 * 1024  # 1GB
                
                manager = MemoryManager(config)
                
                usage = manager.get_memory_usage()
                
                assert usage["rss_mb"] == 100.0
                assert usage["vms_mb"] == 200.0
                assert usage["percent"] == 5.0
                assert usage["available_mb"] == 1024.0
    
    def test_rate_limiter_statistics(self):
        """Test rate limiter statistics collection."""
        config = RateLimitConfig(requests_per_second=10.0)
        rate_limiter = RateLimiter(config)
        
        try:
            stats = rate_limiter.get_stats()
            
            assert "requests_allowed" in stats
            assert "requests_denied" in stats
            assert "requests_timed_out" in stats
            assert "current_tokens" in stats
            assert "requests_last_minute" in stats
        
        finally:
            rate_limiter.shutdown()
    
    def test_lazy_loader_statistics(self):
        """Test lazy loader statistics collection."""
        config = LazyLoadConfig(max_workers=2)
        mock_adapter = Mock()
        
        loader = LazyEmailLoader(config, mock_adapter)
        
        try:
            stats = loader.get_stats()
            
            assert "cached_emails" in stats
            assert "loaded_emails" in stats
            assert "loading_emails" in stats
            assert "cache_hits" in stats
            assert "cache_misses" in stats
        
        finally:
            loader.shutdown()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])