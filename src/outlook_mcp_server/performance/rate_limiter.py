"""Rate limiting and timeout handling for MCP requests."""

import asyncio
import time
import threading
from typing import Dict, Any, Optional, Callable, Awaitable
from collections import defaultdict, deque
from dataclasses import dataclass
import logging
from ..models.exceptions import ValidationError, OutlookMCPError


logger = logging.getLogger(__name__)


@dataclass
class RateLimitConfig:
    """Configuration for rate limiting."""
    requests_per_second: float = 10.0
    requests_per_minute: int = 300
    requests_per_hour: int = 1000
    burst_size: int = 20
    timeout_seconds: float = 30.0
    cleanup_interval: int = 60


class TokenBucket:
    """Token bucket implementation for rate limiting."""
    
    def __init__(self, capacity: int, refill_rate: float):
        """
        Initialize token bucket.
        
        Args:
            capacity: Maximum number of tokens
            refill_rate: Tokens added per second
        """
        self.capacity = capacity
        self.refill_rate = refill_rate
        self.tokens = float(capacity)
        self.last_refill = time.time()
        self._lock = threading.Lock()
    
    def consume(self, tokens: int = 1) -> bool:
        """
        Try to consume tokens from bucket.
        
        Args:
            tokens: Number of tokens to consume
            
        Returns:
            bool: True if tokens were consumed, False if not enough tokens
        """
        with self._lock:
            now = time.time()
            
            # Add tokens based on elapsed time
            elapsed = now - self.last_refill
            self.tokens = min(self.capacity, self.tokens + elapsed * self.refill_rate)
            self.last_refill = now
            
            # Check if we have enough tokens
            if self.tokens >= tokens:
                self.tokens -= tokens
                return True
            
            return False
    
    def get_wait_time(self, tokens: int = 1) -> float:
        """
        Get time to wait until tokens are available.
        
        Args:
            tokens: Number of tokens needed
            
        Returns:
            float: Seconds to wait
        """
        with self._lock:
            if self.tokens >= tokens:
                return 0.0
            
            needed_tokens = tokens - self.tokens
            return needed_tokens / self.refill_rate


class RequestTracker:
    """Track request counts for different time windows."""
    
    def __init__(self):
        """Initialize request tracker."""
        self.requests_per_minute: deque = deque()
        self.requests_per_hour: deque = deque()
        self._lock = threading.Lock()
    
    def add_request(self, timestamp: float = None) -> None:
        """Add a request timestamp."""
        if timestamp is None:
            timestamp = time.time()
        
        with self._lock:
            self.requests_per_minute.append(timestamp)
            self.requests_per_hour.append(timestamp)
            
            # Clean old entries
            self._cleanup_old_requests(timestamp)
    
    def get_request_count(self, window_seconds: int) -> int:
        """Get request count for time window."""
        with self._lock:
            now = time.time()
            cutoff = now - window_seconds
            
            if window_seconds <= 60:
                # Use minute queue
                return sum(1 for ts in self.requests_per_minute if ts > cutoff)
            else:
                # Use hour queue
                return sum(1 for ts in self.requests_per_hour if ts > cutoff)
    
    def _cleanup_old_requests(self, current_time: float) -> None:
        """Remove old request timestamps."""
        minute_cutoff = current_time - 60
        hour_cutoff = current_time - 3600
        
        # Clean minute queue
        while self.requests_per_minute and self.requests_per_minute[0] < minute_cutoff:
            self.requests_per_minute.popleft()
        
        # Clean hour queue
        while self.requests_per_hour and self.requests_per_hour[0] < hour_cutoff:
            self.requests_per_hour.popleft()


class RateLimiter:
    """Rate limiter with multiple strategies and time windows."""
    
    def __init__(self, config: RateLimitConfig):
        """
        Initialize rate limiter.
        
        Args:
            config: Rate limiting configuration
        """
        self.config = config
        
        # Token bucket for burst control
        self.token_bucket = TokenBucket(
            capacity=config.burst_size,
            refill_rate=config.requests_per_second
        )
        
        # Request trackers per client/method
        self.client_trackers: Dict[str, RequestTracker] = defaultdict(RequestTracker)
        self.method_trackers: Dict[str, RequestTracker] = defaultdict(RequestTracker)
        self.global_tracker = RequestTracker()
        
        # Statistics
        self._stats = {
            "requests_allowed": 0,
            "requests_denied": 0,
            "requests_timed_out": 0,
            "total_wait_time": 0.0
        }
        
        self._lock = threading.RLock()
        
        # Start cleanup thread
        self._shutdown = False
        self._cleanup_thread = threading.Thread(
            target=self._cleanup_worker,
            daemon=True,
            name="rate-limiter-cleanup"
        )
        self._cleanup_thread.start()
        
        logger.info(f"Rate limiter initialized: {config}")
    
    async def acquire(self, 
                     client_id: str = "default",
                     method: str = "unknown",
                     timeout: Optional[float] = None) -> bool:
        """
        Acquire permission to make a request.
        
        Args:
            client_id: Identifier for the client
            method: Method being called
            timeout: Optional timeout override
            
        Returns:
            bool: True if request is allowed
            
        Raises:
            ValidationError: If rate limit exceeded
        """
        if timeout is None:
            timeout = self.config.timeout_seconds
        
        start_time = time.time()
        
        try:
            # Check rate limits
            await self._check_rate_limits(client_id, method, timeout)
            
            # Record successful acquisition
            self._record_request(client_id, method)
            
            with self._lock:
                self._stats["requests_allowed"] += 1
            
            logger.debug(f"Rate limit acquired for {client_id}:{method}")
            return True
            
        except asyncio.TimeoutError:
            with self._lock:
                self._stats["requests_timed_out"] += 1
            
            logger.warning(f"Rate limit timeout for {client_id}:{method}")
            raise ValidationError(
                f"Request timed out waiting for rate limit: {timeout}s",
                "rate_limit"
            )
        
        except Exception as e:
            with self._lock:
                self._stats["requests_denied"] += 1
            
            logger.warning(f"Rate limit denied for {client_id}:{method}: {str(e)}")
            raise
    
    async def _check_rate_limits(self, 
                                client_id: str, 
                                method: str, 
                                timeout: float) -> None:
        """Check all rate limit conditions."""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            # Check token bucket (burst control)
            if not self.token_bucket.consume():
                wait_time = self.token_bucket.get_wait_time()
                if wait_time > 0:
                    await asyncio.sleep(min(wait_time, 0.1))
                    continue
            
            # Check per-second rate
            if not self._check_requests_per_second():
                await asyncio.sleep(0.1)
                continue
            
            # Check per-minute rate (global)
            if self.global_tracker.get_request_count(60) >= self.config.requests_per_minute:
                wait_time = self._calculate_wait_time(60, self.config.requests_per_minute)
                if wait_time > 0:
                    await asyncio.sleep(min(wait_time, 1.0))
                    continue
            
            # Check per-hour rate (global)
            if self.global_tracker.get_request_count(3600) >= self.config.requests_per_hour:
                wait_time = self._calculate_wait_time(3600, self.config.requests_per_hour)
                if wait_time > 0:
                    raise ValidationError(
                        f"Hourly rate limit exceeded: {self.config.requests_per_hour}/hour",
                        "rate_limit"
                    )
            
            # Check client-specific limits (more lenient)
            client_per_minute = self.client_trackers[client_id].get_request_count(60)
            if client_per_minute >= self.config.requests_per_minute // 2:  # Half of global limit per client
                wait_time = self._calculate_wait_time(60, self.config.requests_per_minute // 2)
                if wait_time > 0:
                    await asyncio.sleep(min(wait_time, 1.0))
                    continue
            
            # All checks passed
            return
        
        # Timeout exceeded
        raise asyncio.TimeoutError()
    
    def _check_requests_per_second(self) -> bool:
        """Check if we're within per-second rate limit."""
        now = time.time()
        recent_requests = self.global_tracker.get_request_count(1)
        return recent_requests < self.config.requests_per_second
    
    def _calculate_wait_time(self, window_seconds: int, limit: int) -> float:
        """Calculate time to wait for rate limit window."""
        now = time.time()
        cutoff = now - window_seconds
        
        # Find oldest request in window
        if window_seconds <= 60:
            requests = [ts for ts in self.global_tracker.requests_per_minute if ts > cutoff]
        else:
            requests = [ts for ts in self.global_tracker.requests_per_hour if ts > cutoff]
        
        if len(requests) < limit:
            return 0.0
        
        # Wait until oldest request falls out of window
        oldest_request = min(requests)
        return (oldest_request + window_seconds) - now
    
    def _record_request(self, client_id: str, method: str) -> None:
        """Record a successful request."""
        now = time.time()
        
        with self._lock:
            self.global_tracker.add_request(now)
            self.client_trackers[client_id].add_request(now)
            self.method_trackers[method].add_request(now)
    
    def _cleanup_worker(self) -> None:
        """Background worker for cleaning up old data."""
        logger.debug("Starting rate limiter cleanup worker")
        
        while not self._shutdown:
            try:
                time.sleep(self.config.cleanup_interval)
                
                if self._shutdown:
                    break
                
                self._cleanup_old_data()
                
            except Exception as e:
                logger.error(f"Error in rate limiter cleanup: {str(e)}")
    
    def _cleanup_old_data(self) -> None:
        """Clean up old tracking data."""
        with self._lock:
            now = time.time()
            
            # Clean up client trackers that haven't been used recently
            inactive_clients = []
            for client_id, tracker in self.client_trackers.items():
                if not tracker.requests_per_minute or tracker.requests_per_minute[-1] < now - 3600:
                    inactive_clients.append(client_id)
            
            for client_id in inactive_clients:
                del self.client_trackers[client_id]
            
            # Clean up method trackers
            inactive_methods = []
            for method, tracker in self.method_trackers.items():
                if not tracker.requests_per_minute or tracker.requests_per_minute[-1] < now - 3600:
                    inactive_methods.append(method)
            
            for method in inactive_methods:
                del self.method_trackers[method]
            
            logger.debug(f"Rate limiter cleanup: removed {len(inactive_clients)} clients, "
                        f"{len(inactive_methods)} methods")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get rate limiter statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update({
                "active_clients": len(self.client_trackers),
                "active_methods": len(self.method_trackers),
                "current_tokens": self.token_bucket.tokens,
                "requests_last_minute": self.global_tracker.get_request_count(60),
                "requests_last_hour": self.global_tracker.get_request_count(3600)
            })
            return stats
    
    def shutdown(self) -> None:
        """Shutdown the rate limiter."""
        logger.info("Shutting down rate limiter")
        
        self._shutdown = True
        
        if self._cleanup_thread.is_alive():
            self._cleanup_thread.join(timeout=5.0)
        
        logger.info("Rate limiter shutdown complete")


class TimeoutManager:
    """Manages request timeouts and cancellation."""
    
    def __init__(self, default_timeout: float = 30.0):
        """
        Initialize timeout manager.
        
        Args:
            default_timeout: Default timeout in seconds
        """
        self.default_timeout = default_timeout
        self._active_requests: Dict[str, asyncio.Task] = {}
        self._lock = threading.RLock()
        
        # Statistics
        self._stats = {
            "requests_started": 0,
            "requests_completed": 0,
            "requests_timed_out": 0,
            "requests_cancelled": 0
        }
        
        logger.info(f"Timeout manager initialized with default timeout: {default_timeout}s")
    
    async def execute_with_timeout(self,
                                  coro: Awaitable[Any],
                                  timeout: Optional[float] = None,
                                  request_id: str = None) -> Any:
        """
        Execute coroutine with timeout.
        
        Args:
            coro: Coroutine to execute
            timeout: Timeout in seconds
            request_id: Optional request identifier
            
        Returns:
            Result of coroutine execution
            
        Raises:
            asyncio.TimeoutError: If execution times out
        """
        if timeout is None:
            timeout = self.default_timeout
        
        if request_id is None:
            request_id = f"req-{int(time.time() * 1000)}"
        
        with self._lock:
            self._stats["requests_started"] += 1
        
        try:
            # Create task and track it
            task = asyncio.create_task(coro)
            
            with self._lock:
                self._active_requests[request_id] = task
            
            # Execute with timeout
            result = await asyncio.wait_for(task, timeout=timeout)
            
            with self._lock:
                self._stats["requests_completed"] += 1
            
            logger.debug(f"Request {request_id} completed successfully")
            return result
            
        except asyncio.TimeoutError:
            with self._lock:
                self._stats["requests_timed_out"] += 1
            
            logger.warning(f"Request {request_id} timed out after {timeout}s")
            
            # Cancel the task
            await self._cancel_request(request_id)
            raise
            
        except asyncio.CancelledError:
            with self._lock:
                self._stats["requests_cancelled"] += 1
            
            logger.info(f"Request {request_id} was cancelled")
            raise
            
        finally:
            # Clean up tracking
            with self._lock:
                self._active_requests.pop(request_id, None)
    
    async def _cancel_request(self, request_id: str) -> None:
        """Cancel a specific request."""
        with self._lock:
            task = self._active_requests.get(request_id)
        
        if task and not task.done():
            task.cancel()
            try:
                await task
            except asyncio.CancelledError:
                pass
            
            logger.debug(f"Cancelled request {request_id}")
    
    async def cancel_all_requests(self) -> None:
        """Cancel all active requests."""
        with self._lock:
            active_requests = list(self._active_requests.items())
        
        logger.info(f"Cancelling {len(active_requests)} active requests")
        
        for request_id, task in active_requests:
            if not task.done():
                await self._cancel_request(request_id)
    
    def get_stats(self) -> Dict[str, Any]:
        """Get timeout manager statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update({
                "active_requests": len(self._active_requests),
                "default_timeout": self.default_timeout
            })
            return stats