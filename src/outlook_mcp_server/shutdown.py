"""
Graceful shutdown handling for Outlook MCP Server.

This module provides comprehensive shutdown handling to ensure proper cleanup
of resources, completion of in-flight requests, and safe termination of the
server process.
"""

import asyncio
import signal
import threading
import time
from typing import Dict, Any, Optional, Callable, List
from dataclasses import dataclass
from enum import Enum

from .logging.logger import get_logger


class ShutdownPhase(Enum):
    """Shutdown phases for coordinated shutdown process."""
    RUNNING = "running"
    SHUTDOWN_REQUESTED = "shutdown_requested"
    STOPPING_NEW_REQUESTS = "stopping_new_requests"
    DRAINING_REQUESTS = "draining_requests"
    CLEANING_UP = "cleaning_up"
    STOPPED = "stopped"


@dataclass
class ShutdownStats:
    """Statistics about the shutdown process."""
    shutdown_requested_at: float
    shutdown_completed_at: Optional[float] = None
    requests_in_flight_at_shutdown: int = 0
    requests_completed_during_shutdown: int = 0
    cleanup_tasks_completed: int = 0
    total_shutdown_time: Optional[float] = None


class GracefulShutdownHandler:
    """
    Handles graceful shutdown of the Outlook MCP Server.
    
    This class coordinates the shutdown process to ensure:
    - No new requests are accepted after shutdown is initiated
    - In-flight requests are allowed to complete (with timeout)
    - Resources are properly cleaned up
    - Shutdown process is logged and monitored
    """
    
    def __init__(self, 
                 shutdown_timeout: int = 30,
                 drain_timeout: int = 15,
                 cleanup_timeout: int = 10):
        """
        Initialize graceful shutdown handler.
        
        Args:
            shutdown_timeout: Total time allowed for shutdown process
            drain_timeout: Time to wait for in-flight requests to complete
            cleanup_timeout: Time allowed for cleanup operations
        """
        self.shutdown_timeout = shutdown_timeout
        self.drain_timeout = drain_timeout
        self.cleanup_timeout = cleanup_timeout
        
        self.logger = get_logger(__name__)
        
        # Shutdown state
        self.phase = ShutdownPhase.RUNNING
        self.shutdown_event = threading.Event()
        self.shutdown_requested = False
        
        # Request tracking
        self.active_requests: Dict[str, float] = {}  # request_id -> start_time
        self.request_lock = threading.Lock()
        
        # Cleanup callbacks
        self.cleanup_callbacks: List[Callable[[], None]] = []
        self.async_cleanup_callbacks: List[Callable[[], asyncio.Task]] = []
        
        # Statistics
        self.stats = ShutdownStats(shutdown_requested_at=0)
        
        # Signal handlers
        self._original_handlers: Dict[int, Any] = {}
        self._setup_signal_handlers()
    
    def _setup_signal_handlers(self) -> None:
        """Setup signal handlers for graceful shutdown."""
        signals_to_handle = [signal.SIGINT, signal.SIGTERM]
        
        # Add Windows-specific signals
        if hasattr(signal, 'SIGBREAK'):
            signals_to_handle.append(signal.SIGBREAK)
        
        for sig in signals_to_handle:
            try:
                # Store original handler
                self._original_handlers[sig] = signal.signal(sig, self._signal_handler)
                self.logger.debug(f"Registered signal handler for {sig}")
            except (OSError, ValueError) as e:
                self.logger.warning(f"Could not register handler for signal {sig}: {e}")
    
    def _signal_handler(self, signum: int, frame) -> None:
        """Handle shutdown signals."""
        self.logger.info(f"Received signal {signum}, initiating graceful shutdown")
        self.initiate_shutdown()
    
    def initiate_shutdown(self) -> None:
        """Initiate the graceful shutdown process."""
        if self.shutdown_requested:
            self.logger.warning("Shutdown already requested, ignoring duplicate request")
            return
        
        self.logger.info("Initiating graceful shutdown")
        
        self.shutdown_requested = True
        self.phase = ShutdownPhase.SHUTDOWN_REQUESTED
        self.stats.shutdown_requested_at = time.time()
        
        # Count active requests at shutdown
        with self.request_lock:
            self.stats.requests_in_flight_at_shutdown = len(self.active_requests)
        
        # Signal shutdown event
        self.shutdown_event.set()
        
        self.logger.info(f"Shutdown initiated with {self.stats.requests_in_flight_at_shutdown} active requests")
    
    async def shutdown(self) -> ShutdownStats:
        """
        Execute the graceful shutdown process.
        
        Returns:
            ShutdownStats with information about the shutdown process
        """
        if not self.shutdown_requested:
            self.initiate_shutdown()
        
        shutdown_start = time.time()
        
        try:
            # Phase 1: Stop accepting new requests
            await self._stop_new_requests()
            
            # Phase 2: Drain existing requests
            await self._drain_requests()
            
            # Phase 3: Cleanup resources
            await self._cleanup_resources()
            
            # Phase 4: Complete shutdown
            self.phase = ShutdownPhase.STOPPED
            self.stats.shutdown_completed_at = time.time()
            self.stats.total_shutdown_time = self.stats.shutdown_completed_at - self.stats.shutdown_requested_at
            
            self.logger.info(f"Graceful shutdown completed in {self.stats.total_shutdown_time:.2f} seconds")
            
        except asyncio.TimeoutError:
            self.logger.error(f"Shutdown timed out after {self.shutdown_timeout} seconds")
            self.phase = ShutdownPhase.STOPPED
            self.stats.shutdown_completed_at = time.time()
            self.stats.total_shutdown_time = self.stats.shutdown_completed_at - self.stats.shutdown_requested_at
        
        except Exception as e:
            self.logger.error(f"Error during shutdown: {e}", exc_info=True)
            self.phase = ShutdownPhase.STOPPED
            self.stats.shutdown_completed_at = time.time()
            self.stats.total_shutdown_time = self.stats.shutdown_completed_at - self.stats.shutdown_requested_at
        
        finally:
            # Restore original signal handlers
            self._restore_signal_handlers()
        
        return self.stats
    
    async def _stop_new_requests(self) -> None:
        """Stop accepting new requests."""
        self.phase = ShutdownPhase.STOPPING_NEW_REQUESTS
        self.logger.info("Stopping acceptance of new requests")
        
        # This phase is mainly informational - the server should check
        # self.is_shutdown_requested() before processing new requests
        
        await asyncio.sleep(0.1)  # Brief pause to ensure state is propagated
    
    async def _drain_requests(self) -> None:
        """Wait for in-flight requests to complete."""
        self.phase = ShutdownPhase.DRAINING_REQUESTS
        self.logger.info(f"Draining in-flight requests (timeout: {self.drain_timeout}s)")
        
        drain_start = time.time()
        
        while time.time() - drain_start < self.drain_timeout:
            with self.request_lock:
                active_count = len(self.active_requests)
            
            if active_count == 0:
                self.logger.info("All requests completed successfully")
                break
            
            self.logger.debug(f"Waiting for {active_count} requests to complete")
            await asyncio.sleep(0.5)
        
        # Log final status
        with self.request_lock:
            remaining_requests = len(self.active_requests)
        
        if remaining_requests > 0:
            self.logger.warning(f"Shutdown proceeding with {remaining_requests} requests still active")
        else:
            self.logger.info("Request draining completed successfully")
    
    async def _cleanup_resources(self) -> None:
        """Execute cleanup callbacks."""
        self.phase = ShutdownPhase.CLEANING_UP
        self.logger.info("Executing cleanup operations")
        
        cleanup_start = time.time()
        
        try:
            # Execute synchronous cleanup callbacks
            for callback in self.cleanup_callbacks:
                try:
                    callback()
                    self.stats.cleanup_tasks_completed += 1
                except Exception as e:
                    self.logger.error(f"Error in cleanup callback: {e}", exc_info=True)
            
            # Execute asynchronous cleanup callbacks
            cleanup_tasks = []
            for callback in self.async_cleanup_callbacks:
                try:
                    task = callback()
                    if asyncio.iscoroutine(task):
                        cleanup_tasks.append(task)
                    else:
                        self.logger.warning("Async cleanup callback did not return coroutine")
                except Exception as e:
                    self.logger.error(f"Error creating async cleanup task: {e}", exc_info=True)
            
            # Wait for async cleanup tasks with timeout
            if cleanup_tasks:
                try:
                    await asyncio.wait_for(
                        asyncio.gather(*cleanup_tasks, return_exceptions=True),
                        timeout=self.cleanup_timeout
                    )
                    self.stats.cleanup_tasks_completed += len(cleanup_tasks)
                except asyncio.TimeoutError:
                    self.logger.warning(f"Some cleanup tasks timed out after {self.cleanup_timeout}s")
            
            cleanup_duration = time.time() - cleanup_start
            self.logger.info(f"Cleanup completed in {cleanup_duration:.2f}s ({self.stats.cleanup_tasks_completed} tasks)")
            
        except Exception as e:
            self.logger.error(f"Error during cleanup: {e}", exc_info=True)
    
    def _restore_signal_handlers(self) -> None:
        """Restore original signal handlers."""
        for sig, handler in self._original_handlers.items():
            try:
                signal.signal(sig, handler)
            except (OSError, ValueError) as e:
                self.logger.warning(f"Could not restore handler for signal {sig}: {e}")
    
    def register_request(self, request_id: str) -> None:
        """
        Register a new request for tracking.
        
        Args:
            request_id: Unique identifier for the request
        """
        with self.request_lock:
            self.active_requests[request_id] = time.time()
    
    def unregister_request(self, request_id: str) -> None:
        """
        Unregister a completed request.
        
        Args:
            request_id: Unique identifier for the request
        """
        with self.request_lock:
            if request_id in self.active_requests:
                del self.active_requests[request_id]
                if self.shutdown_requested:
                    self.stats.requests_completed_during_shutdown += 1
    
    def register_cleanup_callback(self, callback: Callable[[], None]) -> None:
        """
        Register a synchronous cleanup callback.
        
        Args:
            callback: Function to call during cleanup
        """
        self.cleanup_callbacks.append(callback)
    
    def register_async_cleanup_callback(self, callback: Callable[[], asyncio.Task]) -> None:
        """
        Register an asynchronous cleanup callback.
        
        Args:
            callback: Function that returns a coroutine for cleanup
        """
        self.async_cleanup_callbacks.append(callback)
    
    def is_shutdown_requested(self) -> bool:
        """Check if shutdown has been requested."""
        return self.shutdown_requested
    
    def should_accept_requests(self) -> bool:
        """Check if new requests should be accepted."""
        return self.phase == ShutdownPhase.RUNNING
    
    def get_active_request_count(self) -> int:
        """Get the number of currently active requests."""
        with self.request_lock:
            return len(self.active_requests)
    
    def get_shutdown_stats(self) -> ShutdownStats:
        """Get current shutdown statistics."""
        return self.stats
    
    def wait_for_shutdown(self, timeout: Optional[float] = None) -> bool:
        """
        Wait for shutdown to be requested.
        
        Args:
            timeout: Maximum time to wait (None for indefinite)
            
        Returns:
            True if shutdown was requested, False if timeout occurred
        """
        return self.shutdown_event.wait(timeout)


# Global shutdown handler instance
_shutdown_handler: Optional[GracefulShutdownHandler] = None


def get_shutdown_handler() -> GracefulShutdownHandler:
    """Get the global shutdown handler instance."""
    global _shutdown_handler
    if _shutdown_handler is None:
        _shutdown_handler = GracefulShutdownHandler()
    return _shutdown_handler


def setup_graceful_shutdown(shutdown_timeout: int = 30,
                          drain_timeout: int = 15,
                          cleanup_timeout: int = 10) -> GracefulShutdownHandler:
    """
    Setup graceful shutdown handling.
    
    Args:
        shutdown_timeout: Total time allowed for shutdown process
        drain_timeout: Time to wait for in-flight requests to complete
        cleanup_timeout: Time allowed for cleanup operations
        
    Returns:
        GracefulShutdownHandler instance
    """
    global _shutdown_handler
    _shutdown_handler = GracefulShutdownHandler(
        shutdown_timeout=shutdown_timeout,
        drain_timeout=drain_timeout,
        cleanup_timeout=cleanup_timeout
    )
    return _shutdown_handler