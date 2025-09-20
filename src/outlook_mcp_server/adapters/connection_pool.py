"""Connection pool for managing Outlook COM connections."""

import logging
import threading
import time
from typing import Optional, List, Dict, Any
from queue import Queue, Empty, Full
from contextlib import contextmanager
import win32com.client
import pythoncom
from ..models.exceptions import OutlookConnectionError


logger = logging.getLogger(__name__)


class OutlookConnection:
    """Represents a single Outlook COM connection."""
    
    def __init__(self, connection_id: str):
        """Initialize Outlook connection."""
        self.connection_id = connection_id
        self.outlook_app: Optional[Any] = None
        self.namespace: Optional[Any] = None
        self.created_at = time.time()
        self.last_used = time.time()
        self.use_count = 0
        self.is_active = False
        self._lock = threading.Lock()
    
    def connect(self) -> bool:
        """Establish COM connection to Outlook."""
        with self._lock:
            try:
                logger.debug(f"Creating Outlook connection {self.connection_id}")
                
                # Initialize COM for this thread
                pythoncom.CoInitialize()
                
                # Try to get existing Outlook instance first
                try:
                    self.outlook_app = win32com.client.GetActiveObject("Outlook.Application")
                    logger.debug(f"Connected to existing Outlook instance: {self.connection_id}")
                except:
                    # If no existing instance, create new one
                    self.outlook_app = win32com.client.Dispatch("Outlook.Application")
                    logger.debug(f"Created new Outlook instance: {self.connection_id}")
                
                # Get the MAPI namespace
                self.namespace = self.outlook_app.GetNamespace("MAPI")
                
                # Test the connection
                self._test_connection()
                
                self.is_active = True
                self.last_used = time.time()
                
                logger.info(f"Outlook connection {self.connection_id} established successfully")
                return True
                
            except Exception as e:
                logger.error(f"Failed to create Outlook connection {self.connection_id}: {str(e)}")
                self._cleanup()
                raise OutlookConnectionError(f"Failed to create connection: {str(e)}")
    
    def _test_connection(self) -> None:
        """Test the connection by accessing basic functionality."""
        try:
            if not self.namespace:
                raise OutlookConnectionError("Namespace not available")
            
            # Try to access the default inbox folder
            inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if not inbox:
                raise OutlookConnectionError("Cannot access default inbox folder")
            
            logger.debug(f"Connection test successful for {self.connection_id}")
            
        except Exception as e:
            logger.error(f"Connection test failed for {self.connection_id}: {str(e)}")
            raise OutlookConnectionError(f"Connection test failed: {str(e)}")
    
    def is_healthy(self) -> bool:
        """Check if connection is healthy and usable."""
        with self._lock:
            if not self.is_active or not self.outlook_app or not self.namespace:
                return False
            
            try:
                # Test connection by accessing a basic property
                self.namespace.GetDefaultFolder(6)  # Try to access inbox
                return True
            except:
                self.is_active = False
                return False
    
    def mark_used(self) -> None:
        """Mark connection as recently used."""
        with self._lock:
            self.last_used = time.time()
            self.use_count += 1
    
    def get_age(self) -> float:
        """Get connection age in seconds."""
        return time.time() - self.created_at
    
    def get_idle_time(self) -> float:
        """Get time since last use in seconds."""
        return time.time() - self.last_used
    
    def disconnect(self) -> None:
        """Disconnect and cleanup resources."""
        with self._lock:
            logger.debug(f"Disconnecting Outlook connection {self.connection_id}")
            self._cleanup()
    
    def _cleanup(self) -> None:
        """Clean up COM objects and connection state."""
        try:
            self.namespace = None
            self.outlook_app = None
            self.is_active = False
            
            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass  # Ignore errors during COM cleanup
                
        except Exception as e:
            logger.error(f"Error during connection cleanup {self.connection_id}: {str(e)}")


class OutlookConnectionPool:
    """Pool for managing multiple Outlook COM connections."""
    
    def __init__(self, 
                 min_connections: int = 1,
                 max_connections: int = 5,
                 max_idle_time: int = 300,  # 5 minutes
                 max_connection_age: int = 3600,  # 1 hour
                 health_check_interval: int = 60):  # 1 minute
        """
        Initialize connection pool.
        
        Args:
            min_connections: Minimum number of connections to maintain
            max_connections: Maximum number of connections allowed
            max_idle_time: Maximum idle time before connection is closed (seconds)
            max_connection_age: Maximum age of connection before renewal (seconds)
            health_check_interval: Interval for health checks (seconds)
        """
        self.min_connections = min_connections
        self.max_connections = max_connections
        self.max_idle_time = max_idle_time
        self.max_connection_age = max_connection_age
        self.health_check_interval = health_check_interval
        
        self._pool: Queue[OutlookConnection] = Queue(maxsize=max_connections)
        self._all_connections: Dict[str, OutlookConnection] = {}
        self._connection_counter = 0
        self._lock = threading.RLock()
        self._shutdown = False
        
        # Statistics
        self._stats = {
            "connections_created": 0,
            "connections_destroyed": 0,
            "connections_borrowed": 0,
            "connections_returned": 0,
            "pool_hits": 0,
            "pool_misses": 0
        }
        
        # Start background maintenance thread
        self._maintenance_thread = threading.Thread(
            target=self._maintenance_worker,
            daemon=True,
            name="outlook-pool-maintenance"
        )
        self._maintenance_thread.start()
        
        logger.info(f"Outlook connection pool initialized: min={min_connections}, max={max_connections}")
    
    def initialize(self) -> None:
        """Initialize the pool with minimum connections."""
        with self._lock:
            logger.info("Initializing connection pool with minimum connections")
            
            for _ in range(self.min_connections):
                try:
                    connection = self._create_connection()
                    self._pool.put_nowait(connection)
                except Exception as e:
                    logger.error(f"Failed to create initial connection: {str(e)}")
                    # Continue trying to create other connections
    
    @contextmanager
    def get_connection(self, timeout: float = 10.0):
        """
        Get a connection from the pool using context manager.
        
        Args:
            timeout: Timeout for getting connection from pool
            
        Yields:
            OutlookConnection: Active Outlook connection
            
        Raises:
            OutlookConnectionError: If no connection available within timeout
        """
        connection = None
        try:
            connection = self._borrow_connection(timeout)
            yield connection
        finally:
            if connection:
                self._return_connection(connection)
    
    def _borrow_connection(self, timeout: float) -> OutlookConnection:
        """Borrow a connection from the pool."""
        with self._lock:
            if self._shutdown:
                raise OutlookConnectionError("Connection pool is shutdown")
            
            start_time = time.time()
            
            while time.time() - start_time < timeout:
                try:
                    # Try to get connection from pool
                    connection = self._pool.get_nowait()
                    
                    # Check if connection is healthy
                    if connection.is_healthy():
                        connection.mark_used()
                        self._stats["connections_borrowed"] += 1
                        self._stats["pool_hits"] += 1
                        logger.debug(f"Borrowed healthy connection {connection.connection_id}")
                        return connection
                    else:
                        # Connection is unhealthy, destroy it
                        logger.warning(f"Removing unhealthy connection {connection.connection_id}")
                        self._destroy_connection(connection)
                        
                except Empty:
                    # Pool is empty, try to create new connection
                    if len(self._all_connections) < self.max_connections:
                        try:
                            connection = self._create_connection()
                            connection.mark_used()
                            self._stats["connections_borrowed"] += 1
                            self._stats["pool_misses"] += 1
                            logger.debug(f"Created new connection {connection.connection_id}")
                            return connection
                        except Exception as e:
                            logger.error(f"Failed to create new connection: {str(e)}")
                
                # Wait a bit before retrying
                time.sleep(0.1)
            
            raise OutlookConnectionError(f"No connection available within {timeout} seconds")
    
    def _return_connection(self, connection: OutlookConnection) -> None:
        """Return a connection to the pool."""
        with self._lock:
            if self._shutdown:
                self._destroy_connection(connection)
                return
            
            # Check if connection is still healthy and not too old
            if (connection.is_healthy() and 
                connection.get_age() < self.max_connection_age):
                
                try:
                    self._pool.put_nowait(connection)
                    self._stats["connections_returned"] += 1
                    logger.debug(f"Returned connection {connection.connection_id} to pool")
                except Full:
                    # Pool is full, destroy the connection
                    logger.debug(f"Pool full, destroying connection {connection.connection_id}")
                    self._destroy_connection(connection)
            else:
                # Connection is unhealthy or too old
                logger.debug(f"Destroying aged/unhealthy connection {connection.connection_id}")
                self._destroy_connection(connection)
    
    def _create_connection(self) -> OutlookConnection:
        """Create a new Outlook connection."""
        with self._lock:
            self._connection_counter += 1
            connection_id = f"outlook-conn-{self._connection_counter}"
            
            connection = OutlookConnection(connection_id)
            connection.connect()
            
            self._all_connections[connection_id] = connection
            self._stats["connections_created"] += 1
            
            logger.debug(f"Created new Outlook connection: {connection_id}")
            return connection
    
    def _destroy_connection(self, connection: OutlookConnection) -> None:
        """Destroy a connection and remove from tracking."""
        with self._lock:
            try:
                connection.disconnect()
                
                if connection.connection_id in self._all_connections:
                    del self._all_connections[connection.connection_id]
                
                self._stats["connections_destroyed"] += 1
                logger.debug(f"Destroyed connection: {connection.connection_id}")
                
            except Exception as e:
                logger.error(f"Error destroying connection {connection.connection_id}: {str(e)}")
    
    def _maintenance_worker(self) -> None:
        """Background worker for pool maintenance."""
        logger.debug("Starting connection pool maintenance worker")
        
        while not self._shutdown:
            try:
                time.sleep(self.health_check_interval)
                
                if self._shutdown:
                    break
                
                self._perform_maintenance()
                
            except Exception as e:
                logger.error(f"Error in pool maintenance: {str(e)}")
    
    def _perform_maintenance(self) -> None:
        """Perform pool maintenance tasks."""
        with self._lock:
            logger.debug("Performing connection pool maintenance")
            
            # Remove idle and aged connections
            connections_to_remove = []
            
            for connection in list(self._all_connections.values()):
                if (connection.get_idle_time() > self.max_idle_time or
                    connection.get_age() > self.max_connection_age or
                    not connection.is_healthy()):
                    
                    connections_to_remove.append(connection)
            
            # Remove connections from pool queue
            temp_connections = []
            while not self._pool.empty():
                try:
                    conn = self._pool.get_nowait()
                    if conn not in connections_to_remove:
                        temp_connections.append(conn)
                except Empty:
                    break
            
            # Put back good connections
            for conn in temp_connections:
                try:
                    self._pool.put_nowait(conn)
                except Full:
                    break
            
            # Destroy bad connections
            for connection in connections_to_remove:
                self._destroy_connection(connection)
            
            # Ensure minimum connections
            current_pool_size = self._pool.qsize()
            if current_pool_size < self.min_connections:
                needed = self.min_connections - current_pool_size
                for _ in range(needed):
                    if len(self._all_connections) < self.max_connections:
                        try:
                            connection = self._create_connection()
                            self._pool.put_nowait(connection)
                        except Exception as e:
                            logger.error(f"Failed to create maintenance connection: {str(e)}")
                            break
            
            logger.debug(f"Pool maintenance complete. Active connections: {len(self._all_connections)}, "
                        f"Pool size: {self._pool.qsize()}")
    
    def get_stats(self) -> Dict[str, Any]:
        """Get connection pool statistics."""
        with self._lock:
            stats = self._stats.copy()
            stats.update({
                "active_connections": len(self._all_connections),
                "pool_size": self._pool.qsize(),
                "max_connections": self.max_connections,
                "min_connections": self.min_connections
            })
            return stats
    
    def shutdown(self) -> None:
        """Shutdown the connection pool."""
        logger.info("Shutting down connection pool")
        
        with self._lock:
            self._shutdown = True
            
            # Destroy all connections
            while not self._pool.empty():
                try:
                    connection = self._pool.get_nowait()
                    self._destroy_connection(connection)
                except Empty:
                    break
            
            # Destroy any remaining tracked connections
            for connection in list(self._all_connections.values()):
                self._destroy_connection(connection)
            
            self._all_connections.clear()
        
        # Wait for maintenance thread to finish
        if self._maintenance_thread.is_alive():
            self._maintenance_thread.join(timeout=5.0)
        
        logger.info("Connection pool shutdown complete")