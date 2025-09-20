"""
Health check endpoints and monitoring for Outlook MCP Server.

This module provides health check functionality for monitoring the server status,
Outlook connectivity, and system resources. It can be used by monitoring systems,
load balancers, and deployment tools to verify server health.
"""

import asyncio
import json
import time
import psutil
from datetime import datetime, timedelta
from typing import Dict, Any, Optional, List
from dataclasses import dataclass, asdict

from ..logging.logger import get_logger


@dataclass
class HealthStatus:
    """Health status data structure."""
    status: str  # "healthy", "degraded", "unhealthy"
    timestamp: str
    uptime_seconds: float
    outlook_connected: bool
    server_running: bool
    checks: Dict[str, Any]
    metrics: Dict[str, Any]


class HealthChecker:
    """Health check manager for the Outlook MCP Server."""
    
    def __init__(self, server_instance=None):
        """
        Initialize health checker.
        
        Args:
            server_instance: Reference to the main server instance
        """
        self.server = server_instance
        self.logger = get_logger(__name__)
        self.start_time = time.time()
        
        # Health check configuration
        self.check_timeout = 5.0  # seconds
        self.critical_checks = ["outlook_connection", "server_running"]
        self.warning_checks = ["memory_usage", "disk_space"]
        
        # Metrics tracking
        self.metrics_history: List[Dict[str, Any]] = []
        self.max_history_size = 100
    
    async def perform_health_check(self) -> HealthStatus:
        """
        Perform comprehensive health check.
        
        Returns:
            HealthStatus object with detailed health information
        """
        self.logger.debug("Performing health check")
        
        checks = {}
        metrics = {}
        
        try:
            # Core server checks
            checks.update(await self._check_server_status())
            checks.update(await self._check_outlook_connection())
            
            # System resource checks
            checks.update(self._check_system_resources())
            
            # Performance metrics
            metrics.update(self._collect_performance_metrics())
            
            # Determine overall health status
            overall_status = self._determine_overall_status(checks)
            
            # Create health status object
            health_status = HealthStatus(
                status=overall_status,
                timestamp=datetime.utcnow().isoformat(),
                uptime_seconds=time.time() - self.start_time,
                outlook_connected=checks.get("outlook_connection", {}).get("status") == "pass",
                server_running=checks.get("server_running", {}).get("status") == "pass",
                checks=checks,
                metrics=metrics
            )
            
            # Store metrics for trending
            self._store_metrics(metrics)
            
            self.logger.debug(f"Health check completed: {overall_status}")
            return health_status
            
        except Exception as e:
            self.logger.error(f"Health check failed: {e}", exc_info=True)
            
            # Return unhealthy status on check failure
            return HealthStatus(
                status="unhealthy",
                timestamp=datetime.utcnow().isoformat(),
                uptime_seconds=time.time() - self.start_time,
                outlook_connected=False,
                server_running=False,
                checks={"health_check_error": {"status": "fail", "message": str(e)}},
                metrics={}
            )
    
    async def _check_server_status(self) -> Dict[str, Any]:
        """Check server running status."""
        checks = {}
        
        try:
            if self.server:
                is_running = self.server.is_running()
                is_healthy = self.server.is_healthy()
                
                checks["server_running"] = {
                    "status": "pass" if is_running else "fail",
                    "message": "Server is running" if is_running else "Server is not running",
                    "details": {
                        "is_running": is_running,
                        "is_healthy": is_healthy
                    }
                }
                
                # Get server statistics
                if is_running:
                    stats = self.server.get_server_stats()
                    checks["server_stats"] = {
                        "status": "pass",
                        "message": "Server statistics available",
                        "details": stats
                    }
            else:
                checks["server_running"] = {
                    "status": "fail",
                    "message": "Server instance not available"
                }
                
        except Exception as e:
            checks["server_running"] = {
                "status": "fail",
                "message": f"Error checking server status: {e}"
            }
        
        return checks
    
    async def _check_outlook_connection(self) -> Dict[str, Any]:
        """Check Outlook connection status."""
        checks = {}
        
        try:
            if self.server and self.server.outlook_adapter:
                # Test connection with timeout
                connection_test = asyncio.create_task(
                    self._test_outlook_connection()
                )
                
                try:
                    is_connected = await asyncio.wait_for(
                        connection_test, 
                        timeout=self.check_timeout
                    )
                    
                    checks["outlook_connection"] = {
                        "status": "pass" if is_connected else "fail",
                        "message": "Outlook connection active" if is_connected else "Outlook connection failed",
                        "details": {
                            "connected": is_connected,
                            "adapter_available": True
                        }
                    }
                    
                except asyncio.TimeoutError:
                    checks["outlook_connection"] = {
                        "status": "fail",
                        "message": "Outlook connection check timed out",
                        "details": {"timeout": self.check_timeout}
                    }
                    
            else:
                checks["outlook_connection"] = {
                    "status": "fail",
                    "message": "Outlook adapter not available"
                }
                
        except Exception as e:
            checks["outlook_connection"] = {
                "status": "fail",
                "message": f"Error checking Outlook connection: {e}"
            }
        
        return checks
    
    async def _test_outlook_connection(self) -> bool:
        """Test Outlook connection asynchronously."""
        try:
            if self.server and self.server.outlook_adapter:
                return self.server.outlook_adapter.is_connected()
            return False
        except Exception:
            return False
    
    def _check_system_resources(self) -> Dict[str, Any]:
        """Check system resource usage."""
        checks = {}
        
        try:
            # Memory usage check
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            
            checks["memory_usage"] = {
                "status": "pass" if memory_percent < 85 else "warn" if memory_percent < 95 else "fail",
                "message": f"Memory usage: {memory_percent:.1f}%",
                "details": {
                    "percent": memory_percent,
                    "available_gb": memory.available / (1024**3),
                    "total_gb": memory.total / (1024**3)
                }
            }
            
            # Disk space check
            disk = psutil.disk_usage('.')
            disk_percent = (disk.used / disk.total) * 100
            
            checks["disk_space"] = {
                "status": "pass" if disk_percent < 85 else "warn" if disk_percent < 95 else "fail",
                "message": f"Disk usage: {disk_percent:.1f}%",
                "details": {
                    "percent": disk_percent,
                    "free_gb": disk.free / (1024**3),
                    "total_gb": disk.total / (1024**3)
                }
            }
            
            # CPU usage check
            cpu_percent = psutil.cpu_percent(interval=1)
            
            checks["cpu_usage"] = {
                "status": "pass" if cpu_percent < 80 else "warn" if cpu_percent < 95 else "fail",
                "message": f"CPU usage: {cpu_percent:.1f}%",
                "details": {
                    "percent": cpu_percent,
                    "count": psutil.cpu_count()
                }
            }
            
        except Exception as e:
            checks["system_resources"] = {
                "status": "fail",
                "message": f"Error checking system resources: {e}"
            }
        
        return checks
    
    def _collect_performance_metrics(self) -> Dict[str, Any]:
        """Collect performance metrics."""
        metrics = {}
        
        try:
            # Server metrics
            if self.server:
                stats = self.server.get_server_stats()
                metrics.update({
                    "requests_processed": stats.get("requests_processed", 0),
                    "requests_successful": stats.get("requests_successful", 0),
                    "requests_failed": stats.get("requests_failed", 0),
                    "success_rate": self._calculate_success_rate(stats),
                    "uptime_seconds": stats.get("uptime_seconds", 0)
                })
            
            # System metrics
            process = psutil.Process()
            metrics.update({
                "memory_usage_mb": process.memory_info().rss / (1024**2),
                "cpu_percent": process.cpu_percent(),
                "open_files": len(process.open_files()),
                "threads": process.num_threads()
            })
            
            # Add timestamp
            metrics["timestamp"] = datetime.utcnow().isoformat()
            
        except Exception as e:
            self.logger.warning(f"Error collecting performance metrics: {e}")
            metrics["collection_error"] = str(e)
        
        return metrics
    
    def _calculate_success_rate(self, stats: Dict[str, Any]) -> float:
        """Calculate request success rate."""
        total = stats.get("requests_processed", 0)
        successful = stats.get("requests_successful", 0)
        
        if total == 0:
            return 100.0  # No requests processed yet
        
        return (successful / total) * 100.0
    
    def _determine_overall_status(self, checks: Dict[str, Any]) -> str:
        """Determine overall health status based on individual checks."""
        has_critical_failure = False
        has_warning = False
        
        for check_name, check_result in checks.items():
            status = check_result.get("status", "unknown")
            
            if status == "fail":
                if check_name in self.critical_checks:
                    has_critical_failure = True
                else:
                    has_warning = True
            elif status == "warn":
                has_warning = True
        
        if has_critical_failure:
            return "unhealthy"
        elif has_warning:
            return "degraded"
        else:
            return "healthy"
    
    def _store_metrics(self, metrics: Dict[str, Any]) -> None:
        """Store metrics for historical tracking."""
        self.metrics_history.append(metrics.copy())
        
        # Limit history size
        if len(self.metrics_history) > self.max_history_size:
            self.metrics_history = self.metrics_history[-self.max_history_size:]
    
    def get_metrics_history(self, minutes: int = 60) -> List[Dict[str, Any]]:
        """
        Get metrics history for the specified time period.
        
        Args:
            minutes: Number of minutes of history to return
            
        Returns:
            List of metrics dictionaries
        """
        cutoff_time = datetime.utcnow() - timedelta(minutes=minutes)
        
        filtered_metrics = []
        for metric in self.metrics_history:
            try:
                metric_time = datetime.fromisoformat(metric.get("timestamp", ""))
                if metric_time >= cutoff_time:
                    filtered_metrics.append(metric)
            except (ValueError, TypeError):
                continue
        
        return filtered_metrics
    
    def to_dict(self, health_status: HealthStatus) -> Dict[str, Any]:
        """Convert health status to dictionary."""
        return asdict(health_status)
    
    def to_json(self, health_status: HealthStatus) -> str:
        """Convert health status to JSON string."""
        return json.dumps(self.to_dict(health_status), indent=2)


# Convenience functions for external use
async def get_health_status(server_instance=None) -> HealthStatus:
    """
    Get current health status.
    
    Args:
        server_instance: Optional server instance reference
        
    Returns:
        HealthStatus object
    """
    checker = HealthChecker(server_instance)
    return await checker.perform_health_check()


async def is_server_healthy(server_instance=None) -> bool:
    """
    Quick health check - returns True if server is healthy.
    
    Args:
        server_instance: Optional server instance reference
        
    Returns:
        True if server is healthy, False otherwise
    """
    try:
        status = await get_health_status(server_instance)
        return status.status == "healthy"
    except Exception:
        return False