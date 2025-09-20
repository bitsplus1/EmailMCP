"""
Configuration and utilities for integration testing.
"""

import os
import tempfile
from typing import Dict, Any, List
from dataclasses import dataclass
from datetime import datetime


@dataclass
class TestConfig:
    """Configuration for integration tests."""
    outlook_required: bool = True
    max_test_duration: int = 300  # 5 minutes
    performance_thresholds: Dict[str, float] = None
    concurrent_clients: int = 5
    load_test_requests: int = 10
    memory_limit_mb: int = 500
    
    def __post_init__(self):
        if self.performance_thresholds is None:
            self.performance_thresholds = {
                "folder_listing_avg": 2.0,  # seconds
                "email_listing_avg": 5.0,   # seconds
                "email_retrieval_avg": 3.0, # seconds
                "search_avg": 10.0,         # seconds
                "concurrent_max": 15.0      # seconds
            }


@dataclass
class TestResults:
    """Results from integration test runs."""
    test_name: str
    start_time: datetime
    end_time: datetime
    success: bool
    error_message: str = ""
    performance_metrics: Dict[str, float] = None
    
    def __post_init__(self):
        if self.performance_metrics is None:
            self.performance_metrics = {}
    
    @property
    def duration(self) -> float:
        """Test duration in seconds."""
        return (self.end_time - self.start_time).total_seconds()


class IntegrationTestRunner:
    """Runner for integration tests with configuration and reporting."""
    
    def __init__(self, config: TestConfig = None):
        self.config = config or TestConfig()
        self.results: List[TestResults] = []
        self.temp_dir = tempfile.mkdtemp(prefix="outlook_mcp_integration_")
    
    def create_server_config(self) -> Dict[str, Any]:
        """Create server configuration for testing."""
        return {
            "log_level": "WARNING",  # Reduce noise during testing
            "log_dir": self.temp_dir,
            "max_concurrent_requests": self.config.concurrent_clients * 2,
            "request_timeout": 30,
            "outlook_connection_timeout": 10,
            "enable_performance_logging": True,
            "enable_console_output": False
        }
    
    def check_prerequisites(self) -> bool:
        """Check if prerequisites for integration testing are met."""
        if not self.config.outlook_required:
            return True
        
        try:
            from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
            adapter = OutlookAdapter()
            result = adapter.connect()
            if result:
                adapter.disconnect()
                return True
            return False
        except Exception:
            return False
    
    def record_result(self, result: TestResults):
        """Record a test result."""
        self.results.append(result)
    
    def get_summary(self) -> Dict[str, Any]:
        """Get summary of test results."""
        total_tests = len(self.results)
        successful_tests = sum(1 for r in self.results if r.success)
        failed_tests = total_tests - successful_tests
        
        total_duration = sum(r.duration for r in self.results)
        avg_duration = total_duration / total_tests if total_tests > 0 else 0
        
        return {
            "total_tests": total_tests,
            "successful_tests": successful_tests,
            "failed_tests": failed_tests,
            "success_rate": successful_tests / total_tests if total_tests > 0 else 0,
            "total_duration": total_duration,
            "average_duration": avg_duration,
            "performance_metrics": self._aggregate_performance_metrics()
        }
    
    def _aggregate_performance_metrics(self) -> Dict[str, float]:
        """Aggregate performance metrics from all tests."""
        all_metrics = {}
        
        for result in self.results:
            for metric, value in result.performance_metrics.items():
                if metric not in all_metrics:
                    all_metrics[metric] = []
                all_metrics[metric].append(value)
        
        # Calculate averages
        aggregated = {}
        for metric, values in all_metrics.items():
            if values:
                aggregated[f"{metric}_avg"] = sum(values) / len(values)
                aggregated[f"{metric}_min"] = min(values)
                aggregated[f"{metric}_max"] = max(values)
        
        return aggregated
    
    def cleanup(self):
        """Cleanup test resources."""
        import shutil
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        except Exception:
            pass


# Test data generators for integration testing
class TestDataGenerator:
    """Generate test data for integration tests."""
    
    @staticmethod
    def generate_mcp_requests(num_requests: int = 10) -> List[Dict[str, Any]]:
        """Generate a variety of MCP requests for testing."""
        requests = []
        
        # Folder requests
        for i in range(num_requests // 4):
            requests.append({
                "jsonrpc": "2.0",
                "id": f"folders_{i}",
                "method": "get_folders",
                "params": {}
            })
        
        # Email listing requests
        folders = ["Inbox", "Sent Items", "Drafts"]
        for i in range(num_requests // 4):
            requests.append({
                "jsonrpc": "2.0",
                "id": f"list_emails_{i}",
                "method": "list_emails",
                "params": {
                    "folder": folders[i % len(folders)],
                    "limit": 5 + (i % 10),
                    "unread_only": i % 2 == 0
                }
            })
        
        # Search requests
        search_terms = ["test", "email", "subject", "important", "meeting"]
        for i in range(num_requests // 4):
            requests.append({
                "jsonrpc": "2.0",
                "id": f"search_{i}",
                "method": "search_emails",
                "params": {
                    "query": search_terms[i % len(search_terms)],
                    "limit": 3 + (i % 7)
                }
            })
        
        # Fill remaining with mixed requests
        remaining = num_requests - len(requests)
        for i in range(remaining):
            if i % 2 == 0:
                requests.append({
                    "jsonrpc": "2.0",
                    "id": f"mixed_folders_{i}",
                    "method": "get_folders",
                    "params": {}
                })
            else:
                requests.append({
                    "jsonrpc": "2.0",
                    "id": f"mixed_list_{i}",
                    "method": "list_emails",
                    "params": {"folder": "Inbox", "limit": 3}
                })
        
        return requests
    
    @staticmethod
    def generate_invalid_requests() -> List[Dict[str, Any]]:
        """Generate invalid requests for error testing."""
        return [
            # Invalid method
            {
                "jsonrpc": "2.0",
                "id": "invalid_method",
                "method": "nonexistent_method",
                "params": {}
            },
            # Missing parameters
            {
                "jsonrpc": "2.0",
                "id": "missing_params",
                "method": "get_email",
                "params": {}
            },
            # Invalid parameter types
            {
                "jsonrpc": "2.0",
                "id": "invalid_param_type",
                "method": "list_emails",
                "params": {"limit": "not_a_number"}
            },
            # Invalid folder
            {
                "jsonrpc": "2.0",
                "id": "invalid_folder",
                "method": "list_emails",
                "params": {"folder": "NonExistentFolder123"}
            },
            # Invalid email ID
            {
                "jsonrpc": "2.0",
                "id": "invalid_email",
                "method": "get_email",
                "params": {"email_id": "invalid_id_format"}
            }
        ]


# Performance monitoring utilities
class PerformanceMonitor:
    """Monitor performance during integration tests."""
    
    def __init__(self):
        self.metrics = {}
        self.start_times = {}
    
    def start_timer(self, operation: str):
        """Start timing an operation."""
        import time
        self.start_times[operation] = time.time()
    
    def end_timer(self, operation: str) -> float:
        """End timing an operation and return duration."""
        import time
        if operation in self.start_times:
            duration = time.time() - self.start_times[operation]
            if operation not in self.metrics:
                self.metrics[operation] = []
            self.metrics[operation].append(duration)
            del self.start_times[operation]
            return duration
        return 0.0
    
    def get_stats(self, operation: str) -> Dict[str, float]:
        """Get statistics for an operation."""
        if operation not in self.metrics or not self.metrics[operation]:
            return {}
        
        times = self.metrics[operation]
        return {
            "count": len(times),
            "total": sum(times),
            "average": sum(times) / len(times),
            "min": min(times),
            "max": max(times)
        }
    
    def get_all_stats(self) -> Dict[str, Dict[str, float]]:
        """Get statistics for all operations."""
        return {op: self.get_stats(op) for op in self.metrics.keys()}


# Memory monitoring utilities
class MemoryMonitor:
    """Monitor memory usage during integration tests."""
    
    def __init__(self):
        self.snapshots = []
        self.baseline = None
    
    def take_snapshot(self, label: str = ""):
        """Take a memory usage snapshot."""
        try:
            import psutil
            process = psutil.Process()
            memory_info = process.memory_info()
            
            snapshot = {
                "label": label,
                "timestamp": datetime.now(),
                "rss_mb": memory_info.rss / 1024 / 1024,
                "vms_mb": memory_info.vms / 1024 / 1024
            }
            
            self.snapshots.append(snapshot)
            
            if self.baseline is None:
                self.baseline = snapshot
            
            return snapshot
        except ImportError:
            return {"error": "psutil not available"}
    
    def get_memory_increase(self) -> float:
        """Get memory increase since baseline in MB."""
        if not self.snapshots or self.baseline is None:
            return 0.0
        
        current = self.snapshots[-1]
        return current["rss_mb"] - self.baseline["rss_mb"]
    
    def check_memory_limit(self, limit_mb: float) -> bool:
        """Check if memory usage is within limit."""
        if not self.snapshots:
            return True
        
        current = self.snapshots[-1]
        return current["rss_mb"] <= limit_mb


# Environment validation
def validate_test_environment() -> Dict[str, Any]:
    """Validate the test environment and return status."""
    validation = {
        "outlook_available": False,
        "python_version": "",
        "required_modules": {},
        "system_info": {},
        "recommendations": []
    }
    
    # Check Python version
    import sys
    validation["python_version"] = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    
    # Check required modules
    required_modules = [
        "win32com.client",
        "pythoncom", 
        "pytest",
        "asyncio",
        "psutil"
    ]
    
    for module in required_modules:
        try:
            __import__(module)
            validation["required_modules"][module] = "available"
        except ImportError:
            validation["required_modules"][module] = "missing"
            validation["recommendations"].append(f"Install {module}")
    
    # Check Outlook availability
    try:
        from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
        adapter = OutlookAdapter()
        if adapter.connect():
            validation["outlook_available"] = True
            adapter.disconnect()
        else:
            validation["recommendations"].append("Start Microsoft Outlook")
    except Exception as e:
        validation["recommendations"].append(f"Fix Outlook connection: {str(e)}")
    
    # System information
    try:
        import platform
        validation["system_info"] = {
            "platform": platform.platform(),
            "processor": platform.processor(),
            "python_implementation": platform.python_implementation()
        }
    except Exception:
        pass
    
    return validation


if __name__ == "__main__":
    # Quick environment validation
    validation = validate_test_environment()
    
    print("Integration Test Environment Validation")
    print("=" * 50)
    print(f"Python Version: {validation['python_version']}")
    print(f"Outlook Available: {validation['outlook_available']}")
    
    print("\nRequired Modules:")
    for module, status in validation["required_modules"].items():
        print(f"  {module}: {status}")
    
    if validation["recommendations"]:
        print("\nRecommendations:")
        for rec in validation["recommendations"]:
            print(f"  - {rec}")
    
    print("\nSystem Info:")
    for key, value in validation["system_info"].items():
        print(f"  {key}: {value}")