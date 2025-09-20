#!/usr/bin/env python3
"""
Integration test runner for Outlook MCP Server.

This script runs comprehensive integration tests including:
- End-to-end functionality tests
- Concurrent request handling tests  
- Performance benchmarks
- Error scenario testing
- Memory and resource monitoring

Usage:
    python tests/run_integration_tests.py [options]
    
Options:
    --quick         Run only quick tests (skip performance benchmarks)
    --performance   Run only performance tests
    --concurrent    Run only concurrent tests
    --no-outlook    Skip tests requiring Outlook
    --verbose       Verbose output
    --report        Generate detailed report
"""

import argparse
import asyncio
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from integration_test_config import (
    TestConfig, TestResults, IntegrationTestRunner,
    TestDataGenerator, PerformanceMonitor, MemoryMonitor,
    validate_test_environment
)


class IntegrationTestSuite:
    """Main integration test suite runner."""
    
    def __init__(self, config: TestConfig, verbose: bool = False):
        self.config = config
        self.verbose = verbose
        self.runner = IntegrationTestRunner(config)
        self.performance_monitor = PerformanceMonitor()
        self.memory_monitor = MemoryMonitor()
        self.test_data = TestDataGenerator()
        
    async def run_all_tests(self, test_types: List[str] = None) -> Dict[str, Any]:
        """Run all integration tests."""
        if test_types is None:
            test_types = ["basic", "concurrent", "performance", "error_scenarios"]
        
        print("Starting Outlook MCP Server Integration Tests")
        print("=" * 60)
        
        # Validate environment
        if not self._validate_environment():
            return {"error": "Environment validation failed"}
        
        # Take baseline memory snapshot
        self.memory_monitor.take_snapshot("baseline")
        
        results = {}
        
        try:
            if "basic" in test_types:
                results["basic"] = await self._run_basic_tests()
            
            if "concurrent" in test_types:
                results["concurrent"] = await self._run_concurrent_tests()
            
            if "performance" in test_types:
                results["performance"] = await self._run_performance_tests()
            
            if "error_scenarios" in test_types:
                results["error_scenarios"] = await self._run_error_scenario_tests()
            
            # Final memory check
            self.memory_monitor.take_snapshot("final")
            
            # Generate summary
            results["summary"] = self._generate_summary()
            
        except Exception as e:
            results["error"] = f"Test suite failed: {str(e)}"
            if self.verbose:
                import traceback
                results["traceback"] = traceback.format_exc()
        
        finally:
            self.runner.cleanup()
        
        return results
    
    def _validate_environment(self) -> bool:
        """Validate test environment."""
        validation = validate_test_environment()
        
        if self.verbose:
            print("Environment Validation:")
            print(f"  Outlook Available: {validation['outlook_available']}")
            print(f"  Python Version: {validation['python_version']}")
            
            missing_modules = [
                module for module, status in validation["required_modules"].items()
                if status == "missing"
            ]
            
            if missing_modules:
                print(f"  Missing Modules: {missing_modules}")
                return False
        
        if self.config.outlook_required and not validation["outlook_available"]:
            print("ERROR: Outlook is required but not available")
            return False
        
        return True
    
    async def _run_basic_tests(self) -> Dict[str, Any]:
        """Run basic functionality tests."""
        print("\nRunning Basic Functionality Tests...")
        
        from src.outlook_mcp_server.server import OutlookMCPServer
        
        server_config = self.runner.create_server_config()
        server = OutlookMCPServer(server_config)
        
        results = {
            "tests_run": 0,
            "tests_passed": 0,
            "tests_failed": 0,
            "details": []
        }
        
        try:
            await server.start()
            
            # Test 1: Server startup and health
            test_result = await self._test_server_health(server)
            results["details"].append(test_result)
            results["tests_run"] += 1
            if test_result["success"]:
                results["tests_passed"] += 1
            else:
                results["tests_failed"] += 1
            
            # Test 2: Folder listing
            test_result = await self._test_folder_listing(server)
            results["details"].append(test_result)
            results["tests_run"] += 1
            if test_result["success"]:
                results["tests_passed"] += 1
            else:
                results["tests_failed"] += 1
            
            # Test 3: Email listing
            test_result = await self._test_email_listing(server)
            results["details"].append(test_result)
            results["tests_run"] += 1
            if test_result["success"]:
                results["tests_passed"] += 1
            else:
                results["tests_failed"] += 1
            
            # Test 4: Email search
            test_result = await self._test_email_search(server)
            results["details"].append(test_result)
            results["tests_run"] += 1
            if test_result["success"]:
                results["tests_passed"] += 1
            else:
                results["tests_failed"] += 1
            
        except Exception as e:
            results["error"] = str(e)
        finally:
            if server.is_running():
                await server.stop()
        
        return results
    
    async def _run_concurrent_tests(self) -> Dict[str, Any]:
        """Run concurrent request handling tests."""
        print("\nRunning Concurrent Request Tests...")
        
        from src.outlook_mcp_server.server import OutlookMCPServer
        
        server_config = self.runner.create_server_config()
        server = OutlookMCPServer(server_config)
        
        results = {
            "concurrent_clients": self.config.concurrent_clients,
            "requests_per_client": 3,
            "total_requests": 0,
            "successful_requests": 0,
            "failed_requests": 0,
            "average_response_time": 0.0,
            "max_response_time": 0.0
        }
        
        try:
            await server.start()
            
            # Generate test requests
            test_requests = self.test_data.generate_mcp_requests(
                self.config.concurrent_clients * 3
            )
            
            # Run concurrent requests
            start_time = time.time()
            
            async def make_request(request_data):
                request_start = time.time()
                try:
                    response = await server.handle_request(request_data)
                    request_end = time.time()
                    return {
                        "success": "error" not in response,
                        "response_time": request_end - request_start,
                        "response": response
                    }
                except Exception as e:
                    request_end = time.time()
                    return {
                        "success": False,
                        "response_time": request_end - request_start,
                        "error": str(e)
                    }
            
            # Execute requests concurrently
            tasks = [make_request(req) for req in test_requests]
            request_results = await asyncio.gather(*tasks, return_exceptions=True)
            
            end_time = time.time()
            
            # Analyze results
            response_times = []
            for result in request_results:
                if isinstance(result, dict):
                    results["total_requests"] += 1
                    if result["success"]:
                        results["successful_requests"] += 1
                    else:
                        results["failed_requests"] += 1
                    response_times.append(result["response_time"])
            
            if response_times:
                results["average_response_time"] = sum(response_times) / len(response_times)
                results["max_response_time"] = max(response_times)
            
            results["total_duration"] = end_time - start_time
            results["requests_per_second"] = results["total_requests"] / results["total_duration"]
            
        except Exception as e:
            results["error"] = str(e)
        finally:
            if server.is_running():
                await server.stop()
        
        return results
    
    async def _run_performance_tests(self) -> Dict[str, Any]:
        """Run performance benchmark tests."""
        print("\nRunning Performance Benchmark Tests...")
        
        from src.outlook_mcp_server.server import OutlookMCPServer
        
        server_config = self.runner.create_server_config()
        server = OutlookMCPServer(server_config)
        
        results = {
            "benchmarks": {},
            "memory_usage": {},
            "performance_summary": {}
        }
        
        try:
            await server.start()
            
            # Benchmark 1: Folder listing performance
            results["benchmarks"]["folder_listing"] = await self._benchmark_folder_listing(server)
            
            # Benchmark 2: Email listing performance
            results["benchmarks"]["email_listing"] = await self._benchmark_email_listing(server)
            
            # Benchmark 3: Search performance
            results["benchmarks"]["search_performance"] = await self._benchmark_search(server)
            
            # Memory usage analysis
            results["memory_usage"] = {
                "baseline_mb": self.memory_monitor.baseline["rss_mb"] if self.memory_monitor.baseline else 0,
                "current_mb": self.memory_monitor.snapshots[-1]["rss_mb"] if self.memory_monitor.snapshots else 0,
                "increase_mb": self.memory_monitor.get_memory_increase(),
                "within_limit": self.memory_monitor.check_memory_limit(self.config.memory_limit_mb)
            }
            
            # Performance summary
            results["performance_summary"] = self._analyze_performance_results(results["benchmarks"])
            
        except Exception as e:
            results["error"] = str(e)
        finally:
            if server.is_running():
                await server.stop()
        
        return results
    
    async def _run_error_scenario_tests(self) -> Dict[str, Any]:
        """Run error scenario tests."""
        print("\nRunning Error Scenario Tests...")
        
        from src.outlook_mcp_server.server import OutlookMCPServer
        
        server_config = self.runner.create_server_config()
        server = OutlookMCPServer(server_config)
        
        results = {
            "error_tests": [],
            "total_tests": 0,
            "handled_correctly": 0
        }
        
        try:
            await server.start()
            
            # Test invalid requests
            invalid_requests = self.test_data.generate_invalid_requests()
            
            for request in invalid_requests:
                test_result = await self._test_error_handling(server, request)
                results["error_tests"].append(test_result)
                results["total_tests"] += 1
                if test_result["handled_correctly"]:
                    results["handled_correctly"] += 1
            
            # Test connection recovery scenarios
            recovery_test = await self._test_connection_recovery(server)
            results["error_tests"].append(recovery_test)
            results["total_tests"] += 1
            if recovery_test["handled_correctly"]:
                results["handled_correctly"] += 1
            
        except Exception as e:
            results["error"] = str(e)
        finally:
            if server.is_running():
                await server.stop()
        
        return results
    
    # Individual test methods
    async def _test_server_health(self, server) -> Dict[str, Any]:
        """Test server health and basic functionality."""
        try:
            is_running = server.is_running()
            is_healthy = server.is_healthy()
            server_info = server.get_server_info()
            server_stats = server.get_server_stats()
            
            success = is_running and is_healthy and "name" in server_info
            
            return {
                "test_name": "server_health",
                "success": success,
                "details": {
                    "is_running": is_running,
                    "is_healthy": is_healthy,
                    "server_info": server_info,
                    "server_stats": server_stats
                }
            }
        except Exception as e:
            return {
                "test_name": "server_health",
                "success": False,
                "error": str(e)
            }
    
    async def _test_folder_listing(self, server) -> Dict[str, Any]:
        """Test folder listing functionality."""
        try:
            request_data = {
                "jsonrpc": "2.0",
                "id": "test_folders",
                "method": "get_folders",
                "params": {}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            success = "result" in response and "folders" in response["result"]
            folder_count = len(response["result"]["folders"]) if success else 0
            
            return {
                "test_name": "folder_listing",
                "success": success,
                "response_time": end_time - start_time,
                "details": {
                    "folder_count": folder_count,
                    "response": response
                }
            }
        except Exception as e:
            return {
                "test_name": "folder_listing",
                "success": False,
                "error": str(e)
            }
    
    async def _test_email_listing(self, server) -> Dict[str, Any]:
        """Test email listing functionality."""
        try:
            request_data = {
                "jsonrpc": "2.0",
                "id": "test_emails",
                "method": "list_emails",
                "params": {"folder": "Inbox", "limit": 5}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            success = "result" in response or ("error" in response and "not found" in response["error"]["message"].lower())
            email_count = len(response["result"]) if "result" in response else 0
            
            return {
                "test_name": "email_listing",
                "success": success,
                "response_time": end_time - start_time,
                "details": {
                    "email_count": email_count,
                    "response": response
                }
            }
        except Exception as e:
            return {
                "test_name": "email_listing",
                "success": False,
                "error": str(e)
            }
    
    async def _test_email_search(self, server) -> Dict[str, Any]:
        """Test email search functionality."""
        try:
            request_data = {
                "jsonrpc": "2.0",
                "id": "test_search",
                "method": "search_emails",
                "params": {"query": "test", "limit": 3}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            success = "result" in response or ("error" in response and response["error"]["code"] != -32603)
            result_count = len(response["result"]) if "result" in response else 0
            
            return {
                "test_name": "email_search",
                "success": success,
                "response_time": end_time - start_time,
                "details": {
                    "result_count": result_count,
                    "response": response
                }
            }
        except Exception as e:
            return {
                "test_name": "email_search",
                "success": False,
                "error": str(e)
            }
    
    async def _benchmark_folder_listing(self, server) -> Dict[str, Any]:
        """Benchmark folder listing performance."""
        iterations = 10
        times = []
        
        for i in range(iterations):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"bench_folders_{i}",
                "method": "get_folders",
                "params": {}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            times.append(end_time - start_time)
        
        return {
            "iterations": iterations,
            "times": times,
            "average": sum(times) / len(times),
            "min": min(times),
            "max": max(times),
            "total": sum(times)
        }
    
    async def _benchmark_email_listing(self, server) -> Dict[str, Any]:
        """Benchmark email listing performance."""
        iterations = 5
        times = []
        
        for i in range(iterations):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"bench_emails_{i}",
                "method": "list_emails",
                "params": {"folder": "Inbox", "limit": 10}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            times.append(end_time - start_time)
        
        return {
            "iterations": iterations,
            "times": times,
            "average": sum(times) / len(times),
            "min": min(times),
            "max": max(times),
            "total": sum(times)
        }
    
    async def _benchmark_search(self, server) -> Dict[str, Any]:
        """Benchmark search performance."""
        iterations = 3
        times = []
        search_terms = ["test", "email", "subject"]
        
        for i in range(iterations):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"bench_search_{i}",
                "method": "search_emails",
                "params": {"query": search_terms[i % len(search_terms)], "limit": 5}
            }
            
            start_time = time.time()
            response = await server.handle_request(request_data)
            end_time = time.time()
            
            times.append(end_time - start_time)
        
        return {
            "iterations": iterations,
            "times": times,
            "average": sum(times) / len(times),
            "min": min(times),
            "max": max(times),
            "total": sum(times)
        }
    
    async def _test_error_handling(self, server, invalid_request) -> Dict[str, Any]:
        """Test error handling for invalid requests."""
        try:
            response = await server.handle_request(invalid_request)
            
            # Should return error response
            handled_correctly = "error" in response and "code" in response["error"]
            
            return {
                "test_name": f"error_handling_{invalid_request.get('method', 'unknown')}",
                "request": invalid_request,
                "response": response,
                "handled_correctly": handled_correctly
            }
        except Exception as e:
            return {
                "test_name": f"error_handling_{invalid_request.get('method', 'unknown')}",
                "request": invalid_request,
                "handled_correctly": False,
                "error": str(e)
            }
    
    async def _test_connection_recovery(self, server) -> Dict[str, Any]:
        """Test connection recovery scenarios."""
        try:
            # Test server response when Outlook connection is lost
            original_health = server.is_healthy()
            
            # Simulate disconnection
            if server.outlook_adapter:
                server.outlook_adapter.disconnect()
            
            # Try to make a request
            request_data = {
                "jsonrpc": "2.0",
                "id": "recovery_test",
                "method": "get_folders",
                "params": {}
            }
            
            response = await server.handle_request(request_data)
            
            # Should handle gracefully with error response
            handled_correctly = "error" in response
            
            return {
                "test_name": "connection_recovery",
                "original_health": original_health,
                "response": response,
                "handled_correctly": handled_correctly
            }
        except Exception as e:
            return {
                "test_name": "connection_recovery",
                "handled_correctly": False,
                "error": str(e)
            }
    
    def _analyze_performance_results(self, benchmarks: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze performance benchmark results."""
        summary = {}
        
        for test_name, results in benchmarks.items():
            if "average" in results:
                threshold_key = f"{test_name}_avg"
                threshold = self.config.performance_thresholds.get(threshold_key, float('inf'))
                
                summary[test_name] = {
                    "average_time": results["average"],
                    "threshold": threshold,
                    "within_threshold": results["average"] <= threshold,
                    "performance_rating": "good" if results["average"] <= threshold else "poor"
                }
        
        return summary
    
    def _generate_summary(self) -> Dict[str, Any]:
        """Generate overall test summary."""
        return {
            "timestamp": datetime.now().isoformat(),
            "config": {
                "concurrent_clients": self.config.concurrent_clients,
                "memory_limit_mb": self.config.memory_limit_mb,
                "performance_thresholds": self.config.performance_thresholds
            },
            "performance_metrics": self.performance_monitor.get_all_stats(),
            "memory_snapshots": self.memory_monitor.snapshots,
            "memory_increase_mb": self.memory_monitor.get_memory_increase()
        }


def main():
    """Main entry point for integration test runner."""
    parser = argparse.ArgumentParser(description="Outlook MCP Server Integration Tests")
    parser.add_argument("--quick", action="store_true", help="Run only quick tests")
    parser.add_argument("--performance", action="store_true", help="Run only performance tests")
    parser.add_argument("--concurrent", action="store_true", help="Run only concurrent tests")
    parser.add_argument("--no-outlook", action="store_true", help="Skip tests requiring Outlook")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--report", action="store_true", help="Generate detailed report")
    parser.add_argument("--output", "-o", help="Output file for report")
    
    args = parser.parse_args()
    
    # Configure test types
    test_types = []
    if args.quick:
        test_types = ["basic"]
    elif args.performance:
        test_types = ["performance"]
    elif args.concurrent:
        test_types = ["concurrent"]
    else:
        test_types = ["basic", "concurrent", "performance", "error_scenarios"]
    
    # Create test configuration
    config = TestConfig(
        outlook_required=not args.no_outlook,
        concurrent_clients=5 if args.quick else 10,
        load_test_requests=5 if args.quick else 20
    )
    
    # Run tests
    test_suite = IntegrationTestSuite(config, args.verbose)
    
    try:
        results = asyncio.run(test_suite.run_all_tests(test_types))
        
        # Print summary
        print("\n" + "=" * 60)
        print("INTEGRATION TEST RESULTS")
        print("=" * 60)
        
        if "error" in results:
            print(f"ERROR: {results['error']}")
            return 1
        
        # Print results for each test type
        for test_type, test_results in results.items():
            if test_type == "summary":
                continue
                
            print(f"\n{test_type.upper()} TESTS:")
            
            if isinstance(test_results, dict):
                if "tests_run" in test_results:
                    print(f"  Tests Run: {test_results['tests_run']}")
                    print(f"  Tests Passed: {test_results['tests_passed']}")
                    print(f"  Tests Failed: {test_results['tests_failed']}")
                
                if "total_requests" in test_results:
                    print(f"  Total Requests: {test_results['total_requests']}")
                    print(f"  Successful: {test_results['successful_requests']}")
                    print(f"  Failed: {test_results['failed_requests']}")
                    print(f"  Avg Response Time: {test_results.get('average_response_time', 0):.3f}s")
                
                if "benchmarks" in test_results:
                    print("  Performance Benchmarks:")
                    for bench_name, bench_results in test_results["benchmarks"].items():
                        if "average" in bench_results:
                            print(f"    {bench_name}: {bench_results['average']:.3f}s avg")
        
        # Generate report if requested
        if args.report:
            output_file = args.output or f"integration_test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            with open(output_file, 'w') as f:
                json.dump(results, f, indent=2, default=str)
            print(f"\nDetailed report saved to: {output_file}")
        
        # Determine exit code
        overall_success = True
        for test_type, test_results in results.items():
            if isinstance(test_results, dict) and "tests_failed" in test_results:
                if test_results["tests_failed"] > 0:
                    overall_success = False
                    break
        
        return 0 if overall_success else 1
        
    except KeyboardInterrupt:
        print("\nTests interrupted by user")
        return 1
    except Exception as e:
        print(f"\nTest suite failed: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())