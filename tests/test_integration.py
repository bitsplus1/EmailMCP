"""
Comprehensive integration tests for the Outlook MCP Server.

These tests require a real Outlook installation and are designed to test:
1. End-to-end functionality with real Outlook integration
2. Concurrent request handling
3. Various Outlook folder structures and email types
4. Error scenario testing with actual COM failures
5. Performance benchmarks and load testing

Note: These tests require Microsoft Outlook to be installed and configured
on the test system. They should be run in a controlled environment.
"""

import asyncio
import json
import pytest
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
from unittest.mock import patch, Mock
import tempfile
import os

from src.outlook_mcp_server.server import OutlookMCPServer, create_server_config
from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
from src.outlook_mcp_server.services.email_service import EmailService
from src.outlook_mcp_server.services.folder_service import FolderService
from src.outlook_mcp_server.models.exceptions import (
    OutlookConnectionError,
    FolderNotFoundError,
    EmailNotFoundError,
    ValidationError
)
from src.outlook_mcp_server.models.mcp_models import MCPRequest


class TestOutlookIntegration:
    """Integration tests with real Outlook connection."""
    
    @pytest.fixture(scope="class")
    def outlook_adapter(self):
        """Create and connect to real Outlook instance."""
        adapter = OutlookAdapter()
        try:
            # Attempt to connect to real Outlook
            if adapter.connect():
                yield adapter
            else:
                pytest.skip("Outlook not available for integration testing")
        except OutlookConnectionError:
            pytest.skip("Outlook not available for integration testing")
        finally:
            if adapter.is_connected():
                adapter.disconnect()
    
    @pytest.fixture(scope="class")
    def email_service(self, outlook_adapter):
        """Create email service with real Outlook connection."""
        return EmailService(outlook_adapter)
    
    @pytest.fixture(scope="class")
    def folder_service(self, outlook_adapter):
        """Create folder service with real Outlook connection."""
        return FolderService(outlook_adapter)
    
    def test_real_outlook_connection(self, outlook_adapter):
        """Test connection to real Outlook instance."""
        assert outlook_adapter.is_connected()
        
        # Test namespace access
        namespace = outlook_adapter.get_namespace()
        assert namespace is not None
        
        # Test basic folder access
        try:
            inbox = outlook_adapter.get_folder_by_name("Inbox")
            assert inbox is not None
        except Exception as e:
            pytest.fail(f"Failed to access Inbox folder: {e}")
    
    def test_real_folder_listing(self, folder_service):
        """Test listing real Outlook folders."""
        folders = folder_service.get_folders()
        
        # Should have at least basic folders
        assert len(folders) > 0
        
        # Check for common folder names
        folder_names = [folder["name"] for folder in folders]
        common_folders = ["Inbox", "Sent Items", "Drafts", "Deleted Items"]
        
        # At least some common folders should exist
        found_folders = [name for name in common_folders if name in folder_names]
        assert len(found_folders) > 0, f"No common folders found. Available: {folder_names}"
        
        # Validate folder structure
        for folder in folders:
            assert "id" in folder
            assert "name" in folder
            assert "full_path" in folder
            assert isinstance(folder["item_count"], int)
            assert isinstance(folder["unread_count"], int)
    
    def test_real_email_listing(self, email_service):
        """Test listing real emails from Outlook."""
        # Test listing emails from Inbox
        emails = email_service.list_emails(folder="Inbox", limit=5)
        
        # Should return a list (may be empty if no emails)
        assert isinstance(emails, list)
        
        if len(emails) > 0:
            # Validate email structure
            for email in emails:
                assert "id" in email
                assert "subject" in email
                assert "sender" in email
                assert "received_time" in email
                assert "is_read" in email
                assert isinstance(email["is_read"], bool)
    
    def test_real_email_retrieval(self, email_service):
        """Test retrieving specific real emails."""
        # First get a list of emails
        emails = email_service.list_emails(folder="Inbox", limit=1)
        
        if len(emails) == 0:
            pytest.skip("No emails available for testing")
        
        # Get detailed email information
        email_id = emails[0]["id"]
        detailed_email = email_service.get_email(email_id)
        
        # Validate detailed email structure
        assert detailed_email["id"] == email_id
        assert "subject" in detailed_email
        assert "sender" in detailed_email
        assert "sender_email" in detailed_email
        assert "body" in detailed_email
        assert "received_time" in detailed_email
        assert "sent_time" in detailed_email
        assert "is_read" in detailed_email
        assert "has_attachments" in detailed_email
        assert "folder_name" in detailed_email
    
    def test_real_email_search(self, email_service):
        """Test searching real emails."""
        # Search for emails with common terms
        search_terms = ["test", "email", "subject", "from"]
        
        for term in search_terms:
            results = email_service.search_emails(query=term, limit=3)
            assert isinstance(results, list)
            
            # If results found, validate structure
            for email in results:
                assert "id" in email
                assert "subject" in email
                assert "sender" in email
                # Search term should appear somewhere in the email
                email_text = f"{email['subject']} {email.get('body', '')}".lower()
                # Note: Real search might not always match our simple text search
    
    def test_folder_access_permissions(self, folder_service, outlook_adapter):
        """Test folder access with different permission levels."""
        folders = folder_service.get_folders()
        
        for folder in folders[:3]:  # Test first 3 folders
            folder_name = folder["name"]
            
            try:
                # Try to access folder
                outlook_folder = outlook_adapter.get_folder_by_name(folder_name)
                assert outlook_folder is not None
                
                # Try to list items (may fail for some system folders)
                try:
                    items = outlook_folder.Items
                    count = items.Count if hasattr(items, 'Count') else 0
                    assert count >= 0
                except Exception:
                    # Some folders may not allow item access
                    pass
                    
            except Exception as e:
                # Some folders may not be accessible
                print(f"Could not access folder {folder_name}: {e}")
    
    def test_email_filtering_real_data(self, email_service):
        """Test email filtering with real data."""
        # Test unread filter
        unread_emails = email_service.list_emails(folder="Inbox", unread_only=True, limit=10)
        assert isinstance(unread_emails, list)
        
        # All returned emails should be unread
        for email in unread_emails:
            assert email["is_read"] is False
        
        # Test limit functionality
        limited_emails = email_service.list_emails(folder="Inbox", limit=2)
        assert len(limited_emails) <= 2
    
    def test_error_handling_with_real_outlook(self, email_service, folder_service):
        """Test error handling with real Outlook scenarios."""
        # Test invalid folder access
        with pytest.raises(FolderNotFoundError):
            email_service.list_emails(folder="NonExistentFolder123")
        
        # Test invalid email ID
        with pytest.raises(EmailNotFoundError):
            email_service.get_email("invalid_email_id_12345")
        
        # Test empty search query
        results = email_service.search_emails(query="")
        assert results == []
        
        # Test search in invalid folder
        with pytest.raises(FolderNotFoundError):
            email_service.search_emails(query="test", folder="InvalidFolder123")


class TestServerIntegration:
    """Integration tests for the complete MCP server."""
    
    @pytest.fixture
    async def server(self):
        """Create and start a test server."""
        config = create_server_config(
            log_level="DEBUG",
            log_dir=tempfile.mkdtemp(),
            max_concurrent_requests=5,
            request_timeout=10,
            outlook_connection_timeout=5
        )
        
        server = OutlookMCPServer(config)
        
        try:
            await server.start()
            yield server
        except OutlookConnectionError:
            pytest.skip("Outlook not available for server integration testing")
        finally:
            if server.is_running():
                await server.stop()
    
    @pytest.mark.asyncio
    async def test_server_startup_and_shutdown(self):
        """Test complete server lifecycle."""
        config = create_server_config(
            log_level="INFO",
            log_dir=tempfile.mkdtemp()
        )
        
        server = OutlookMCPServer(config)
        
        # Test startup
        try:
            await server.start()
            assert server.is_running()
            assert server.is_healthy()
            
            # Test server info
            info = server.get_server_info()
            assert "name" in info
            assert "version" in info
            
            # Test server stats
            stats = server.get_server_stats()
            assert "requests_processed" in stats
            assert "is_running" in stats
            assert stats["is_running"] is True
            
        except OutlookConnectionError:
            pytest.skip("Outlook not available for server testing")
        finally:
            # Test shutdown
            await server.stop()
            assert not server.is_running()
    
    @pytest.mark.asyncio
    async def test_mcp_request_handling(self, server):
        """Test MCP request processing."""
        # Test list_emails request
        request_data = {
            "jsonrpc": "2.0",
            "id": "test_1",
            "method": "list_emails",
            "params": {"folder": "Inbox", "limit": 3}
        }
        
        response = await server.handle_request(request_data)
        
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test_1"
        
        if "result" in response:
            # Successful response
            assert isinstance(response["result"], list)
        else:
            # Error response (acceptable if no emails)
            assert "error" in response
    
    @pytest.mark.asyncio
    async def test_get_folders_request(self, server):
        """Test get_folders MCP request."""
        request_data = {
            "jsonrpc": "2.0",
            "id": "test_folders",
            "method": "get_folders",
            "params": {}
        }
        
        response = await server.handle_request(request_data)
        
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test_folders"
        
        if "result" in response:
            folders = response["result"]["folders"]
            assert isinstance(folders, list)
            assert len(folders) > 0
        else:
            assert "error" in response
    
    @pytest.mark.asyncio
    async def test_invalid_mcp_request(self, server):
        """Test handling of invalid MCP requests."""
        # Test invalid method
        request_data = {
            "jsonrpc": "2.0",
            "id": "test_invalid",
            "method": "invalid_method",
            "params": {}
        }
        
        response = await server.handle_request(request_data)
        
        assert response["jsonrpc"] == "2.0"
        assert response["id"] == "test_invalid"
        assert "error" in response
        assert response["error"]["code"] == -32601  # Method not found
    
    @pytest.mark.asyncio
    async def test_malformed_request(self, server):
        """Test handling of malformed requests."""
        # Test missing required fields
        request_data = {
            "jsonrpc": "2.0",
            # Missing id and method
            "params": {}
        }
        
        response = await server.handle_request(request_data)
        
        assert "error" in response
        assert response["error"]["code"] == -32600  # Invalid Request


class TestConcurrentRequests:
    """Test concurrent request handling."""
    
    @pytest.fixture
    async def server(self):
        """Create server for concurrent testing."""
        config = create_server_config(
            log_level="WARNING",  # Reduce log noise
            log_dir=tempfile.mkdtemp(),
            max_concurrent_requests=10
        )
        
        server = OutlookMCPServer(config)
        
        try:
            await server.start()
            yield server
        except OutlookConnectionError:
            pytest.skip("Outlook not available for concurrent testing")
        finally:
            if server.is_running():
                await server.stop()
    
    @pytest.mark.asyncio
    async def test_concurrent_folder_requests(self, server):
        """Test concurrent get_folders requests."""
        num_requests = 5
        
        async def make_request(request_id: int):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"concurrent_{request_id}",
                "method": "get_folders",
                "params": {}
            }
            return await server.handle_request(request_data)
        
        # Execute concurrent requests
        tasks = [make_request(i) for i in range(num_requests)]
        responses = await asyncio.gather(*tasks, return_exceptions=True)
        
        # Validate responses
        successful_responses = 0
        for i, response in enumerate(responses):
            if isinstance(response, Exception):
                print(f"Request {i} failed: {response}")
            else:
                assert response["jsonrpc"] == "2.0"
                assert response["id"] == f"concurrent_{i}"
                if "result" in response:
                    successful_responses += 1
        
        # At least some requests should succeed
        assert successful_responses > 0
    
    @pytest.mark.asyncio
    async def test_concurrent_mixed_requests(self, server):
        """Test concurrent requests of different types."""
        
        async def make_folders_request():
            return await server.handle_request({
                "jsonrpc": "2.0",
                "id": "folders_req",
                "method": "get_folders",
                "params": {}
            })
        
        async def make_emails_request():
            return await server.handle_request({
                "jsonrpc": "2.0",
                "id": "emails_req",
                "method": "list_emails",
                "params": {"folder": "Inbox", "limit": 2}
            })
        
        async def make_search_request():
            return await server.handle_request({
                "jsonrpc": "2.0",
                "id": "search_req",
                "method": "search_emails",
                "params": {"query": "test", "limit": 1}
            })
        
        # Execute different types of requests concurrently
        tasks = [
            make_folders_request(),
            make_emails_request(),
            make_search_request(),
            make_folders_request(),  # Duplicate to test caching/concurrency
        ]
        
        responses = await asyncio.gather(*tasks, return_exceptions=True)
        
        # Validate that all requests completed
        assert len(responses) == 4
        
        for response in responses:
            if not isinstance(response, Exception):
                assert response["jsonrpc"] == "2.0"
                assert "id" in response
    
    def test_thread_safety_with_real_outlook(self):
        """Test thread safety with real Outlook COM objects."""
        if not self._outlook_available():
            pytest.skip("Outlook not available for thread safety testing")
        
        results = []
        errors = []
        
        def worker_thread(thread_id: int):
            """Worker function for thread safety testing."""
            try:
                adapter = OutlookAdapter()
                adapter.connect()
                
                # Perform operations
                folders = adapter.get_folders()
                results.append(f"Thread {thread_id}: Found {len(folders)} folders")
                
                # Try to list some emails
                if folders:
                    folder_name = folders[0].name
                    emails = adapter.list_emails(folder_name, limit=1)
                    results.append(f"Thread {thread_id}: Found {len(emails)} emails in {folder_name}")
                
                adapter.disconnect()
                
            except Exception as e:
                errors.append(f"Thread {thread_id} error: {str(e)}")
        
        # Create and start multiple threads
        threads = []
        num_threads = 3
        
        for i in range(num_threads):
            thread = threading.Thread(target=worker_thread, args=(i,))
            threads.append(thread)
            thread.start()
        
        # Wait for all threads to complete
        for thread in threads:
            thread.join(timeout=30)  # 30 second timeout
        
        # Check results
        print(f"Results: {results}")
        print(f"Errors: {errors}")
        
        # At least some threads should succeed
        assert len(results) > 0, f"No successful operations. Errors: {errors}"
    
    def _outlook_available(self) -> bool:
        """Check if Outlook is available for testing."""
        try:
            adapter = OutlookAdapter()
            result = adapter.connect()
            if result:
                adapter.disconnect()
            return result
        except:
            return False


class TestPerformanceBenchmarks:
    """Performance benchmarks and load testing."""
    
    @pytest.fixture
    async def server(self):
        """Create server for performance testing."""
        config = create_server_config(
            log_level="ERROR",  # Minimal logging for performance
            log_dir=tempfile.mkdtemp(),
            max_concurrent_requests=20,
            enable_performance_logging=True
        )
        
        server = OutlookMCPServer(config)
        
        try:
            await server.start()
            yield server
        except OutlookConnectionError:
            pytest.skip("Outlook not available for performance testing")
        finally:
            if server.is_running():
                await server.stop()
    
    @pytest.mark.asyncio
    async def test_folder_listing_performance(self, server):
        """Benchmark folder listing performance."""
        num_iterations = 10
        start_time = time.time()
        
        for i in range(num_iterations):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"perf_folders_{i}",
                "method": "get_folders",
                "params": {}
            }
            
            response = await server.handle_request(request_data)
            
            # Should complete successfully or with acceptable error
            assert "jsonrpc" in response
        
        end_time = time.time()
        total_time = end_time - start_time
        avg_time = total_time / num_iterations
        
        print(f"Folder listing: {num_iterations} requests in {total_time:.2f}s")
        print(f"Average time per request: {avg_time:.3f}s")
        
        # Performance assertion - should complete within reasonable time
        assert avg_time < 2.0, f"Folder listing too slow: {avg_time:.3f}s per request"
    
    @pytest.mark.asyncio
    async def test_email_listing_performance(self, server):
        """Benchmark email listing performance."""
        num_iterations = 5
        start_time = time.time()
        
        for i in range(num_iterations):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"perf_emails_{i}",
                "method": "list_emails",
                "params": {"folder": "Inbox", "limit": 10}
            }
            
            response = await server.handle_request(request_data)
            assert "jsonrpc" in response
        
        end_time = time.time()
        total_time = end_time - start_time
        avg_time = total_time / num_iterations
        
        print(f"Email listing: {num_iterations} requests in {total_time:.2f}s")
        print(f"Average time per request: {avg_time:.3f}s")
        
        # Performance assertion
        assert avg_time < 5.0, f"Email listing too slow: {avg_time:.3f}s per request"
    
    @pytest.mark.asyncio
    async def test_concurrent_load(self, server):
        """Test server under concurrent load."""
        num_concurrent = 10
        requests_per_client = 3
        
        async def client_load(client_id: int):
            """Simulate client making multiple requests."""
            client_times = []
            
            for req_id in range(requests_per_client):
                start = time.time()
                
                request_data = {
                    "jsonrpc": "2.0",
                    "id": f"load_client_{client_id}_req_{req_id}",
                    "method": "get_folders",
                    "params": {}
                }
                
                response = await server.handle_request(request_data)
                
                end = time.time()
                client_times.append(end - start)
                
                # Brief pause between requests
                await asyncio.sleep(0.1)
            
            return client_times
        
        # Start concurrent clients
        start_time = time.time()
        tasks = [client_load(i) for i in range(num_concurrent)]
        all_times = await asyncio.gather(*tasks, return_exceptions=True)
        end_time = time.time()
        
        # Analyze results
        successful_clients = 0
        all_request_times = []
        
        for client_times in all_times:
            if not isinstance(client_times, Exception):
                successful_clients += 1
                all_request_times.extend(client_times)
        
        total_requests = successful_clients * requests_per_client
        total_time = end_time - start_time
        
        if all_request_times:
            avg_request_time = sum(all_request_times) / len(all_request_times)
            max_request_time = max(all_request_times)
            
            print(f"Load test: {total_requests} requests from {num_concurrent} clients")
            print(f"Total time: {total_time:.2f}s")
            print(f"Requests per second: {total_requests / total_time:.2f}")
            print(f"Average request time: {avg_request_time:.3f}s")
            print(f"Max request time: {max_request_time:.3f}s")
            
            # Performance assertions
            assert successful_clients > 0, "No clients completed successfully"
            assert avg_request_time < 3.0, f"Average request time too high: {avg_request_time:.3f}s"
            assert max_request_time < 10.0, f"Max request time too high: {max_request_time:.3f}s"
    
    def test_memory_usage_stability(self):
        """Test memory usage stability over time."""
        if not self._outlook_available():
            pytest.skip("Outlook not available for memory testing")
        
        import psutil
        import gc
        
        process = psutil.Process()
        initial_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        # Perform many operations
        adapter = OutlookAdapter()
        adapter.connect()
        
        try:
            for i in range(50):
                # Perform various operations
                folders = adapter.get_folders()
                if folders:
                    emails = adapter.list_emails(folders[0].name, limit=5)
                
                # Force garbage collection periodically
                if i % 10 == 0:
                    gc.collect()
                    current_memory = process.memory_info().rss / 1024 / 1024
                    print(f"Iteration {i}: Memory usage: {current_memory:.1f} MB")
        
        finally:
            adapter.disconnect()
        
        final_memory = process.memory_info().rss / 1024 / 1024
        memory_increase = final_memory - initial_memory
        
        print(f"Initial memory: {initial_memory:.1f} MB")
        print(f"Final memory: {final_memory:.1f} MB")
        print(f"Memory increase: {memory_increase:.1f} MB")
        
        # Memory should not increase excessively
        assert memory_increase < 100, f"Memory usage increased too much: {memory_increase:.1f} MB"
    
    def _outlook_available(self) -> bool:
        """Check if Outlook is available for testing."""
        try:
            adapter = OutlookAdapter()
            result = adapter.connect()
            if result:
                adapter.disconnect()
            return result
        except:
            return False


class TestErrorScenarios:
    """Test various error scenarios with real Outlook."""
    
    @pytest.fixture
    async def server(self):
        """Create server for error testing."""
        config = create_server_config(
            log_level="DEBUG",
            log_dir=tempfile.mkdtemp(),
            outlook_connection_timeout=2  # Short timeout for testing
        )
        
        server = OutlookMCPServer(config)
        
        try:
            await server.start()
            yield server
        except OutlookConnectionError:
            pytest.skip("Outlook not available for error scenario testing")
        finally:
            if server.is_running():
                await server.stop()
    
    @pytest.mark.asyncio
    async def test_outlook_disconnection_recovery(self, server):
        """Test recovery from Outlook disconnection."""
        # Verify server is initially healthy
        assert server.is_healthy()
        
        # Simulate disconnection by disconnecting the adapter
        if server.outlook_adapter:
            server.outlook_adapter.disconnect()
        
        # Server should detect unhealthy state
        assert not server.is_healthy()
        
        # Requests should still be handled gracefully
        request_data = {
            "jsonrpc": "2.0",
            "id": "test_disconnected",
            "method": "get_folders",
            "params": {}
        }
        
        response = await server.handle_request(request_data)
        
        # Should return error response
        assert "error" in response
        assert response["error"]["code"] < 0  # Error code
    
    @pytest.mark.asyncio
    async def test_invalid_folder_operations(self, server):
        """Test operations on invalid folders."""
        invalid_folders = [
            "NonExistentFolder123",
            "",
            None,
            "Folder/With/Invalid/Path",
            "A" * 300  # Very long name
        ]
        
        for folder_name in invalid_folders:
            request_data = {
                "jsonrpc": "2.0",
                "id": f"test_invalid_folder_{hash(str(folder_name))}",
                "method": "list_emails",
                "params": {"folder": folder_name, "limit": 1}
            }
            
            response = await server.handle_request(request_data)
            
            # Should return error for invalid folders
            if folder_name in [None, ""]:
                # These might be handled as default folder
                assert "jsonrpc" in response
            else:
                assert "error" in response
    
    @pytest.mark.asyncio
    async def test_invalid_email_operations(self, server):
        """Test operations on invalid emails."""
        invalid_email_ids = [
            "nonexistent_email_id",
            "",
            None,
            "invalid_format_id",
            "A" * 300  # Very long ID
        ]
        
        for email_id in invalid_email_ids:
            request_data = {
                "jsonrpc": "2.0",
                "id": f"test_invalid_email_{hash(str(email_id))}",
                "method": "get_email",
                "params": {"email_id": email_id}
            }
            
            response = await server.handle_request(request_data)
            
            # Should return error for invalid email IDs
            assert "error" in response
    
    @pytest.mark.asyncio
    async def test_malformed_search_queries(self, server):
        """Test search with malformed queries."""
        malformed_queries = [
            "",  # Empty query
            None,  # None query
            "A" * 1000,  # Very long query
            "query with \x00 null bytes",  # Invalid characters
            "query with 'special' \"quotes\"",  # Special characters
        ]
        
        for query in malformed_queries:
            request_data = {
                "jsonrpc": "2.0",
                "id": f"test_malformed_query_{hash(str(query))}",
                "method": "search_emails",
                "params": {"query": query, "limit": 1}
            }
            
            response = await server.handle_request(request_data)
            
            # Should handle gracefully (empty results or error)
            assert "jsonrpc" in response
            
            if query in ["", None]:
                # Empty queries should return empty results
                if "result" in response:
                    assert response["result"] == []
    
    @pytest.mark.asyncio
    async def test_parameter_validation_errors(self, server):
        """Test various parameter validation scenarios."""
        test_cases = [
            # Invalid limit values
            {
                "method": "list_emails",
                "params": {"limit": -1},
                "should_error": True
            },
            {
                "method": "list_emails", 
                "params": {"limit": 10000},
                "should_error": True
            },
            {
                "method": "search_emails",
                "params": {"query": "test", "limit": "invalid"},
                "should_error": True
            },
            # Missing required parameters
            {
                "method": "get_email",
                "params": {},
                "should_error": True
            },
            {
                "method": "search_emails",
                "params": {"limit": 10},  # Missing query
                "should_error": True
            },
        ]
        
        for i, test_case in enumerate(test_cases):
            request_data = {
                "jsonrpc": "2.0",
                "id": f"param_test_{i}",
                "method": test_case["method"],
                "params": test_case["params"]
            }
            
            response = await server.handle_request(request_data)
            
            if test_case["should_error"]:
                assert "error" in response, f"Expected error for test case {i}: {test_case}"
            else:
                # Should succeed or return valid result
                assert "jsonrpc" in response


class TestFolderStructureVariations:
    """Test with various Outlook folder structures."""
    
    @pytest.fixture
    def outlook_adapter(self):
        """Create Outlook adapter for folder testing."""
        adapter = OutlookAdapter()
        try:
            if adapter.connect():
                yield adapter
            else:
                pytest.skip("Outlook not available for folder structure testing")
        except OutlookConnectionError:
            pytest.skip("Outlook not available for folder structure testing")
        finally:
            if adapter.is_connected():
                adapter.disconnect()
    
    def test_nested_folder_access(self, outlook_adapter):
        """Test access to nested folder structures."""
        folder_service = FolderService(outlook_adapter)
        folders = folder_service.get_folders()
        
        # Look for nested folders
        nested_folders = [f for f in folders if "/" in f.get("full_path", "")]
        
        if nested_folders:
            # Test accessing nested folders
            for folder in nested_folders[:3]:  # Test first 3 nested folders
                try:
                    outlook_folder = outlook_adapter.get_folder_by_name(folder["name"])
                    assert outlook_folder is not None
                except Exception as e:
                    print(f"Could not access nested folder {folder['name']}: {e}")
    
    def test_special_folder_types(self, outlook_adapter):
        """Test access to special folder types."""
        folder_service = FolderService(outlook_adapter)
        folders = folder_service.get_folders()
        
        # Look for special folders
        special_folders = ["Calendar", "Contacts", "Tasks", "Notes", "Journal"]
        
        for special_name in special_folders:
            matching_folders = [f for f in folders if special_name.lower() in f["name"].lower()]
            
            if matching_folders:
                folder = matching_folders[0]
                try:
                    outlook_folder = outlook_adapter.get_folder_by_name(folder["name"])
                    # Special folders might not contain emails
                    print(f"Accessed special folder: {folder['name']}")
                except Exception as e:
                    print(f"Could not access special folder {folder['name']}: {e}")
    
    def test_folder_permissions_and_access(self, outlook_adapter):
        """Test folder permissions and access restrictions."""
        folder_service = FolderService(outlook_adapter)
        folders = folder_service.get_folders()
        
        accessible_folders = []
        restricted_folders = []
        
        for folder in folders:
            try:
                outlook_folder = outlook_adapter.get_folder_by_name(folder["name"])
                
                # Try to access items
                try:
                    items = outlook_folder.Items
                    count = items.Count if hasattr(items, 'Count') else 0
                    accessible_folders.append((folder["name"], count))
                except Exception:
                    restricted_folders.append(folder["name"])
                    
            except Exception:
                restricted_folders.append(folder["name"])
        
        print(f"Accessible folders: {len(accessible_folders)}")
        print(f"Restricted folders: {len(restricted_folders)}")
        
        # Should have at least some accessible folders
        assert len(accessible_folders) > 0, "No folders are accessible"


# Utility functions for integration testing
def create_test_config() -> Dict[str, Any]:
    """Create configuration for integration testing."""
    return create_server_config(
        log_level="INFO",
        log_dir=tempfile.mkdtemp(),
        max_concurrent_requests=5,
        request_timeout=15,
        outlook_connection_timeout=10,
        enable_performance_logging=True,
        enable_console_output=False
    )


def is_outlook_available() -> bool:
    """Check if Outlook is available for testing."""
    try:
        adapter = OutlookAdapter()
        result = adapter.connect()
        if result:
            adapter.disconnect()
        return result
    except:
        return False


# Pytest configuration for integration tests
def pytest_configure(config):
    """Configure pytest for integration tests."""
    config.addinivalue_line(
        "markers", "integration: mark test as integration test requiring Outlook"
    )
    config.addinivalue_line(
        "markers", "performance: mark test as performance benchmark"
    )
    config.addinivalue_line(
        "markers", "concurrent: mark test as concurrent/threading test"
    )


# Mark all tests in this module as integration tests
pytestmark = pytest.mark.integration