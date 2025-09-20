# Outlook MCP Server - Usage Examples

This document provides practical examples and use cases for the Outlook MCP Server, demonstrating how to integrate and use the server in various scenarios.

## Table of Contents

- [Basic Usage Examples](#basic-usage-examples)
- [Advanced Use Cases](#advanced-use-cases)
- [Integration Examples](#integration-examples)
- [Error Handling Examples](#error-handling-examples)
- [Performance Optimization](#performance-optimization)
- [Real-World Scenarios](#real-world-scenarios)

## Basic Usage Examples

### Example 1: Getting Started - List Recent Emails

This example shows how to connect to the server and list recent emails from the Inbox.

```python
import json
import asyncio
from mcp_client import MCPClient  # Hypothetical MCP client library

async def list_recent_emails():
    """List the 10 most recent emails from Inbox."""
    
    # Connect to the Outlook MCP Server
    client = MCPClient("outlook-mcp-server")
    await client.connect()
    
    try:
        # Request recent emails
        request = {
            "jsonrpc": "2.0",
            "id": "1",
            "method": "list_emails",
            "params": {
                "folder": "Inbox",
                "limit": 10
            }
        }
        
        response = await client.send_request(request)
        
        if "result" in response:
            emails = response["result"]
            print(f"Found {len(emails)} recent emails:")
            
            for email in emails:
                print(f"- {email['subject']} (from: {email['sender']})")
                print(f"  Received: {email['received_time']}")
                print(f"  Read: {'Yes' if email['is_read'] else 'No'}")
                print()
        else:
            print(f"Error: {response['error']['message']}")
            
    finally:
        await client.disconnect()

# Run the example
asyncio.run(list_recent_emails())
```

### Example 2: Reading a Specific Email

```python
async def read_email_details(email_id):
    """Get detailed information for a specific email."""
    
    client = MCPClient("outlook-mcp-server")
    await client.connect()
    
    try:
        request = {
            "jsonrpc": "2.0",
            "id": "2",
            "method": "get_email",
            "params": {
                "email_id": email_id
            }
        }
        
        response = await client.send_request(request)
        
        if "result" in response:
            email = response["result"]
            
            print(f"Subject: {email['subject']}")
            print(f"From: {email['sender']} <{email['sender_email']}>")
            print(f"To: {', '.join(email['recipients'])}")
            
            if email['cc_recipients']:
                print(f"CC: {', '.join(email['cc_recipients'])}")
            
            print(f"Received: {email['received_time']}")
            print(f"Size: {email['size']} bytes")
            
            if email['has_attachments']:
                print(f"Attachments: {email['attachment_count']}")
                for attachment in email.get('attachments', []):
                    print(f"  - {attachment['name']} ({attachment['size']} bytes)")
            
            print(f"\nBody:\n{email['body']}")
            
        else:
            print(f"Error: {response['error']['message']}")
            
    finally:
        await client.disconnect()

# Example usage
email_id = "AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK"
asyncio.run(read_email_details(email_id))
```

### Example 3: Searching for Emails

```python
async def search_project_emails():
    """Search for emails related to a specific project."""
    
    client = MCPClient("outlook-mcp-server")
    await client.connect()
    
    try:
        # Search for project-related emails
        request = {
            "jsonrpc": "2.0",
            "id": "3",
            "method": "search_emails",
            "params": {
                "query": "project alpha AND (status OR update)",
                "limit": 25
            }
        }
        
        response = await client.send_request(request)
        
        if "result" in response:
            emails = response["result"]
            
            print(f"Found {len(emails)} project-related emails:")
            
            # Group by sender
            by_sender = {}
            for email in emails:
                sender = email['sender']
                if sender not in by_sender:
                    by_sender[sender] = []
                by_sender[sender].append(email)
            
            for sender, sender_emails in by_sender.items():
                print(f"\nFrom {sender} ({len(sender_emails)} emails):")
                for email in sender_emails:
                    print(f"  - {email['subject']}")
                    print(f"    {email['received_time']} | {email['folder_name']}")
        else:
            print(f"Error: {response['error']['message']}")
            
    finally:
        await client.disconnect()

asyncio.run(search_project_emails())
```

### Example 4: Exploring Folder Structure

```python
async def explore_folders():
    """List all available folders and their contents."""
    
    client = MCPClient("outlook-mcp-server")
    await client.connect()
    
    try:
        # Get all folders
        request = {
            "jsonrpc": "2.0",
            "id": "4",
            "method": "get_folders",
            "params": {}
        }
        
        response = await client.send_request(request)
        
        if "result" in response:
            folders = response["result"]
            
            print("Available folders:")
            print("=" * 50)
            
            # Sort folders by full path for better organization
            folders.sort(key=lambda f: f['full_path'])
            
            for folder in folders:
                indent = "  " * (folder['full_path'].count('/') - 1)
                unread_info = f" ({folder['unread_count']} unread)" if folder['unread_count'] > 0 else ""
                
                print(f"{indent}{folder['name']} - {folder['item_count']} items{unread_info}")
                
        else:
            print(f"Error: {response['error']['message']}")
            
    finally:
        await client.disconnect()

asyncio.run(explore_folders())
```

## Advanced Use Cases

### Example 5: Email Processing Pipeline

```python
import asyncio
from datetime import datetime, timedelta
from typing import List, Dict, Any

class EmailProcessor:
    """Advanced email processing with filtering and analysis."""
    
    def __init__(self):
        self.client = None
    
    async def connect(self):
        """Connect to the MCP server."""
        self.client = MCPClient("outlook-mcp-server")
        await self.client.connect()
    
    async def disconnect(self):
        """Disconnect from the MCP server."""
        if self.client:
            await self.client.disconnect()
    
    async def process_unread_emails(self, folder: str = "Inbox") -> Dict[str, Any]:
        """Process all unread emails and categorize them."""
        
        # Get unread emails
        request = {
            "jsonrpc": "2.0",
            "id": "process_unread",
            "method": "list_emails",
            "params": {
                "folder": folder,
                "unread_only": True,
                "limit": 100
            }
        }
        
        response = await self.client.send_request(request)
        
        if "error" in response:
            raise Exception(f"Failed to get emails: {response['error']['message']}")
        
        emails = response["result"]
        
        # Categorize emails
        categories = {
            "urgent": [],
            "meetings": [],
            "projects": [],
            "notifications": [],
            "other": []
        }
        
        for email in emails:
            category = self._categorize_email(email)
            categories[category].append(email)
        
        # Generate summary
        summary = {
            "total_unread": len(emails),
            "categories": {cat: len(emails) for cat, emails in categories.items()},
            "urgent_count": len(categories["urgent"]),
            "needs_attention": categories["urgent"] + categories["meetings"]
        }
        
        return {
            "summary": summary,
            "categories": categories
        }
    
    def _categorize_email(self, email: Dict[str, Any]) -> str:
        """Categorize an email based on its content."""
        subject = email.get("subject", "").lower()
        sender = email.get("sender", "").lower()
        importance = email.get("importance", "Normal")
        
        # Check for urgent emails
        if importance == "High" or any(word in subject for word in ["urgent", "asap", "emergency", "critical"]):
            return "urgent"
        
        # Check for meeting-related emails
        if any(word in subject for word in ["meeting", "calendar", "appointment", "schedule", "invite"]):
            return "meetings"
        
        # Check for project-related emails
        if any(word in subject for word in ["project", "milestone", "deliverable", "sprint", "release"]):
            return "projects"
        
        # Check for notifications
        if any(word in sender for word in ["noreply", "notification", "automated", "system"]):
            return "notifications"
        
        return "other"
    
    async def find_emails_from_timeframe(self, days_back: int = 7, sender_filter: str = None) -> List[Dict[str, Any]]:
        """Find emails from a specific timeframe, optionally filtered by sender."""
        
        # Calculate date range
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days_back)
        
        # Build search query
        query_parts = [f"received:{start_date.strftime('%Y-%m-%d')}..{end_date.strftime('%Y-%m-%d')}"]
        
        if sender_filter:
            query_parts.append(f"from:{sender_filter}")
        
        query = " AND ".join(query_parts)
        
        request = {
            "jsonrpc": "2.0",
            "id": "timeframe_search",
            "method": "search_emails",
            "params": {
                "query": query,
                "limit": 200
            }
        }
        
        response = await self.client.send_request(request)
        
        if "error" in response:
            raise Exception(f"Search failed: {response['error']['message']}")
        
        return response["result"]

# Usage example
async def run_email_processing():
    processor = EmailProcessor()
    
    try:
        await processor.connect()
        
        # Process unread emails
        print("Processing unread emails...")
        result = await processor.process_unread_emails()
        
        print(f"Summary: {result['summary']['total_unread']} unread emails")
        print(f"Urgent: {result['summary']['urgent_count']}")
        print(f"Categories: {result['summary']['categories']}")
        
        # Show urgent emails
        if result['categories']['urgent']:
            print("\nURGENT EMAILS:")
            for email in result['categories']['urgent']:
                print(f"- {email['subject']} (from: {email['sender']})")
        
        # Find recent emails from manager
        print("\nRecent emails from manager...")
        manager_emails = await processor.find_emails_from_timeframe(
            days_back=3, 
            sender_filter="manager@company.com"
        )
        
        print(f"Found {len(manager_emails)} emails from manager in last 3 days")
        
    finally:
        await processor.disconnect()

asyncio.run(run_email_processing())
```

### Example 6: Email Analytics Dashboard

```python
import asyncio
from collections import defaultdict, Counter
from datetime import datetime, timedelta
import json

class EmailAnalytics:
    """Generate analytics and insights from email data."""
    
    def __init__(self):
        self.client = None
    
    async def connect(self):
        self.client = MCPClient("outlook-mcp-server")
        await self.client.connect()
    
    async def disconnect(self):
        if self.client:
            await self.client.disconnect()
    
    async def generate_weekly_report(self) -> Dict[str, Any]:
        """Generate a comprehensive weekly email report."""
        
        # Get emails from the last week
        end_date = datetime.now()
        start_date = end_date - timedelta(days=7)
        
        query = f"received:{start_date.strftime('%Y-%m-%d')}..{end_date.strftime('%Y-%m-%d')}"
        
        request = {
            "jsonrpc": "2.0",
            "id": "weekly_report",
            "method": "search_emails",
            "params": {
                "query": query,
                "limit": 1000
            }
        }
        
        response = await self.client.send_request(request)
        
        if "error" in response:
            raise Exception(f"Failed to get weekly emails: {response['error']['message']}")
        
        emails = response["result"]
        
        # Analyze the data
        analysis = {
            "period": {
                "start": start_date.isoformat(),
                "end": end_date.isoformat(),
                "total_emails": len(emails)
            },
            "daily_breakdown": self._analyze_daily_volume(emails),
            "top_senders": self._analyze_top_senders(emails),
            "folder_distribution": self._analyze_folder_distribution(emails),
            "response_patterns": self._analyze_response_patterns(emails),
            "attachment_stats": self._analyze_attachments(emails),
            "unread_analysis": self._analyze_unread_emails(emails)
        }
        
        return analysis
    
    def _analyze_daily_volume(self, emails: List[Dict]) -> Dict[str, int]:
        """Analyze email volume by day."""
        daily_counts = defaultdict(int)
        
        for email in emails:
            try:
                received_date = datetime.fromisoformat(email['received_time'].replace('Z', '+00:00'))
                day_key = received_date.strftime('%Y-%m-%d')
                daily_counts[day_key] += 1
            except (ValueError, KeyError):
                continue
        
        return dict(daily_counts)
    
    def _analyze_top_senders(self, emails: List[Dict], top_n: int = 10) -> List[Dict]:
        """Find the most frequent senders."""
        sender_counts = Counter()
        
        for email in emails:
            sender = email.get('sender_email', email.get('sender', 'Unknown'))
            sender_counts[sender] += 1
        
        return [
            {"sender": sender, "count": count}
            for sender, count in sender_counts.most_common(top_n)
        ]
    
    def _analyze_folder_distribution(self, emails: List[Dict]) -> Dict[str, int]:
        """Analyze email distribution across folders."""
        folder_counts = Counter()
        
        for email in emails:
            folder = email.get('folder_name', 'Unknown')
            folder_counts[folder] += 1
        
        return dict(folder_counts)
    
    def _analyze_response_patterns(self, emails: List[Dict]) -> Dict[str, Any]:
        """Analyze response patterns and email threads."""
        subjects = Counter()
        reply_indicators = 0
        forward_indicators = 0
        
        for email in emails:
            subject = email.get('subject', '')
            subjects[subject] += 1
            
            if subject.lower().startswith(('re:', 'reply:')):
                reply_indicators += 1
            elif subject.lower().startswith(('fwd:', 'fw:', 'forward:')):
                forward_indicators += 1
        
        # Find potential email threads (same subject, multiple emails)
        threads = [(subject, count) for subject, count in subjects.items() if count > 1]
        threads.sort(key=lambda x: x[1], reverse=True)
        
        return {
            "replies": reply_indicators,
            "forwards": forward_indicators,
            "potential_threads": len(threads),
            "top_threads": threads[:5]
        }
    
    def _analyze_attachments(self, emails: List[Dict]) -> Dict[str, Any]:
        """Analyze attachment patterns."""
        total_with_attachments = sum(1 for email in emails if email.get('has_attachments', False))
        total_attachment_count = sum(email.get('attachment_count', 0) for email in emails)
        
        return {
            "emails_with_attachments": total_with_attachments,
            "total_attachments": total_attachment_count,
            "percentage_with_attachments": (total_with_attachments / len(emails) * 100) if emails else 0
        }
    
    def _analyze_unread_emails(self, emails: List[Dict]) -> Dict[str, Any]:
        """Analyze unread email patterns."""
        unread_emails = [email for email in emails if not email.get('is_read', True)]
        
        if not unread_emails:
            return {"unread_count": 0, "unread_percentage": 0}
        
        # Analyze unread by sender
        unread_senders = Counter()
        for email in unread_emails:
            sender = email.get('sender', 'Unknown')
            unread_senders[sender] += 1
        
        return {
            "unread_count": len(unread_emails),
            "unread_percentage": len(unread_emails) / len(emails) * 100,
            "top_unread_senders": [
                {"sender": sender, "count": count}
                for sender, count in unread_senders.most_common(5)
            ]
        }

# Usage example
async def generate_analytics_report():
    analytics = EmailAnalytics()
    
    try:
        await analytics.connect()
        
        print("Generating weekly email analytics report...")
        report = await analytics.generate_weekly_report()
        
        # Print summary
        print(f"\n📊 WEEKLY EMAIL REPORT")
        print(f"Period: {report['period']['start'][:10]} to {report['period']['end'][:10]}")
        print(f"Total Emails: {report['period']['total_emails']}")
        
        # Daily breakdown
        print(f"\n📅 Daily Volume:")
        for day, count in sorted(report['daily_breakdown'].items()):
            print(f"  {day}: {count} emails")
        
        # Top senders
        print(f"\n👥 Top Senders:")
        for sender_info in report['top_senders'][:5]:
            print(f"  {sender_info['sender']}: {sender_info['count']} emails")
        
        # Folder distribution
        print(f"\n📁 Folder Distribution:")
        for folder, count in report['folder_distribution'].items():
            print(f"  {folder}: {count} emails")
        
        # Response patterns
        patterns = report['response_patterns']
        print(f"\n💬 Response Patterns:")
        print(f"  Replies: {patterns['replies']}")
        print(f"  Forwards: {patterns['forwards']}")
        print(f"  Active Threads: {patterns['potential_threads']}")
        
        # Attachment stats
        attachments = report['attachment_stats']
        print(f"\n📎 Attachments:")
        print(f"  Emails with attachments: {attachments['emails_with_attachments']}")
        print(f"  Total attachments: {attachments['total_attachments']}")
        print(f"  Percentage: {attachments['percentage_with_attachments']:.1f}%")
        
        # Unread analysis
        unread = report['unread_analysis']
        print(f"\n📬 Unread Emails:")
        print(f"  Unread count: {unread['unread_count']}")
        print(f"  Unread percentage: {unread['unread_percentage']:.1f}%")
        
        # Save detailed report to file
        with open(f"email_report_{datetime.now().strftime('%Y%m%d')}.json", 'w') as f:
            json.dump(report, f, indent=2, default=str)
        
        print(f"\n✅ Detailed report saved to email_report_{datetime.now().strftime('%Y%m%d')}.json")
        
    finally:
        await analytics.disconnect()

asyncio.run(generate_analytics_report())
```

## Integration Examples

### Example 7: Flask Web API Integration

```python
from flask import Flask, jsonify, request
import asyncio
from concurrent.futures import ThreadPoolExecutor
import threading

app = Flask(__name__)
executor = ThreadPoolExecutor(max_workers=4)

class OutlookAPIWrapper:
    """Thread-safe wrapper for Outlook MCP Server integration."""
    
    def __init__(self):
        self._local = threading.local()
    
    def get_client(self):
        """Get thread-local MCP client."""
        if not hasattr(self._local, 'client'):
            self._local.client = MCPClient("outlook-mcp-server")
        return self._local.client
    
    async def list_emails_async(self, folder=None, unread_only=False, limit=50):
        """Async wrapper for listing emails."""
        client = self.get_client()
        
        if not client.is_connected():
            await client.connect()
        
        request_data = {
            "jsonrpc": "2.0",
            "id": "api_list",
            "method": "list_emails",
            "params": {
                "folder": folder,
                "unread_only": unread_only,
                "limit": limit
            }
        }
        
        response = await client.send_request(request_data)
        
        if "error" in response:
            raise Exception(response["error"]["message"])
        
        return response["result"]

outlook_wrapper = OutlookAPIWrapper()

def run_async(coro):
    """Run async function in thread pool."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        return loop.run_until_complete(coro)
    finally:
        loop.close()

@app.route('/api/emails', methods=['GET'])
def api_list_emails():
    """REST API endpoint for listing emails."""
    try:
        # Get query parameters
        folder = request.args.get('folder')
        unread_only = request.args.get('unread_only', 'false').lower() == 'true'
        limit = min(int(request.args.get('limit', 50)), 1000)
        
        # Run async operation in thread pool
        future = executor.submit(
            run_async,
            outlook_wrapper.list_emails_async(folder, unread_only, limit)
        )
        
        emails = future.result(timeout=30)  # 30 second timeout
        
        return jsonify({
            "success": True,
            "data": emails,
            "count": len(emails)
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/emails/<email_id>', methods=['GET'])
def api_get_email(email_id):
    """REST API endpoint for getting specific email."""
    try:
        async def get_email():
            client = outlook_wrapper.get_client()
            if not client.is_connected():
                await client.connect()
            
            request_data = {
                "jsonrpc": "2.0",
                "id": "api_get",
                "method": "get_email",
                "params": {"email_id": email_id}
            }
            
            response = await client.send_request(request_data)
            
            if "error" in response:
                raise Exception(response["error"]["message"])
            
            return response["result"]
        
        future = executor.submit(run_async, get_email())
        email = future.result(timeout=30)
        
        return jsonify({
            "success": True,
            "data": email
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

@app.route('/api/search', methods=['GET'])
def api_search_emails():
    """REST API endpoint for searching emails."""
    try:
        query = request.args.get('q')
        if not query:
            return jsonify({
                "success": False,
                "error": "Query parameter 'q' is required"
            }), 400
        
        folder = request.args.get('folder')
        limit = min(int(request.args.get('limit', 50)), 1000)
        
        async def search_emails():
            client = outlook_wrapper.get_client()
            if not client.is_connected():
                await client.connect()
            
            request_data = {
                "jsonrpc": "2.0",
                "id": "api_search",
                "method": "search_emails",
                "params": {
                    "query": query,
                    "folder": folder,
                    "limit": limit
                }
            }
            
            response = await client.send_request(request_data)
            
            if "error" in response:
                raise Exception(response["error"]["message"])
            
            return response["result"]
        
        future = executor.submit(run_async, search_emails())
        results = future.result(timeout=30)
        
        return jsonify({
            "success": True,
            "data": results,
            "count": len(results),
            "query": query
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
```

## Error Handling Examples

### Example 8: Robust Error Handling

```python
import asyncio
import logging
from typing import Optional, Dict, Any

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class RobustEmailClient:
    """Email client with comprehensive error handling and retry logic."""
    
    def __init__(self, max_retries: int = 3, retry_delay: float = 1.0):
        self.client = None
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.connected = False
    
    async def connect_with_retry(self) -> bool:
        """Connect to MCP server with retry logic."""
        for attempt in range(self.max_retries):
            try:
                self.client = MCPClient("outlook-mcp-server")
                await self.client.connect()
                self.connected = True
                logger.info("Successfully connected to Outlook MCP Server")
                return True
                
            except Exception as e:
                logger.warning(f"Connection attempt {attempt + 1} failed: {str(e)}")
                if attempt < self.max_retries - 1:
                    await asyncio.sleep(self.retry_delay * (2 ** attempt))  # Exponential backoff
                else:
                    logger.error("All connection attempts failed")
                    return False
        
        return False
    
    async def safe_request(self, method: str, params: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Make a safe request with error handling and retries."""
        if not self.connected:
            if not await self.connect_with_retry():
                return None
        
        request_data = {
            "jsonrpc": "2.0",
            "id": f"safe_{method}",
            "method": method,
            "params": params
        }
        
        for attempt in range(self.max_retries):
            try:
                response = await self.client.send_request(request_data)
                
                if "result" in response:
                    return response["result"]
                elif "error" in response:
                    error = response["error"]
                    error_code = error.get("code", 0)
                    error_message = error.get("message", "Unknown error")
                    error_type = error.get("data", {}).get("type", "UnknownError")
                    
                    # Handle different error types
                    if error_type == "OutlookConnectionError":
                        logger.warning("Outlook connection lost, attempting to reconnect...")
                        self.connected = False
                        if await self.connect_with_retry():
                            continue  # Retry the request
                        else:
                            logger.error("Failed to reconnect to Outlook")
                            return None
                    
                    elif error_type == "RateLimitError":
                        retry_after = error.get("data", {}).get("retry_after", 60)
                        logger.warning(f"Rate limit exceeded, waiting {retry_after} seconds...")
                        await asyncio.sleep(retry_after)
                        continue  # Retry the request
                    
                    elif error_type in ["ValidationError", "EmailNotFoundError", "FolderNotFoundError"]:
                        # These are client errors, don't retry
                        logger.error(f"Client error: {error_message}")
                        return None
                    
                    elif error_type == "PermissionError":
                        logger.error(f"Permission denied: {error_message}")
                        return None
                    
                    else:
                        # Unknown error, log and potentially retry
                        logger.error(f"Unknown error (attempt {attempt + 1}): {error_message}")
                        if attempt < self.max_retries - 1:
                            await asyncio.sleep(self.retry_delay)
                            continue
                        else:
                            return None
                
            except asyncio.TimeoutError:
                logger.warning(f"Request timeout (attempt {attempt + 1})")
                if attempt < self.max_retries - 1:
                    await asyncio.sleep(self.retry_delay)
                    continue
                else:
                    logger.error("Request timed out after all retries")
                    return None
            
            except Exception as e:
                logger.error(f"Unexpected error (attempt {attempt + 1}): {str(e)}")
                if attempt < self.max_retries - 1:
                    await asyncio.sleep(self.retry_delay)
                    continue
                else:
                    logger.error("Request failed after all retries")
                    return None
        
        return None
    
    async def list_emails_safe(self, folder: str = None, unread_only: bool = False, limit: int = 50) -> Optional[List[Dict[str, Any]]]:
        """Safely list emails with error handling."""
        params = {
            "folder": folder,
            "unread_only": unread_only,
            "limit": limit
        }
        
        result = await self.safe_request("list_emails", params)
        return result if result is not None else []
    
    async def get_email_safe(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Safely get email with error handling."""
        if not email_id:
            logger.error("Email ID is required")
            return None
        
        params = {"email_id": email_id}
        return await self.safe_request("get_email", params)
    
    async def search_emails_safe(self, query: str, folder: str = None, limit: int = 50) -> Optional[List[Dict[str, Any]]]:
        """Safely search emails with error handling."""
        if not query or not query.strip():
            logger.error("Search query is required")
            return []
        
        params = {
            "query": query.strip(),
            "folder": folder,
            "limit": limit
        }
        
        result = await self.safe_request("search_emails", params)
        return result if result is not None else []
    
    async def disconnect(self):
        """Safely disconnect from the server."""
        if self.client and self.connected:
            try:
                await self.client.disconnect()
                logger.info("Disconnected from Outlook MCP Server")
            except Exception as e:
                logger.warning(f"Error during disconnect: {str(e)}")
            finally:
                self.connected = False

# Usage example with comprehensive error handling
async def robust_email_processing():
    """Example of robust email processing with error handling."""
    
    client = RobustEmailClient(max_retries=3, retry_delay=2.0)
    
    try:
        # Connect with retry logic
        if not await client.connect_with_retry():
            print("❌ Failed to connect to Outlook MCP Server")
            return
        
        print("✅ Connected to Outlook MCP Server")
        
        # Try to list emails
        print("\n📧 Listing recent emails...")
        emails = await client.list_emails_safe(folder="Inbox", limit=10)
        
        if emails:
            print(f"✅ Found {len(emails)} emails")
            
            # Try to get details for the first email
            if emails:
                first_email_id = emails[0]["id"]
                print(f"\n📖 Getting details for first email...")
                
                email_details = await client.get_email_safe(first_email_id)
                
                if email_details:
                    print(f"✅ Retrieved email: {email_details['subject']}")
                else:
                    print("❌ Failed to retrieve email details")
        else:
            print("❌ No emails found or failed to retrieve emails")
        
        # Try searching
        print(f"\n🔍 Searching for emails...")
        search_results = await client.search_emails_safe("meeting", limit=5)
        
        if search_results:
            print(f"✅ Found {len(search_results)} search results")
        else:
            print("❌ No search results found or search failed")
        
    except Exception as e:
        logger.error(f"Unexpected error in main processing: {str(e)}")
        print(f"❌ Unexpected error: {str(e)}")
    
    finally:
        await client.disconnect()
        print("\n👋 Disconnected from server")

# Run the robust example
asyncio.run(robust_email_processing())
```

## Performance Optimization

### Example 9: Batch Processing and Caching

```python
import asyncio
from typing import List, Dict, Any, Set
import time
from dataclasses import dataclass
from collections import defaultdict

@dataclass
class EmailBatch:
    """Represents a batch of emails for processing."""
    emails: List[Dict[str, Any]]
    folder: str
    timestamp: float

class OptimizedEmailProcessor:
    """Optimized email processor with caching and batch operations."""
    
    def __init__(self, cache_ttl: int = 300):  # 5 minute cache TTL
        self.client = None
        self.email_cache: Dict[str, Dict[str, Any]] = {}
        self.folder_cache: Dict[str, EmailBatch] = {}
        self.cache_ttl = cache_ttl
        self.processed_emails: Set[str] = set()
    
    async def connect(self):
        """Connect to MCP server."""
        self.client = MCPClient("outlook-mcp-server")
        await self.client.connect()
    
    async def disconnect(self):
        """Disconnect from MCP server."""
        if self.client:
            await self.client.disconnect()
    
    def _is_cache_valid(self, timestamp: float) -> bool:
        """Check if cache entry is still valid."""
        return time.time() - timestamp < self.cache_ttl
    
    async def get_emails_optimized(self, folder: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Get emails with caching optimization."""
        
        # Check cache first
        cache_key = f"{folder}:{limit}"
        if cache_key in self.folder_cache:
            cached_batch = self.folder_cache[cache_key]
            if self._is_cache_valid(cached_batch.timestamp):
                print(f"📋 Cache hit for folder '{folder}'")
                return cached_batch.emails
        
        # Cache miss or expired, fetch from server
        print(f"🌐 Fetching emails from server for folder '{folder}'")
        
        request = {
            "jsonrpc": "2.0",
            "id": f"optimized_{folder}",
            "method": "list_emails",
            "params": {
                "folder": folder,
                "limit": limit
            }
        }
        
        response = await self.client.send_request(request)
        
        if "error" in response:
            raise Exception(f"Failed to get emails: {response['error']['message']}")
        
        emails = response["result"]
        
        # Cache the results
        self.folder_cache[cache_key] = EmailBatch(
            emails=emails,
            folder=folder,
            timestamp=time.time()
        )
        
        # Cache individual emails
        for email in emails:
            self.email_cache[email["id"]] = email
        
        return emails
    
    async def get_email_details_batch(self, email_ids: List[str]) -> Dict[str, Dict[str, Any]]:
        """Get details for multiple emails efficiently."""
        
        results = {}
        uncached_ids = []
        
        # Check cache for each email
        for email_id in email_ids:
            if email_id in self.email_cache:
                cached_email = self.email_cache[email_id]
                if "body" in cached_email:  # Full details cached
                    results[email_id] = cached_email
                    continue
            
            uncached_ids.append(email_id)
        
        print(f"📋 Cache hits: {len(results)}, Cache misses: {len(uncached_ids)}")
        
        # Fetch uncached emails
        if uncached_ids:
            # Process in smaller batches to avoid overwhelming the server
            batch_size = 5
            for i in range(0, len(uncached_ids), batch_size):
                batch = uncached_ids[i:i + batch_size]
                
                # Create concurrent requests for the batch
                tasks = []
                for email_id in batch:
                    task = self._get_single_email_details(email_id)
                    tasks.append(task)
                
                # Wait for all requests in the batch
                batch_results = await asyncio.gather(*tasks, return_exceptions=True)
                
                # Process results
                for email_id, result in zip(batch, batch_results):
                    if isinstance(result, Exception):
                        print(f"❌ Error getting email {email_id}: {str(result)}")
                        continue
                    
                    if result:
                        results[email_id] = result
                        self.email_cache[email_id] = result
                
                # Small delay between batches to be respectful
                if i + batch_size < len(uncached_ids):
                    await asyncio.sleep(0.1)
        
        return results
    
    async def _get_single_email_details(self, email_id: str) -> Dict[str, Any]:
        """Get details for a single email."""
        request = {
            "jsonrpc": "2.0",
            "id": f"details_{email_id}",
            "method": "get_email",
            "params": {"email_id": email_id}
        }
        
        response = await self.client.send_request(request)
        
        if "error" in response:
            raise Exception(f"Failed to get email details: {response['error']['message']}")
        
        return response["result"]
    
    async def process_emails_efficiently(self, folders: List[str], process_func) -> Dict[str, Any]:
        """Process emails from multiple folders efficiently."""
        
        start_time = time.time()
        
        # Step 1: Get email lists for all folders concurrently
        print("🚀 Fetching email lists from all folders...")
        
        folder_tasks = []
        for folder in folders:
            task = self.get_emails_optimized(folder, limit=50)
            folder_tasks.append((folder, task))
        
        folder_results = {}
        for folder, task in folder_tasks:
            try:
                emails = await task
                folder_results[folder] = emails
                print(f"✅ {folder}: {len(emails)} emails")
            except Exception as e:
                print(f"❌ {folder}: Error - {str(e)}")
                folder_results[folder] = []
        
        # Step 2: Collect all unique email IDs that need detailed processing
        all_email_ids = set()
        for emails in folder_results.values():
            for email in emails:
                if email["id"] not in self.processed_emails:
                    all_email_ids.add(email["id"])
        
        print(f"📊 Total unique emails to process: {len(all_email_ids)}")
        
        # Step 3: Get detailed information for all emails
        if all_email_ids:
            print("📖 Getting detailed email information...")
            detailed_emails = await self.get_email_details_batch(list(all_email_ids))
        else:
            detailed_emails = {}
        
        # Step 4: Process emails using the provided function
        print("⚙️ Processing emails...")
        
        processing_results = defaultdict(list)
        total_processed = 0
        
        for folder, emails in folder_results.items():
            for email in emails:
                email_id = email["id"]
                
                # Skip if already processed
                if email_id in self.processed_emails:
                    continue
                
                # Get detailed email data
                detailed_email = detailed_emails.get(email_id, email)
                
                # Process the email
                try:
                    result = await process_func(detailed_email, folder)
                    if result:
                        processing_results[folder].append(result)
                    
                    self.processed_emails.add(email_id)
                    total_processed += 1
                    
                except Exception as e:
                    print(f"❌ Error processing email {email_id}: {str(e)}")
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        # Return summary
        return {
            "folders_processed": len(folders),
            "total_emails_found": sum(len(emails) for emails in folder_results.values()),
            "emails_processed": total_processed,
            "processing_time_seconds": processing_time,
            "emails_per_second": total_processed / processing_time if processing_time > 0 else 0,
            "results_by_folder": dict(processing_results),
            "cache_stats": {
                "cached_emails": len(self.email_cache),
                "cached_folders": len(self.folder_cache)
            }
        }
    
    def clear_cache(self):
        """Clear all caches."""
        self.email_cache.clear()
        self.folder_cache.clear()
        self.processed_emails.clear()
        print("🧹 Cache cleared")

# Example processing function
async def analyze_email_sentiment(email: Dict[str, Any], folder: str) -> Dict[str, Any]:
    """Example processing function that analyzes email sentiment."""
    
    # Simulate processing time
    await asyncio.sleep(0.01)
    
    subject = email.get("subject", "")
    body = email.get("body", "")
    
    # Simple sentiment analysis (in real use, you'd use a proper library)
    positive_words = ["great", "excellent", "good", "thanks", "appreciate", "wonderful"]
    negative_words = ["problem", "issue", "error", "failed", "urgent", "critical"]
    
    text = (subject + " " + body).lower()
    
    positive_count = sum(1 for word in positive_words if word in text)
    negative_count = sum(1 for word in negative_words if word in text)
    
    if positive_count > negative_count:
        sentiment = "positive"
    elif negative_count > positive_count:
        sentiment = "negative"
    else:
        sentiment = "neutral"
    
    return {
        "email_id": email["id"],
        "subject": subject,
        "sender": email.get("sender", ""),
        "sentiment": sentiment,
        "positive_score": positive_count,
        "negative_score": negative_count,
        "folder": folder
    }

# Usage example
async def run_optimized_processing():
    """Run optimized email processing example."""
    
    processor = OptimizedEmailProcessor(cache_ttl=600)  # 10 minute cache
    
    try:
        await processor.connect()
        
        # Process emails from multiple folders
        folders_to_process = ["Inbox", "Sent Items", "Projects"]
        
        print("🚀 Starting optimized email processing...")
        
        results = await processor.process_emails_efficiently(
            folders=folders_to_process,
            process_func=analyze_email_sentiment
        )
        
        # Print results
        print(f"\n📊 PROCESSING SUMMARY")
        print(f"Folders processed: {results['folders_processed']}")
        print(f"Total emails found: {results['total_emails_found']}")
        print(f"Emails processed: {results['emails_processed']}")
        print(f"Processing time: {results['processing_time_seconds']:.2f} seconds")
        print(f"Processing rate: {results['emails_per_second']:.2f} emails/second")
        
        print(f"\n📋 CACHE STATISTICS")
        print(f"Cached emails: {results['cache_stats']['cached_emails']}")
        print(f"Cached folders: {results['cache_stats']['cached_folders']}")
        
        # Show sentiment analysis results
        print(f"\n😊 SENTIMENT ANALYSIS RESULTS")
        for folder, folder_results in results['results_by_folder'].items():
            if folder_results:
                sentiments = defaultdict(int)
                for result in folder_results:
                    sentiments[result['sentiment']] += 1
                
                print(f"{folder}: {dict(sentiments)}")
        
        # Run again to demonstrate caching
        print(f"\n🔄 Running again to demonstrate caching...")
        
        start_time = time.time()
        results2 = await processor.process_emails_efficiently(
            folders=folders_to_process,
            process_func=analyze_email_sentiment
        )
        
        print(f"Second run took {results2['processing_time_seconds']:.2f} seconds")
        print(f"Speed improvement: {results['processing_time_seconds'] / results2['processing_time_seconds']:.1f}x faster")
        
    finally:
        await processor.disconnect()

asyncio.run(run_optimized_processing())
```

## Real-World Scenarios

### Example 10: Email Monitoring and Alerting System

```python
import asyncio
import json
from datetime import datetime, timedelta
from typing import List, Dict, Any, Callable
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class EmailMonitor:
    """Real-world email monitoring and alerting system."""
    
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.client = None
        self.alert_handlers: List[Callable] = []
        self.monitoring_active = False
        
    async def connect(self):
        """Connect to MCP server."""
        self.client = MCPClient("outlook-mcp-server")
        await self.client.connect()
    
    async def disconnect(self):
        """Disconnect from MCP server."""
        if self.client:
            await self.client.disconnect()
    
    def add_alert_handler(self, handler: Callable):
        """Add an alert handler function."""
        self.alert_handlers.append(handler)
    
    async def start_monitoring(self):
        """Start continuous email monitoring."""
        self.monitoring_active = True
        
        print("🔍 Starting email monitoring...")
        
        while self.monitoring_active:
            try:
                await self._check_alerts()
                
                # Wait for next check interval
                check_interval = self.config.get("check_interval_seconds", 60)
                await asyncio.sleep(check_interval)
                
            except Exception as e:
                print(f"❌ Error during monitoring: {str(e)}")
                await asyncio.sleep(30)  # Wait before retrying
    
    def stop_monitoring(self):
        """Stop email monitoring."""
        self.monitoring_active = False
        print("⏹️ Email monitoring stopped")
    
    async def _check_alerts(self):
        """Check for alert conditions."""
        
        # Check for high-priority emails
        await self._check_high_priority_emails()
        
        # Check for emails from VIP senders
        await self._check_vip_emails()
        
        # Check for keyword alerts
        await self._check_keyword_alerts()
        
        # Check for unusual email volume
        await self._check_email_volume()
    
    async def _check_high_priority_emails(self):
        """Check for high-priority unread emails."""
        
        try:
            request = {
                "jsonrpc": "2.0",
                "id": "priority_check",
                "method": "list_emails",
                "params": {
                    "folder": "Inbox",
                    "unread_only": True,
                    "limit": 50
                }
            }
            
            response = await self.client.send_request(request)
            
            if "error" in response:
                return
            
            emails = response["result"]
            high_priority_emails = [
                email for email in emails 
                if email.get("importance", "Normal") == "High"
            ]
            
            if high_priority_emails:
                await self._trigger_alert("high_priority", {
                    "count": len(high_priority_emails),
                    "emails": high_priority_emails[:5]  # First 5 for alert
                })
                
        except Exception as e:
            print(f"Error checking high priority emails: {str(e)}")
    
    async def _check_vip_emails(self):
        """Check for emails from VIP senders."""
        
        vip_senders = self.config.get("vip_senders", [])
        if not vip_senders:
            return
        
        try:
            # Check each VIP sender
            for sender in vip_senders:
                query = f"from:{sender}"
                
                # Look for emails in the last hour
                one_hour_ago = datetime.now() - timedelta(hours=1)
                query += f" AND received:{one_hour_ago.strftime('%Y-%m-%d')}.."
                
                request = {
                    "jsonrpc": "2.0",
                    "id": f"vip_check_{sender}",
                    "method": "search_emails",
                    "params": {
                        "query": query,
                        "limit": 10
                    }
                }
                
                response = await self.client.send_request(request)
                
                if "error" in response:
                    continue
                
                emails = response["result"]
                unread_emails = [email for email in emails if not email.get("is_read", True)]
                
                if unread_emails:
                    await self._trigger_alert("vip_email", {
                        "sender": sender,
                        "count": len(unread_emails),
                        "emails": unread_emails
                    })
                    
        except Exception as e:
            print(f"Error checking VIP emails: {str(e)}")
    
    async def _check_keyword_alerts(self):
        """Check for emails containing alert keywords."""
        
        alert_keywords = self.config.get("alert_keywords", [])
        if not alert_keywords:
            return
        
        try:
            for keyword in alert_keywords:
                # Search for the keyword in recent emails
                one_hour_ago = datetime.now() - timedelta(hours=1)
                query = f"{keyword} AND received:{one_hour_ago.strftime('%Y-%m-%d')}.."
                
                request = {
                    "jsonrpc": "2.0",
                    "id": f"keyword_check_{keyword}",
                    "method": "search_emails",
                    "params": {
                        "query": query,
                        "limit": 20
                    }
                }
                
                response = await self.client.send_request(request)
                
                if "error" in response:
                    continue
                
                emails = response["result"]
                
                if emails:
                    await self._trigger_alert("keyword_match", {
                        "keyword": keyword,
                        "count": len(emails),
                        "emails": emails[:3]  # First 3 for alert
                    })
                    
        except Exception as e:
            print(f"Error checking keyword alerts: {str(e)}")
    
    async def _check_email_volume(self):
        """Check for unusual email volume."""
        
        try:
            # Get emails from the last hour
            one_hour_ago = datetime.now() - timedelta(hours=1)
            query = f"received:{one_hour_ago.strftime('%Y-%m-%d')}.."
            
            request = {
                "jsonrpc": "2.0",
                "id": "volume_check",
                "method": "search_emails",
                "params": {
                    "query": query,
                    "limit": 1000
                }
            }
            
            response = await self.client.send_request(request)
            
            if "error" in response:
                return
            
            emails = response["result"]
            email_count = len(emails)
            
            # Check against threshold
            volume_threshold = self.config.get("volume_threshold_per_hour", 100)
            
            if email_count > volume_threshold:
                await self._trigger_alert("high_volume", {
                    "count": email_count,
                    "threshold": volume_threshold,
                    "timeframe": "1 hour"
                })
                
        except Exception as e:
            print(f"Error checking email volume: {str(e)}")
    
    async def _trigger_alert(self, alert_type: str, data: Dict[str, Any]):
        """Trigger an alert with the given data."""
        
        alert = {
            "type": alert_type,
            "timestamp": datetime.now().isoformat(),
            "data": data
        }
        
        print(f"🚨 ALERT: {alert_type.upper()}")
        print(f"   Data: {json.dumps(data, indent=2, default=str)}")
        
        # Call all registered alert handlers
        for handler in self.alert_handlers:
            try:
                await handler(alert)
            except Exception as e:
                print(f"Error in alert handler: {str(e)}")

# Alert handler examples
async def console_alert_handler(alert: Dict[str, Any]):
    """Simple console alert handler."""
    alert_type = alert["type"]
    data = alert["data"]
    
    if alert_type == "high_priority":
        print(f"📢 {data['count']} high-priority emails need attention!")
        
    elif alert_type == "vip_email":
        print(f"⭐ New email from VIP: {data['sender']} ({data['count']} emails)")
        
    elif alert_type == "keyword_match":
        print(f"🔍 Keyword '{data['keyword']}' found in {data['count']} emails")
        
    elif alert_type == "high_volume":
        print(f"📊 High email volume: {data['count']} emails in {data['timeframe']}")

async def email_alert_handler(alert: Dict[str, Any]):
    """Email notification alert handler."""
    
    # This would send an email notification
    # Implementation depends on your email setup
    
    alert_type = alert["type"]
    data = alert["data"]
    
    subject = f"Email Alert: {alert_type.replace('_', ' ').title()}"
    
    body = f"""
    Alert Type: {alert_type}
    Timestamp: {alert['timestamp']}
    
    Details:
    {json.dumps(data, indent=2, default=str)}
    """
    
    print(f"📧 Would send email alert: {subject}")
    # In real implementation:
    # send_email_notification(subject, body, recipients)

async def slack_alert_handler(alert: Dict[str, Any]):
    """Slack notification alert handler."""
    
    # This would send a Slack message
    # Implementation depends on your Slack setup
    
    alert_type = alert["type"]
    data = alert["data"]
    
    message = f"🚨 *Email Alert: {alert_type.replace('_', ' ').title()}*\n"
    
    if alert_type == "high_priority":
        message += f"📢 {data['count']} high-priority emails need attention"
        
    elif alert_type == "vip_email":
        message += f"⭐ New email from VIP: {data['sender']}"
        
    print(f"💬 Would send Slack message: {message}")
    # In real implementation:
    # send_slack_message(message, channel)

# Usage example
async def run_email_monitoring():
    """Run the email monitoring system."""
    
    # Configuration
    config = {
        "check_interval_seconds": 30,  # Check every 30 seconds
        "vip_senders": [
            "ceo@company.com",
            "manager@company.com",
            "client@importantclient.com"
        ],
        "alert_keywords": [
            "urgent",
            "critical",
            "emergency",
            "server down",
            "production issue"
        ],
        "volume_threshold_per_hour": 50
    }
    
    # Create monitor
    monitor = EmailMonitor(config)
    
    # Add alert handlers
    monitor.add_alert_handler(console_alert_handler)
    monitor.add_alert_handler(email_alert_handler)
    monitor.add_alert_handler(slack_alert_handler)
    
    try:
        await monitor.connect()
        
        # Start monitoring (this will run indefinitely)
        monitoring_task = asyncio.create_task(monitor.start_monitoring())
        
        # Simulate running for a while (in real use, this would run continuously)
        print("🔍 Email monitoring started. Press Ctrl+C to stop.")
        
        try:
            await monitoring_task
        except KeyboardInterrupt:
            print("\n⏹️ Stopping email monitoring...")
            monitor.stop_monitoring()
            
    finally:
        await monitor.disconnect()

# Run the monitoring system
asyncio.run(run_email_monitoring())
```

This comprehensive examples document demonstrates various real-world usage scenarios for the Outlook MCP Server, from basic operations to advanced integration patterns, error handling, performance optimization, and monitoring systems.