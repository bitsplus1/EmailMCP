# Outlook MCP Server - HTTP API Examples

This document provides comprehensive examples for using the Outlook MCP Server in HTTP mode, which is the recommended approach for testing and web application integration.

## Getting Started

### Starting the HTTP Server

```bash
# Start with the recommended configuration
python main.py http --config docker_config.json

# Start with custom host and port
python main.py http --config docker_config.json --host 0.0.0.0 --port 8080

# Start with debug logging
python main.py http --config docker_config.json --log-level DEBUG
```

The server will be available at `http://localhost:8080` by default.

### Configuration File

The `docker_config.json` file contains the HTTP server configuration:

```json
{
  "server_host": "0.0.0.0",
  "server_port": 8080,
  "log_level": "INFO",
  "enable_console_output": true
}
```

## Basic HTTP API Examples

All requests are sent to the `/mcp` endpoint using POST method with JSON-RPC 2.0 format.

### 1. List Inbox Emails

Get recent emails from your inbox:

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "1",
    "method": "list_inbox_emails",
    "params": {
      "limit": 10
    }
  }'
```

**Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "result": [
    {
      "id": "AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK...",
      "subject": "Project Update Meeting",
      "sender": "John Smith",
      "sender_email": "john.smith@company.com",
      "recipients": ["user@company.com"],
      "body": "Hi, let's schedule our project update meeting for next week...",
      "received_time": "2024-01-15T14:30:00Z",
      "is_read": false,
      "has_attachments": true,
      "folder_name": "Inbox"
    }
  ]
}
```

### 2. Get Unread Emails Only

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "2",
    "method": "list_inbox_emails",
    "params": {
      "unread_only": true,
      "limit": 20
    }
  }'
```

### 3. Get Specific Email Details

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "3",
    "method": "get_email",
    "params": {
      "email_id": "AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK..."
    }
  }'
```

### 4. Search Emails

Search for emails containing specific keywords:

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "4",
    "method": "search_emails",
    "params": {
      "query": "project meeting",
      "limit": 15
    }
  }'
```

### 5. Send Email

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "5",
    "method": "send_email",
    "params": {
      "to": ["recipient@example.com"],
      "subject": "Test Email from MCP Server",
      "body": "This is a test email sent via the Outlook MCP Server HTTP API."
    }
  }'
```

**Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "5",
  "result": {
    "success": true,
    "message": "Email sent successfully",
    "email_id": "AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK..."
  }
}
```

## Advanced Examples

### 6. Send Email with CC and HTML Body

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "6",
    "method": "send_email",
    "params": {
      "to": ["recipient@example.com"],
      "cc": ["manager@example.com"],
      "subject": "Project Status Update",
      "body": "Please see the HTML version for formatting.",
      "body_html": "<h2>Project Status Update</h2><p>The project is <strong>on track</strong> for completion.</p><ul><li>Phase 1: Complete</li><li>Phase 2: In Progress</li></ul>"
    }
  }'
```

### 7. Search with Advanced Query

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "7",
    "method": "search_emails",
    "params": {
      "query": "from:john@company.com AND subject:project",
      "limit": 25
    }
  }'
```

### 8. Get Folder List

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "8",
    "method": "get_folders",
    "params": {}
  }'
```

## Python Integration Examples

### Simple Python Client

```python
import requests
import json

class OutlookMCPClient:
    def __init__(self, base_url="http://localhost:8080"):
        self.base_url = base_url
        self.session = requests.Session()
        self.session.headers.update({"Content-Type": "application/json"})
    
    def _make_request(self, method, params=None):
        """Make a JSON-RPC request to the MCP server."""
        payload = {
            "jsonrpc": "2.0",
            "id": method,
            "method": method,
            "params": params or {}
        }
        
        response = self.session.post(f"{self.base_url}/mcp", json=payload)
        response.raise_for_status()
        
        result = response.json()
        
        if "error" in result:
            raise Exception(f"MCP Error: {result['error']['message']}")
        
        return result.get("result")
    
    def list_inbox_emails(self, unread_only=False, limit=50):
        """List emails from inbox."""
        return self._make_request("list_inbox_emails", {
            "unread_only": unread_only,
            "limit": limit
        })
    
    def get_email(self, email_id):
        """Get specific email by ID."""
        return self._make_request("get_email", {
            "email_id": email_id
        })
    
    def search_emails(self, query, limit=50):
        """Search emails by query."""
        return self._make_request("search_emails", {
            "query": query,
            "limit": limit
        })
    
    def send_email(self, to, subject, body, cc=None, bcc=None, body_html=None):
        """Send an email."""
        params = {
            "to": to if isinstance(to, list) else [to],
            "subject": subject,
            "body": body
        }
        
        if cc:
            params["cc"] = cc if isinstance(cc, list) else [cc]
        if bcc:
            params["bcc"] = bcc if isinstance(bcc, list) else [bcc]
        if body_html:
            params["body_html"] = body_html
        
        return self._make_request("send_email", params)

# Usage example
if __name__ == "__main__":
    client = OutlookMCPClient()
    
    try:
        # List recent emails
        print("Recent emails:")
        emails = client.list_inbox_emails(limit=5)
        for email in emails:
            print(f"- {email['subject']} (from: {email['sender']})")
        
        # Search for project emails
        print("\nProject emails:")
        project_emails = client.search_emails("project", limit=3)
        for email in project_emails:
            print(f"- {email['subject']} ({email['received_time']})")
        
        # Send a test email (uncomment to actually send)
        # result = client.send_email(
        #     to="test@example.com",
        #     subject="Test from Python",
        #     body="This is a test email from the Python client."
        # )
        # print(f"Email sent: {result}")
        
    except Exception as e:
        print(f"Error: {e}")
```

### Async Python Client

```python
import aiohttp
import asyncio
import json

class AsyncOutlookMCPClient:
    def __init__(self, base_url="http://localhost:8080"):
        self.base_url = base_url
    
    async def _make_request(self, session, method, params=None):
        """Make an async JSON-RPC request."""
        payload = {
            "jsonrpc": "2.0",
            "id": method,
            "method": method,
            "params": params or {}
        }
        
        async with session.post(f"{self.base_url}/mcp", json=payload) as response:
            response.raise_for_status()
            result = await response.json()
            
            if "error" in result:
                raise Exception(f"MCP Error: {result['error']['message']}")
            
            return result.get("result")
    
    async def list_inbox_emails(self, session, unread_only=False, limit=50):
        """List emails from inbox."""
        return await self._make_request(session, "list_inbox_emails", {
            "unread_only": unread_only,
            "limit": limit
        })
    
    async def search_emails(self, session, query, limit=50):
        """Search emails by query."""
        return await self._make_request(session, "search_emails", {
            "query": query,
            "limit": limit
        })

# Usage example
async def main():
    client = AsyncOutlookMCPClient()
    
    async with aiohttp.ClientSession() as session:
        try:
            # Concurrent requests
            tasks = [
                client.list_inbox_emails(session, limit=5),
                client.search_emails(session, "meeting", limit=3),
                client.search_emails(session, "project", limit=3)
            ]
            
            results = await asyncio.gather(*tasks)
            
            print(f"Inbox emails: {len(results[0])}")
            print(f"Meeting emails: {len(results[1])}")
            print(f"Project emails: {len(results[2])}")
            
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    asyncio.run(main())
```

## JavaScript/Node.js Integration

### Simple Node.js Client

```javascript
const axios = require('axios');

class OutlookMCPClient {
    constructor(baseUrl = 'http://localhost:8080') {
        this.baseUrl = baseUrl;
        this.client = axios.create({
            baseURL: baseUrl,
            headers: {
                'Content-Type': 'application/json'
            }
        });
    }

    async makeRequest(method, params = {}) {
        const payload = {
            jsonrpc: '2.0',
            id: method,
            method: method,
            params: params
        };

        try {
            const response = await this.client.post('/mcp', payload);
            
            if (response.data.error) {
                throw new Error(`MCP Error: ${response.data.error.message}`);
            }
            
            return response.data.result;
        } catch (error) {
            if (error.response) {
                throw new Error(`HTTP Error: ${error.response.status} - ${error.response.statusText}`);
            }
            throw error;
        }
    }

    async listInboxEmails(unreadOnly = false, limit = 50) {
        return await this.makeRequest('list_inbox_emails', {
            unread_only: unreadOnly,
            limit: limit
        });
    }

    async getEmail(emailId) {
        return await this.makeRequest('get_email', {
            email_id: emailId
        });
    }

    async searchEmails(query, limit = 50) {
        return await this.makeRequest('search_emails', {
            query: query,
            limit: limit
        });
    }

    async sendEmail(to, subject, body, options = {}) {
        const params = {
            to: Array.isArray(to) ? to : [to],
            subject: subject,
            body: body,
            ...options
        };

        return await this.makeRequest('send_email', params);
    }
}

// Usage example
async function main() {
    const client = new OutlookMCPClient();

    try {
        // List recent emails
        console.log('Recent emails:');
        const emails = await client.listInboxEmails(false, 5);
        emails.forEach(email => {
            console.log(`- ${email.subject} (from: ${email.sender})`);
        });

        // Search for project emails
        console.log('\nProject emails:');
        const projectEmails = await client.searchEmails('project', 3);
        projectEmails.forEach(email => {
            console.log(`- ${email.subject} (${email.received_time})`);
        });

        // Send email (uncomment to actually send)
        // const result = await client.sendEmail(
        //     'test@example.com',
        //     'Test from Node.js',
        //     'This is a test email from the Node.js client.'
        // );
        // console.log('Email sent:', result);

    } catch (error) {
        console.error('Error:', error.message);
    }
}

main();
```

## Testing and Debugging

### Health Check

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "health",
    "method": "get_folders",
    "params": {}
  }'
```

### Error Handling Example

```bash
# This will return an error for invalid email ID
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "error_test",
    "method": "get_email",
    "params": {
      "email_id": "invalid_id"
    }
  }'
```

**Error Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "error_test",
  "error": {
    "code": -32602,
    "message": "Email not found",
    "data": {
      "type": "EmailNotFoundError",
      "email_id": "invalid_id"
    }
  }
}
```

## Performance Tips

1. **Use appropriate limits**: Don't request more emails than needed
2. **Cache results**: Store frequently accessed data locally
3. **Batch requests**: Use concurrent requests when possible
4. **Handle errors gracefully**: Implement retry logic for transient failures
5. **Monitor server logs**: Check logs for performance insights

## Troubleshooting

### Common Issues

1. **Connection Refused**: Make sure the server is running on the correct port
2. **Outlook Not Found**: Ensure Outlook is installed and accessible
3. **Permission Denied**: Run the server with appropriate permissions
4. **Timeout Errors**: Increase timeout values in configuration

### Debug Mode

Start the server with debug logging:

```bash
python main.py http --config docker_config.json --log-level DEBUG
```

This provides detailed information about requests, responses, and internal operations.