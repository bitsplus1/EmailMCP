# Outlook MCP Server - Email Functions for LLMs

This document provides comprehensive function descriptions for LLMs (like Gemini) to understand and effectively use the Outlook MCP Server email functions.

## Overview

The Outlook MCP Server provides four core email functions that enable LLMs to interact with Microsoft Outlook for email management tasks. These functions are designed to handle common email workflows including reading, searching, retrieving, and sending emails.

## Server Setup

The recommended way to run the server for testing and integration is using HTTP mode:

```bash
# Start the HTTP server
python main.py http --config docker_config.json

# Server will be available at http://localhost:8080
# All requests are sent to the /mcp endpoint using POST method
```

## Function Descriptions

### 1. `list_inbox_emails`

**Purpose**: Retrieve a list of emails from the user's inbox with complete email content and metadata.

**When to use**:
- User asks to "check my emails" or "show me recent emails"
- Need to summarize recent email activity
- Looking for emails without specific search criteria
- Want to see unread emails only

**Parameters**:
```json
{
  "unread_only": false,  // boolean, optional - Set to true to only get unread emails
  "limit": 50           // integer, optional - Max emails to return (1-100, default: 50)
}
```

**Returns**: Array of email objects with complete information:
```json
[
  {
    "id": "00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE00000",
    "subject": "Meeting Tomorrow",
    "sender": "John Smith",
    "sender_email": "john.smith@company.com",
    "recipients": ["user@company.com"],
    "body": "Hi, just confirming our meeting tomorrow at 2 PM...",
    "body_html": "<p>Hi, just confirming our meeting tomorrow at 2 PM...</p>",
    "received_time": "2024-01-15T14:30:00Z",
    "sent_time": "2024-01-15T14:29:45Z",
    "is_read": false,
    "has_attachments": true,
    "folder_name": "Inbox"
  }
]
```

**Example Usage Scenarios**:
- "Show me my latest 10 emails" → `{"unread_only": false, "limit": 10}`
- "Check for unread emails" → `{"unread_only": true, "limit": 50}`
- "Summarize today's emails" → `{"unread_only": false, "limit": 20}`

**HTTP Request Example**:
```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "1",
    "method": "list_inbox_emails",
    "params": {
      "unread_only": false,
      "limit": 10
    }
  }'
```

### 2. `get_email`

**Purpose**: Retrieve complete details for a specific email using its unique ID.

**When to use**:
- Need full details of a specific email found through `list_inbox_emails` or `search_emails`
- User asks about a specific email by reference
- Need to read the complete content of a particular email
- Following up on emails from search results

**Parameters**:
```json
{
  "email_id": "00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE00000"
}
```

**Returns**: Single email object with same structure as `list_inbox_emails` items.

**Important Notes**:
- Email IDs are long hexadecimal strings (140+ characters)
- Always use the exact ID returned from other functions
- IDs are unique and persistent for each email

**Example Usage Scenarios**:
- After finding emails with `search_emails`, use `get_email` to read full content
- User says "read me the full email from John about the project"
- Need complete email details for response generation

**HTTP Request Example**:
```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "2",
    "method": "get_email",
    "params": {
      "email_id": "00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE00000"
    }
  }'
```

### 3. `send_email`

**Purpose**: Send new emails through Microsoft Outlook.

**When to use**:
- User asks to send an email or reply to someone
- Need to send automated responses or notifications
- Forwarding information via email
- Creating and sending new communications

**Parameters**:
```json
{
  "to": ["recipient@example.com", "another@example.com"],     // required - array of recipient emails
  "subject": "Meeting Follow-up",                             // required - email subject
  "body": "Thank you for the meeting today...",              // required - plain text body
  "cc": ["manager@example.com"],                             // optional - CC recipients
  "bcc": ["archive@example.com"],                            // optional - BCC recipients  
  "body_html": "<p>Thank you for the meeting today...</p>"   // optional - HTML body (overrides plain text)
}
```

**Returns**:
```json
{
  "success": true,
  "message": "Email sent successfully",
  "email_id": "00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE00000"
}
```

**Example Usage Scenarios**:
- "Send an email to john@company.com about the project update"
- "Reply to Sarah's email about the meeting"
- "Send a follow-up email to the team"

**HTTP Request Example**:
```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "3",
    "method": "send_email",
    "params": {
      "to": ["john@company.com"],
      "subject": "Project Update",
      "body": "Hi John, here is the latest update on our project..."
    }
  }'
```

### 4. `search_emails`

**Purpose**: Search for emails across all folders using keywords, sender information, or other criteria.

**When to use**:
- User asks to find emails about specific topics
- Looking for emails from particular people
- Need to locate emails containing specific keywords
- Searching for emails in specific folders

**Parameters**:
```json
{
  "query": "project status meeting",        // required - search terms (searches subject, body, sender)
  "folder_name": "Inbox",                  // optional - specific folder to search (default: all folders)
  "unread_only": false,                    // optional - only search unread emails
  "limit": 50                              // optional - max results to return (default: 50)
}
```

**Search Query Tips**:
- Use keywords that might appear in subject or body: "project", "meeting", "invoice"
- Search for sender names: "John Smith" or "john@company.com"
- Use multiple keywords: "project status update"
- Folder names can be in different languages (English, Chinese, Japanese, etc.)

**Returns**: Array of email objects (same structure as `list_inbox_emails`)

**Example Usage Scenarios**:
- "Find emails about the quarterly report" → `{"query": "quarterly report"}`
- "Show me emails from John in the last week" → `{"query": "John Smith"}`
- "Search for unread emails about meetings" → `{"query": "meeting", "unread_only": true}`

**HTTP Request Example**:
```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "4",
    "method": "search_emails",
    "params": {
      "query": "quarterly report",
      "limit": 20
    }
  }'
```

## Common Workflow Patterns

### Pattern 1: Email Summarization
```
1. Use list_inbox_emails(limit=20) to get recent emails
2. Process subjects, senders, and body content
3. Create summary of key emails and topics
```

### Pattern 2: Find and Reply to Specific Email
```
1. Use search_emails(query="sender name or topic") to find relevant emails
2. Use get_email(email_id) to get full details if needed
3. Use send_email() to send response with appropriate recipients
```

### Pattern 3: Unread Email Processing
```
1. Use list_inbox_emails(unread_only=true) to get unread emails
2. Process each email for important information
3. Use send_email() for any required responses
```

### Pattern 4: Topic-Based Email Search
```
1. Use search_emails(query="topic keywords") to find related emails
2. Use get_email() for detailed analysis of specific results
3. Summarize findings or take action based on content
```

## Response Handling

### Success Responses
All functions return structured data that can be directly processed. Email objects contain both plain text (`body`) and HTML (`body_html`) versions of content.

### Error Handling
Functions may return errors for:
- Invalid email IDs (email not found)
- Permission issues (restricted emails)
- Connection problems (Outlook not available)
- Invalid parameters (malformed requests)

### Performance Considerations
- `list_inbox_emails` is optimized for bulk retrieval (use appropriate limits)
- `get_email` is best for single email details
- `search_emails` performance depends on query complexity and mailbox size
- Current body extraction success rate: ~84%

## Practical Examples

### Example 1: "Summarize my recent emails"
```json
{
  "method": "list_inbox_emails",
  "params": {
    "limit": 15
  }
}
```

### Example 2: "Find emails about the project from last week"
```json
{
  "method": "search_emails",
  "params": {
    "query": "project",
    "limit": 20
  }
}
```

### Example 3: "Send a thank you email to the team"
```json
{
  "method": "send_email",
  "params": {
    "to": ["team@company.com"],
    "subject": "Thank You",
    "body": "Thank you all for your hard work on the project. Great job!"
  }
}
```

### Example 4: "Check for unread emails from my manager"
```json
{
  "method": "search_emails",
  "params": {
    "query": "manager@company.com",
    "unread_only": true
  }
}
```

## Best Practices for LLMs

1. **Always use exact email IDs** returned from other functions - never modify or truncate them
2. **Use appropriate limits** to avoid timeouts - start with smaller limits for testing
3. **Handle empty results gracefully** - not all searches will return results
4. **Process both body and body_html** - some emails may have content in only one format
5. **Use descriptive search queries** - include relevant keywords that would appear in emails
6. **Consider folder names in different languages** - Outlook may use localized folder names
7. **Chain functions logically** - use search/list functions first, then get_email for details, then send_email for responses

## HTTP Server Usage

The recommended way to use this MCP server is through HTTP mode, which provides a simple REST-like interface:

### Starting the Server
```bash
python main.py http --config docker_config.json
```

### Making Requests
All requests are sent to `http://localhost:8080/mcp` using POST method with JSON-RPC 2.0 format:

```bash
curl -X POST http://localhost:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": "unique_request_id",
    "method": "function_name",
    "params": {
      "parameter1": "value1",
      "parameter2": "value2"
    }
  }'
```

### Integration with Programming Languages

**Python Example:**
```python
import requests

def call_mcp_function(method, params):
    response = requests.post('http://localhost:8080/mcp', json={
        "jsonrpc": "2.0",
        "id": method,
        "method": method,
        "params": params
    })
    return response.json()['result']

# Usage
emails = call_mcp_function('list_inbox_emails', {'limit': 10})
```

**JavaScript Example:**
```javascript
async function callMCPFunction(method, params) {
    const response = await fetch('http://localhost:8080/mcp', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            jsonrpc: '2.0',
            id: method,
            method: method,
            params: params
        })
    });
    const result = await response.json();
    return result.result;
}

// Usage
const emails = await callMCPFunction('list_inbox_emails', {limit: 10});
```

This MCP server enables comprehensive email management through Microsoft Outlook, allowing LLMs to help users with email-related tasks efficiently and effectively.