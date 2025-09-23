# Outlook MCP Server API Documentation

## Overview

The Outlook MCP Server provides programmatic access to Microsoft Outlook email functionality through the Model Context Protocol (MCP). This document provides comprehensive documentation for all available functions, parameters, response formats, and usage examples.

## Table of Contents

- [Authentication & Connection](#authentication--connection)
- [Available Methods](#available-methods)
  - [list_inbox_emails](#list_inbox_emails)
  - [list_emails](#list_emails)
  - [get_email](#get_email)
  - [search_emails](#search_emails)
  - [send_email](#send_email)
  - [get_folders](#get_folders)
- [Data Models](#data-models)
- [Error Handling](#error-handling)
- [Rate Limiting](#rate-limiting)
- [Performance Considerations](#performance-considerations)

## Authentication & Connection

The Outlook MCP Server connects to a locally installed Microsoft Outlook application using COM (Component Object Model) interface. No additional authentication is required beyond having Outlook installed and configured on the Windows system.

### Prerequisites

- Microsoft Outlook installed and configured
- Windows operating system (required for COM interface)
- Outlook application running or accessible
- Appropriate permissions to access Outlook data

## Available Methods

### list_inbox_emails

Lists emails from the default inbox folder with filtering options. This is a simplified method that automatically finds and accesses the inbox folder.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `unread_only` | boolean | No | false | Filter to show only unread emails |
| `limit` | integer | No | 50 | Maximum number of emails to return (1-1000) |

#### Request Example

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "list_inbox_emails",
  "params": {
    "unread_only": true,
    "limit": 10
  }
}
```

#### Response Format

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "result": {
    "emails": [
      {
        "id": "AAMkADExMzJmYWE...",
        "subject": "Project Update - Q4 Planning",
        "sender": "John Doe",
        "sender_email": "john.doe@company.com",
        "recipients": ["team@company.com"],
        "received_time": "2024-01-15T10:30:00Z",
        "sent_time": "2024-01-15T10:25:00Z",
        "is_read": false,
        "has_attachments": true,
        "importance": "Normal",
        "folder": "Inbox",
        "size": 15420,
        "body_preview": "Hi team, I wanted to share the latest updates..."
      }
    ]
  }
}
```

#### Usage Examples

**List latest 5 emails from inbox:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "list_inbox_emails",
  "params": {
    "limit": 5
  }
}
```

**List only unread emails from inbox:**
```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "method": "list_inbox_emails",
  "params": {
    "unread_only": true
  }
}
```

---

### list_emails

Lists emails from a specific folder by folder ID with filtering options. Use `get_folders` to obtain folder IDs.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folder_id` | string | Yes | - | Folder ID to list emails from (use get_folders to see available folder IDs) |
| `unread_only` | boolean | No | false | Filter to show only unread emails |
| `limit` | integer | No | 50 | Maximum number of emails to return (1-1000) |

#### Request Example

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "list_emails",
  "params": {
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "unread_only": true,
    "limit": 10
  }
}
```

#### Response Format

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "result": [
    {
      "id": "AAMkADExMzJmYWE...",
      "subject": "Project Update - Q4 Planning",
      "sender": "John Doe",
      "sender_email": "john.doe@company.com",
      "recipients": ["team@company.com"],
      "cc_recipients": ["manager@company.com"],
      "bcc_recipients": [],
      "received_time": "2024-01-15T10:30:00Z",
      "sent_time": "2024-01-15T10:25:00Z",
      "is_read": false,
      "has_attachments": true,
      "importance": "Normal",
      "folder_name": "Inbox",
      "size": 15420,
      "accessible": true,
      "has_body": true,
      "body_preview": "Hi team, I wanted to share the latest updates on our Q4 planning initiative...",
      "attachment_count": 2
    }
  ]
}
```

#### Usage Examples

**List emails from specific folder:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "list_emails",
  "params": {
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000"
  }
}
```

**List unread emails from specific folder:**
```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "method": "list_emails",
  "params": {
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "unread_only": true,
    "limit": 25
  }
}
```

**List emails from Sent Items folder:**
```json
{
  "jsonrpc": "2.0",
  "id": "3",
  "method": "list_emails",
  "params": {
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A91111",
    "limit": 100
  }
}
```

### get_email

Retrieves detailed information for a specific email by its unique identifier.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `email_id` | string | Yes | - | Unique identifier of the email to retrieve |

#### Request Example

```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "method": "get_email",
  "params": {
    "email_id": "AAMkADExMzJmYWE..."
  }
}
```

#### Response Format

```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "result": {
    "id": "AAMkADExMzJmYWE...",
    "subject": "Project Update - Q4 Planning",
    "sender": "John Doe",
    "sender_email": "john.doe@company.com",
    "recipients": ["team@company.com"],
    "cc_recipients": ["manager@company.com"],
    "bcc_recipients": [],
    "body": "Hi team,\n\nI wanted to share the latest updates on our Q4 planning initiative...",
    "body_html": "<html><body><p>Hi team,</p><p>I wanted to share the latest updates...</p></body></html>",
    "received_time": "2024-01-15T10:30:00Z",
    "sent_time": "2024-01-15T10:25:00Z",
    "is_read": false,
    "has_attachments": true,
    "importance": "High",
    "folder_name": "Inbox",
    "size": 15420,
    "attachments": [
      {
        "name": "Q4_Planning_Document.pdf",
        "size": 245760,
        "type": "application/pdf"
      },
      {
        "name": "Budget_Spreadsheet.xlsx",
        "size": 89432,
        "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      }
    ],
    "accessible": true,
    "has_body": true,
    "body_preview": "Hi team, I wanted to share the latest updates on our Q4 planning initiative...",
    "attachment_count": 2
  }
}
```

#### Usage Examples

**Get specific email by ID:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "get_email",
  "params": {
    "email_id": "AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK"
  }
}
```

### search_emails

Searches emails based on user-defined queries across folders or within specific folders.

#### Parameters

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `query` | string | Yes | - | Search query string (max 1000 characters) |
| `folder_id` | string | No | null | Folder ID to limit search to. If not specified, searches all accessible folders |
| `limit` | integer | No | 50 | Maximum number of results to return (1-1000) |

#### Request Example

```json
{
  "jsonrpc": "2.0",
  "id": "3",
  "method": "search_emails",
  "params": {
    "query": "project update",
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "limit": 20
  }
}
```

#### Response Format

```json
{
  "jsonrpc": "2.0",
  "id": "3",
  "result": [
    {
      "id": "AAMkADExMzJmYWE...",
      "subject": "Project Update - Q4 Planning",
      "sender": "John Doe",
      "sender_email": "john.doe@company.com",
      "recipients": ["team@company.com"],
      "cc_recipients": [],
      "bcc_recipients": [],
      "received_time": "2024-01-15T10:30:00Z",
      "sent_time": "2024-01-15T10:25:00Z",
      "is_read": true,
      "has_attachments": false,
      "importance": "Normal",
      "folder_name": "Inbox",
      "size": 8420,
      "accessible": true,
      "has_body": true,
      "body_preview": "Hi team, here's the latest project update for Q4 planning...",
      "attachment_count": 0
    }
  ]
}
```

#### Search Query Syntax

The search functionality supports various query formats:

- **Simple text search**: `"project update"`
- **Sender search**: `"from:john.doe@company.com"`
- **Subject search**: `"subject:meeting"`
- **Date range**: `"received:2024-01-01..2024-01-31"`
- **Boolean operators**: `"project AND update"`
- **Phrase search**: `"exact phrase"`

#### Usage Examples

**Search all folders:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "search_emails",
  "params": {
    "query": "meeting agenda"
  }
}
```

**Search specific folder:**
```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "method": "search_emails",
  "params": {
    "query": "from:manager@company.com",
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "limit": 10
  }
}
```

**Complex search query:**
```json
{
  "jsonrpc": "2.0",
  "id": "3",
  "method": "search_emails",
  "params": {
    "query": "subject:project AND received:2024-01-01..2024-01-31",
    "limit": 50
  }
}
```

### get_folders

Lists all available email folders in Outlook with their hierarchy and metadata.

#### Parameters

None - this method takes no parameters.

#### Request Example

```json
{
  "jsonrpc": "2.0",
  "id": "4",
  "method": "get_folders",
  "params": {}
}
```

#### Response Format

```json
{
  "jsonrpc": "2.0",
  "id": "4",
  "result": [
    {
      "id": "AAMkADExMzJmYWE...",
      "name": "Inbox",
      "full_path": "Mailbox - user@company.com/Inbox",
      "item_count": 245,
      "unread_count": 12,
      "parent_folder": "Mailbox - user@company.com",
      "folder_type": "Mail"
    },
    {
      "id": "AAMkADExMzJmYWF...",
      "name": "Sent Items",
      "full_path": "Mailbox - user@company.com/Sent Items",
      "item_count": 1024,
      "unread_count": 0,
      "parent_folder": "Mailbox - user@company.com",
      "folder_type": "Mail"
    },
    {
      "id": "AAMkADExMzJmYWG...",
      "name": "Projects",
      "full_path": "Mailbox - user@company.com/Inbox/Projects",
      "item_count": 89,
      "unread_count": 5,
      "parent_folder": "Inbox",
      "folder_type": "Mail"
    }
  ]
}
```

#### Usage Examples

**Get all folders:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "get_folders",
  "params": {}
}
```

## Data Models

### Email Data Model

| Field | Type | Description |
|-------|------|-------------|
| `id` | string | Unique email identifier |
| `subject` | string | Email subject line |
| `sender` | string | Sender display name |
| `sender_email` | string | Sender email address |
| `recipients` | array[string] | List of recipient email addresses |
| `cc_recipients` | array[string] | List of CC recipient email addresses |
| `bcc_recipients` | array[string] | List of BCC recipient email addresses |
| `body` | string | Plain text email body |
| `body_html` | string | HTML email body |
| `received_time` | string | ISO 8601 timestamp when email was received |
| `sent_time` | string | ISO 8601 timestamp when email was sent |
| `is_read` | boolean | Whether the email has been read |
| `has_attachments` | boolean | Whether the email has attachments |
| `importance` | string | Email importance level (Low, Normal, High) |
| `folder_name` | string | Name of the folder containing the email |
| `size` | integer | Email size in bytes |
| `attachments` | array[object] | List of attachment objects (only in get_email) |
| `accessible` | boolean | Whether the email is accessible |
| `has_body` | boolean | Whether the email has body content |
| `body_preview` | string | Preview of email body (truncated) |
| `attachment_count` | integer | Number of attachments |

### Folder Data Model

| Field | Type | Description |
|-------|------|-------------|
| `id` | string | Unique folder identifier |
| `name` | string | Folder display name |
| `full_path` | string | Complete folder path |
| `item_count` | integer | Total number of items in folder |
| `unread_count` | integer | Number of unread items in folder |
| `parent_folder` | string | Parent folder name |
| `folder_type` | string | Type of folder (Mail, Calendar, Contacts, etc.) |

### Attachment Data Model

| Field | Type | Description |
|-------|------|-------------|
| `name` | string | Attachment filename |
| `size` | integer | Attachment size in bytes |
| `type` | string | MIME type of the attachment |

## Error Handling

All errors follow the MCP protocol error format:

```json
{
  "jsonrpc": "2.0",
  "id": "request_id",
  "error": {
    "code": -32000,
    "message": "Error description",
    "data": {
      "type": "ErrorType",
      "details": "Additional error information"
    }
  }
}
```

### Error Types

| Error Type | Code | Description |
|------------|------|-------------|
| `ValidationError` | -32602 | Invalid parameters provided |
| `OutlookConnectionError` | -32001 | Cannot connect to Outlook |
| `EmailNotFoundError` | -32002 | Specified email not found |
| `FolderNotFoundError` | -32003 | Specified folder not found |
| `PermissionError` | -32004 | Access denied to resource |
| `SearchError` | -32005 | Search operation failed |
| `TimeoutError` | -32006 | Operation timed out |
| `RateLimitError` | -32007 | Rate limit exceeded |

### Common Error Scenarios

**Invalid Email ID:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "error": {
    "code": -32602,
    "message": "Invalid email ID format",
    "data": {
      "type": "ValidationError",
      "details": "Email ID must be a valid Outlook entry ID"
    }
  }
}
```

**Folder Not Found:**
```json
{
  "jsonrpc": "2.0",
  "id": "2",
  "error": {
    "code": -32003,
    "message": "Folder not found: 'NonExistentFolder'",
    "data": {
      "type": "FolderNotFoundError",
      "details": "The specified folder does not exist or is not accessible"
    }
  }
}
```

**Outlook Connection Error:**
```json
{
  "jsonrpc": "2.0",
  "id": "3",
  "error": {
    "code": -32001,
    "message": "Cannot connect to Outlook",
    "data": {
      "type": "OutlookConnectionError",
      "details": "Outlook application is not running or not accessible"
    }
  }
}
```

## Rate Limiting

The server implements rate limiting to prevent abuse and ensure stable performance:

- **Default limits**: 100 requests per minute per client
- **Burst allowance**: Up to 10 requests in a 10-second window
- **Rate limit headers**: Included in error responses when limits are exceeded

**Rate Limit Exceeded Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "error": {
    "code": -32007,
    "message": "Rate limit exceeded",
    "data": {
      "type": "RateLimitError",
      "details": "Too many requests. Please wait before making additional requests.",
      "retry_after": 60
    }
  }
}
```

## Performance Considerations

### Pagination

- Use the `limit` parameter to control response size
- Maximum limit is 1000 items per request
- For large datasets, make multiple requests with appropriate limits

### Caching

- The server implements intelligent caching for frequently accessed data
- Email content is cached for improved performance
- Folder structures are cached and updated periodically

### Best Practices

1. **Use appropriate limits**: Don't request more data than needed
2. **Cache responses**: Cache responses on the client side when appropriate
3. **Handle errors gracefully**: Implement proper error handling and retry logic
4. **Use specific folders**: Specify folder names when possible to improve performance
5. **Optimize search queries**: Use specific search terms to reduce processing time

### Performance Monitoring

The server provides performance statistics through internal monitoring:

- Request processing times
- Cache hit/miss ratios
- Memory usage statistics
- Connection pool utilization

These statistics are logged and can be used for performance optimization.