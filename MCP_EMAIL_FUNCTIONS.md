# Outlook MCP Server - Email Functions

This document describes the available email functions in the Outlook MCP Server for LLM assistants like Gemini to understand and use effectively.

## Available Functions

### 1. `list_inbox_emails`
**Purpose**: Retrieve a list of emails from the user's inbox with basic information and body content.

**Parameters**:
- `unread_only` (boolean, optional): If true, only return unread emails. Default: false
- `limit` (integer, optional): Maximum number of emails to return. Default: 50, Max: 100

**Returns**: Array of email objects with:
- `id`: Unique email identifier (required for get_email)
- `subject`: Email subject line
- `sender`: Sender's display name
- `sender_email`: Sender's email address
- `recipients`: Array of recipient email addresses
- `body`: Plain text email body content
- `body_html`: HTML email body content
- `received_time`: When email was received (ISO format)
- `sent_time`: When email was sent (ISO format)
- `is_read`: Boolean indicating if email has been read
- `has_attachments`: Boolean indicating if email has attachments
- `folder_name`: Name of the folder containing the email

**Use Cases**:
- Get recent emails for summarization
- Find emails by scanning subjects and senders
- Check for unread messages
- Get email content for analysis or response

**Example Usage**:
```json
{
  "method": "list_inbox_emails",
  "params": {
    "unread_only": false,
    "limit": 10
  }
}
```

### 2. `get_email`
**Purpose**: Retrieve detailed information for a specific email by its ID.

**Parameters**:
- `email_id` (string, required): The unique identifier of the email (obtained from list_inbox_emails)

**Returns**: Single email object with complete details (same structure as list_inbox_emails items)

**Use Cases**:
- Get full details of a specific email
- Retrieve complete body content for long emails
- Access all metadata for a particular email
- Follow up on emails found through list_inbox_emails

**Example Usage**:
```json
{
  "method": "get_email", 
  "params": {
    "email_id": "00000000DB2820C5F3F8204492F273035529BA6807009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000C8F3FCAAA615D740A192DA27F69F25450000365C4DE00000"
  }
}
```

### 3. `send_email`
**Purpose**: Send a new email through Outlook.

**Parameters**:
- `to` (array of strings, required): Recipient email addresses
- `subject` (string, required): Email subject line
- `body` (string, required): Plain text email body
- `cc` (array of strings, optional): CC recipient email addresses
- `bcc` (array of strings, optional): BCC recipient email addresses
- `body_html` (string, optional): HTML formatted email body (takes precedence over plain text)

**Returns**: 
- `success`: Boolean indicating if email was sent successfully
- `message`: Status message
- `email_id`: ID of the sent email (if successful)

**Use Cases**:
- Send replies to received emails
- Send new emails based on user requests
- Forward information via email
- Send automated responses or notifications

**Example Usage**:
```json
{
  "method": "send_email",
  "params": {
    "to": ["recipient@example.com"],
    "subject": "Meeting Follow-up",
    "body": "Thank you for the meeting today. Here are the action items we discussed...",
    "cc": ["manager@example.com"]
  }
}
```

### 4. `search_emails`
**Purpose**: Search for emails across all folders using keywords or criteria.

**Parameters**:
- `query` (string, required): Search query (searches in subject, body, sender)
- `folder_name` (string, optional): Specific folder to search in (default: all folders)
- `unread_only` (boolean, optional): Only search unread emails. Default: false
- `limit` (integer, optional): Maximum number of results. Default: 50

**Returns**: Array of email objects (same structure as list_inbox_emails)

**Use Cases**:
- Find emails containing specific keywords
- Search for emails from particular senders
- Locate emails about specific topics or projects
- Find emails in specific folders

**Example Usage**:
```json
{
  "method": "search_emails",
  "params": {
    "query": "project status meeting",
    "unread_only": false,
    "limit": 20
  }
}
```

## Common Workflow Patterns

### Pattern 1: Email Summarization
1. Use `list_inbox_emails` with appropriate limit
2. Process the returned emails to extract key information
3. Create summary based on subjects, senders, and body content

### Pattern 2: Email Response
1. Use `list_inbox_emails` or `search_emails` to find relevant emails
2. Use `get_email` to get full details if needed
3. Use `send_email` to send response with appropriate recipients and content

### Pattern 3: Email Search and Analysis
1. Use `search_emails` with relevant keywords
2. Use `get_email` for detailed analysis of specific results
3. Process and analyze the content as needed

### Pattern 4: Unread Email Processing
1. Use `list_inbox_emails` with `unread_only: true`
2. Process each unread email
3. Use `get_email` for detailed content if needed
4. Use `send_email` for responses if required

## Important Notes

### Email IDs
- Email IDs are long hexadecimal strings (140+ characters)
- Always use the exact ID returned from list_inbox_emails or search_emails
- IDs are unique and persistent for each email

### Body Content
- Both `body` (plain text) and `body_html` (HTML) are available
- Some emails may have empty body content due to formatting or permissions
- Current success rate for body extraction: ~84%

### Error Handling
- Functions return appropriate error messages for invalid parameters
- Email not found errors occur if using invalid or expired email IDs
- Permission errors may occur for restricted emails

### Performance
- `list_inbox_emails` is optimized for bulk email retrieval
- `get_email` is best for detailed single email access
- Use appropriate limits to avoid timeouts on large mailboxes

## Example Scenarios

### Scenario 1: "Summarize my recent emails"
```json
{
  "method": "list_inbox_emails",
  "params": {
    "limit": 20
  }
}
```

### Scenario 2: "Find emails about project X"
```json
{
  "method": "search_emails", 
  "params": {
    "query": "project X",
    "limit": 10
  }
}
```

### Scenario 3: "Reply to the latest email from John"
1. Search: `{"method": "search_emails", "params": {"query": "from:john@example.com", "limit": 1}}`
2. Get details: `{"method": "get_email", "params": {"email_id": "..."}}`
3. Send reply: `{"method": "send_email", "params": {"to": ["john@example.com"], "subject": "RE: ...", "body": "..."}}`

### Scenario 4: "Check unread emails"
```json
{
  "method": "list_inbox_emails",
  "params": {
    "unread_only": true,
    "limit": 50
  }
}
```

This MCP server provides comprehensive email management capabilities through Microsoft Outlook, enabling LLM assistants to help users with email-related tasks efficiently and effectively.