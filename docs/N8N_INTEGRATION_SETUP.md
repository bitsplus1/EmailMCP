# n8n Integration Setup Guide

This comprehensive guide provides step-by-step instructions for integrating the Outlook MCP Server with your n8n instance. You can choose between two integration methods based on your security requirements and deployment scenario.

## Integration Methods Overview

The Outlook MCP Server supports two communication modes:

### Method A: stdio Mode (Recommended for Local Use)
- ✅ **Maximum Security** - No network exposure, local process execution only
- ✅ **No Network Ports** - Communication through standard input/output
- ✅ **Simple Setup** - No firewall or port configuration needed
- ✅ **Same Machine Only** - n8n and MCP server must be on the same computer
- 🔧 **Uses**: n8n Execute Command node

### Method B: HTTP Mode (For Remote Access)
- 🌐 **Remote Access** - n8n can be on a different computer than MCP server
- 🔌 **Network Communication** - Uses HTTP requests on configurable port (default: 8080)
- ⚠️ **Security Consideration** - Requires network configuration and security measures
- 🔧 **Uses**: n8n HTTP Request node

### Which Method Should You Choose?

| Aspect | Method A (stdio) | Method B (HTTP) |
|--------|------------------|-----------------|
| **Security** | ✅ Maximum (no network exposure) | ⚠️ Requires network security |
| **Setup Complexity** | ✅ Simple (no server management) | ⚠️ Moderate (server management) |
| **Performance** | ✅ Fast (direct process) | ⚠️ Network overhead |
| **Remote Access** | ❌ Same machine only | ✅ Remote access capable |
| **Resource Usage** | ⚠️ New process per request | ✅ Persistent server |
| **Debugging** | ⚠️ Limited (process output) | ✅ Server logs and endpoints |
| **Scalability** | ⚠️ Process overhead | ✅ Handles concurrent requests |

**Recommended Choice:**
- **Use Method A (stdio)** if n8n is installed **natively on Windows** (same machine as MCP server)
- **Use Method B (HTTP)** when you have:
  - **n8n running in Docker** (even on the same Windows machine)
  - Remote n8n instance
  - High-frequency requests

**⚠️ Important for Docker Users**: If you're running n8n in Docker (like Docker Desktop), you **must use Method B (HTTP)** because the Docker container cannot directly execute Windows programs.

**💡 Testing Tip**: When testing the HTTP server, you'll need **two command prompts**:
- **Terminal 1**: Running the HTTP server (`python main.py http`)
- **Terminal 2**: Testing the endpoints (`curl http://127.0.0.1:8080/health`)

## Table of Contents

- [Prerequisites](#prerequisites)
- [System Requirements](#system-requirements)
- [MCP Server Installation](#mcp-server-installation)
- [n8n Configuration](#n8n-configuration)
- [Connection Validation](#connection-validation)
- [Troubleshooting](#troubleshooting)
- [Security Considerations](#security-considerations)
- [Next Steps](#next-steps)

## Prerequisites

Before beginning the integration setup, ensure you have the following prerequisites in place:

### Required Software

1. **Microsoft Outlook** (2016 or later)
   - Must be installed and configured with at least one email account
   - Outlook must be accessible via COM interface
   - Account should be actively syncing emails

2. **n8n** (Latest version recommended)
   - Local n8n installation (desktop app or self-hosted)
   - Administrative access to n8n configuration
   - Ability to create and modify workflows

3. **Python** (3.8 or later)
   - Required for running the MCP server
   - pip package manager installed
   - Administrative privileges for package installation

### System Access Requirements

- **Windows Operating System** (Windows 10 or later)
- **Administrator privileges** for initial setup
- **Network access** on localhost (127.0.0.1)
- **Firewall permissions** for local communication

### Outlook Configuration Verification

⚠️ **IMPORTANT**: **Always launch Outlook BEFORE starting the MCP server!**

Before proceeding, verify your Outlook setup:

```powershell
# Test Outlook COM access (PowerShell)
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    Write-Host "✅ Outlook COM access successful"
    Write-Host "📧 Available accounts: $($namespace.Accounts.Count)"
    $outlook.Quit()
} catch {
    Write-Host "❌ Outlook COM access failed: $($_.Exception.Message)"
}
```

## System Requirements

### Hardware Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| **RAM** | 4 GB | 8 GB or more |
| **Storage** | 2 GB free space | 5 GB free space |
| **CPU** | Dual-core 2.0 GHz | Quad-core 2.5 GHz or higher |
| **Network** | Localhost connectivity | Localhost connectivity |

### Software Compatibility

| Software | Version | Notes |
|----------|---------|-------|
| **Windows** | 10 (1903+) or 11 | Required for COM interface |
| **Microsoft Outlook** | 2016, 2019, 365 | Must support COM automation |
| **Python** | 3.8 - 3.12 | Tested versions |
| **n8n** | 1.0+ | Latest version recommended |

### Network Requirements

**Method A (stdio Mode):**
- **No network ports required**: Communication via standard input/output
- **Local process execution**: n8n must be able to execute Python processes
- **File system access**: n8n needs read access to MCP server files
- **Same machine only**: n8n and MCP server must be on same computer

**Method B (HTTP Mode):**
- **Port availability**: Default port 8080 (configurable)
- **Network access**: HTTP communication on localhost or network
- **Firewall configuration**: May need to allow port 8080
- **Remote access capable**: n8n can be on different computer than MCP server

## MCP Server Installation

### Step 1: Download and Extract

```bash
# Option A: Clone from repository
git clone https://github.com/bitsplus1/EmailMCP.git
cd outlook-mcp-server

# Option B: Download ZIP and extract
# Download from: https://github.com/bitsplus1/EmailMCP/archive/main.zip
# Extract to desired directory and navigate to it
```

### Step 2: Install Python Dependencies

```bash
# Install required packages
pip install -r requirements.txt

# Verify installation
pip list | findstr "pywin32"
pip list | findstr "pytest"
```

Expected output:
```
pywin32                   306
pytest                    7.4.3
pytest-asyncio           0.21.1
```

### Step 3: Test MCP Server Installation

```bash
# Test Outlook connection
python main.py test
```

Expected successful output:
```
🔍 Testing Outlook MCP Server...
✅ Python environment: OK
✅ Required packages: OK
✅ Outlook COM interface: OK
✅ Outlook connection: OK
✅ Email access: OK (Found X folders)
✅ MCP server functionality: OK

🎉 Installation test completed successfully!
Server is ready for n8n integration.
```

### Step 4: Create Configuration File

```bash
# Generate default configuration
python main.py create-config
```

This creates `outlook_mcp_server_config.json` with default settings:

```json
{
  "log_level": "INFO",
  "log_dir": "logs",
  "max_concurrent_requests": 10,
  "request_timeout": 30,
  "outlook_connection_timeout": 10,
  "enable_performance_logging": true,
  "enable_console_output": true,
  "server_host": "127.0.0.1",
  "server_port": 8080,
  "rate_limiting": {
    "enabled": true,
    "requests_per_minute": 100,
    "burst_allowance": 10
  },
  "caching": {
    "enabled": true,
    "email_cache_ttl": 300,
    "folder_cache_ttl": 600,
    "max_cache_size_mb": 100
  },
  "security": {
    "allowed_folders": [],
    "blocked_folders": ["Deleted Items"],
    "max_email_size_mb": 50,
    "sanitize_html": true
  }
}
```

### Step 5: Start MCP Server

```bash
# Start server in stdio mode (for n8n integration)
python main.py stdio

# Or start with custom configuration
python main.py stdio --config outlook_mcp_server_config.json

# For testing, use interactive mode
python main.py interactive --log-level DEBUG
```

Successful startup output:
```
📧 Outlook MCP Server Starting...
✅ Configuration loaded
✅ Outlook connection established
✅ MCP protocol handler initialized
✅ Server ready for connections

🚀 Server running in stdio mode
📡 Ready to receive n8n requests
```

## n8n Configuration

### Step 1: Verify n8n Installation

Ensure n8n is properly installed and running:

```bash
# Check n8n version
n8n --version

# Start n8n (if not already running)
n8n start

# Or for desktop app users, launch the n8n desktop application
```

## Method A: stdio Mode Setup (Recommended for Local Use)

### Step A1: Start MCP Server in stdio Mode

**You don't need to start the server manually** - the Execute Command node will start it automatically for each request.

To test that the server works:
```bash
cd C:\path\to\outlook-mcp-server
python main.py test
```

#### Pre-made Batch Files for n8n

The MCP server includes ready-to-use batch files for n8n integration:

- **`n8n_test_folders.bat`** - Gets list of all Outlook folders
- **`n8n_list_emails.bat`** - Lists first 5 emails from Inbox

Test these batch files:
```bash
# Test folder listing
cmd /c n8n_test_folders.bat

# Test email listing  
cmd /c n8n_list_emails.bat
```

You should see JSON responses like:
```json
{"jsonrpc": "2.0", "id": "n8n-test", "error": {"code": -32000, "message": "No active session. Handshake required."}}
```

The "handshake required" error is expected for single requests - this is normal MCP protocol behavior.

### Step A2: Configure n8n Execute Command Node

1. **Open n8n** in your browser (typically http://localhost:5678)

2. **Create a new workflow**

3. **Add Execute Command node**:
   - Click the **"+"** button to add a new node
   - Search for **"Execute Command"**
   - Click on **"Execute Command"** to add it to your workflow

4. **Configure the Execute Command node**:

   In n8n version 1.111.0, the Execute Command node has a simplified interface with just a **Command** field.

   **Recommended approach - Use cmd with batch file:**

   **Command**: 
   ```cmd
   cmd /c "C:\Users\USER\Documents\Github Repo\EmailMCP\mcp_test.bat"
   ```

   **Alternative approach - PowerShell with batch file:**
   ```powershell
   powershell -Command "& 'C:\Users\USER\Documents\Github Repo\EmailMCP\mcp_test.bat'"
   ```

   **Alternative approach - Full command with cmd:**
   ```cmd
   cmd /c "cd /d \"C:\Users\USER\Documents\Github Repo\EmailMCP\" && echo {\"jsonrpc\": \"2.0\", \"id\": \"test-folders\", \"method\": \"get_folders\", \"params\": {}} | python main.py single-request"
   ```

   **Why this is needed**: n8n on Windows sometimes uses Unix shell (`/bin/sh`) instead of Windows Command Prompt, causing path issues with spaces.

5. **Save the node configuration**

#### Alternative Method (Using PowerShell):

If the above doesn't work, try using PowerShell syntax:

**Command**:
```powershell
powershell -Command "cd 'C:\path\to\outlook-mcp-server'; echo '{\"jsonrpc\": \"2.0\", \"id\": \"test-folders\", \"method\": \"get_folders\", \"params\": {}}' | python main.py stdio"
```

#### Visual Guide for Execute Command Node (n8n 1.111.0):

**Method 1: Using Batch File (Recommended)**
```
┌─────────────────────────────────────┐
│ Execute Command                     │
├─────────────────────────────────────┤
│ Command:                            │
│ C:\path\to\outlook-mcp-server\      │
│ n8n_test_folders.bat                │
└─────────────────────────────────────┘
```

**Method 2: Full Command**
```
┌─────────────────────────────────────┐
│ Execute Command                     │
├─────────────────────────────────────┤
│ Command:                            │
│ cd C:\path\to\outlook-mcp-server && │
│ echo {"jsonrpc": "2.0", "id":       │
│ "test-folders", "method":           │
│ "get_folders", "params": {}} |      │
│ python main.py single-request       │
└─────────────────────────────────────┘
```

#### Recommended Approach: Use Batch Files (Easiest)

The MCP server includes pre-made batch files for easy n8n integration:

**For get_folders (list all Outlook folders):**
```bash
C:\path\to\outlook-mcp-server\n8n_test_folders.bat
```

**For list_emails (get emails from Inbox):**
```bash
C:\path\to\outlook-mcp-server\n8n_list_emails.bat
```

#### Manual Command Approach:

If you prefer to write the full command:

**For get_folders:**
```bash
cd C:\path\to\outlook-mcp-server && echo {"jsonrpc": "2.0", "id": "folders", "method": "get_folders", "params": {}} | python main.py single-request
```

**For list_emails:**
```bash
cd C:\path\to\outlook-mcp-server && echo {"jsonrpc": "2.0", "id": "emails", "method": "list_emails", "params": {"folder": "Inbox", "limit": 10}} | python main.py single-request
```

**For search_emails:**
```bash
cd C:\path\to\outlook-mcp-server && echo {"jsonrpc": "2.0", "id": "search", "method": "search_emails", "params": {"query": "subject:test", "limit": 5}} | python main.py single-request
```

**Important Notes:**
- Use `single-request` mode instead of `stdio` for n8n Execute Command
- Replace `C:\path\to\outlook-mcp-server` with your actual installation path
- The batch files are the easiest and most reliable method

---

## Method B: HTTP Mode Setup (For Remote Access & Docker n8n)

**⚠️ Docker Users**: This is the **only method** that works if you're running n8n in Docker, even on the same Windows machine.

### Step B1: Start MCP Server in HTTP Mode

⚠️ **CRITICAL REQUIREMENT**: **Outlook MUST be running before starting the MCP server!**

1. **Launch Outlook first** (REQUIRED!):
```bash
start outlook
```
**Wait for Outlook to fully load and log into your email account before proceeding.**

2. **Open a command prompt** and navigate to your MCP server directory:
```bash
cd "C:\Users\USER\Documents\Github Repo\EmailMCP"
```

3. **Start the HTTP server**:

   **For Docker n8n users (recommended)**:
   ```bash
   python main.py http --config docker_config.json
   ```

   **For native n8n users**:
   ```bash
   python main.py http
   ```

You should see output like:
```
✅ MCP HTTP server started successfully on http://0.0.0.0:8080
Available endpoints:
  - POST http://127.0.0.1:8080/mcp (MCP requests)
  - GET  http://127.0.0.1:8080/health (Health check)
```

4. **Keep this command prompt open** - the server needs to stay running to handle requests.

5. **Test the HTTP server** (open a **new command prompt**):

   **Test from Windows host**:
   ```bash
   # Test with curl
   curl -s http://127.0.0.1:8080/health
   
   # Or test with PowerShell
   Invoke-RestMethod -Uri "http://127.0.0.1:8080/health" -Method GET
   ```

   **⚠️ IMPORTANT FOR DOCKER N8N USERS**: If you're running n8n in Docker Desktop, you **MUST** use your Windows machine's IP address that starts with **192.168.x.x** (not 127.0.0.1 or localhost).

   **Find your Windows IP address**:
   ```bash
   # Method 1: Using ipconfig
   ipconfig | findstr "IPv4"
   
   # Method 2: Using PowerShell
   (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "Wi-Fi" -PrefixLength 24).IPAddress
   
   # Method 3: Using wmic
   wmic computersystem get name && ipconfig | findstr /i "ipv4"
   ```

   **Test with your Windows IP** (for Docker access):
   ```bash
   # Replace 192.168.1.164 with your actual Windows IP (starts with 192.168.x.x)
   curl -s http://192.168.1.164:8080/health
   ```

   **Expected response**:
   ```json
   {
     "status": "healthy",
     "timestamp": "2025-09-21T12:36:17.464976",
     "server_info": {
       "name": "outlook-mcp-server",
       "version": "1.0.0",
       "capabilities": {...}
     }
   }
   ```

### Step B2: Configure n8n HTTP Request Node

1. **Open n8n** in your browser

2. **Create a new workflow**

3. **Add HTTP Request node**:
   - Click the **"+"** button to add a new node
   - Search for **"HTTP Request"**
   - Click on **"HTTP Request"** to add it to your workflow

4. **Configure the HTTP Request node**:

   **Basic Configuration:**
   - **Method**: Select `POST` from dropdown
   - **URL**: Choose the URL that worked in your connectivity test:
     - **Docker n8n (REQUIRED)**: `http://192.168.x.x:8080/mcp` (replace with your Windows IP that starts with 192.168)
     - **Alternative for Docker**: `http://host.docker.internal:8080/mcp` (if host.docker.internal worked)
     - **Native n8n**: `http://127.0.0.1:8080/mcp`
     
   **⚠️ Docker Users**: You **MUST** use your Windows machine's IP address (192.168.x.x format) because Docker containers cannot access localhost/127.0.0.1 on the Windows host.

   **Headers Section:**
   - Click **"Add Header"**
   - **Name**: `Content-Type`
   - **Value**: `application/json`

   **Body Section:**
   - **Body Content Type**: Select `JSON` from dropdown
   - **JSON**: 
   ```json
   {
     "jsonrpc": "2.0",
     "id": "test-folders",
     "method": "get_folders",
     "params": {}
   }
   ```

   **Options:**
   - **Timeout**: `30000` (30 seconds)
   - **Retry on Fail**: Toggle `ON`
   - **Max Retries**: `3`

5. **Save the node configuration**

#### Visual Guide for HTTP Request Node:
```
┌─────────────────────────────────────┐
│ HTTP Request                        │
├─────────────────────────────────────┤
│ Method: [POST ▼]                    │
│ URL: http://127.0.0.1:8080/mcp      │
├─────────────────────────────────────┤
│ Headers:                            │
│ └─ Content-Type: application/json   │
├─────────────────────────────────────┤
│ Body Content Type: [JSON ▼]         │
│ JSON:                               │
│ {                                   │
│   "jsonrpc": "2.0",                 │
│   "id": "test-folders",             │
│   "method": "get_folders",          │
│   "params": {}                      │
│ }                                   │
├─────────────────────────────────────┤
│ Options:                            │
│ ┌─ Timeout: 30000                   │
│ ┌─ Retry on Fail: [ON]              │
│ └─ Max Retries: 3                   │
└─────────────────────────────────────┘
```

---

## MCP Function Examples for n8n HTTP Request

The Outlook MCP Server provides four main functions. Here are complete n8n HTTP Request node configurations for each:

### 🗂️ Function 1: get_folders
**Purpose**: Get a list of all available Outlook folders

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "get-folders-request",
  "method": "get_folders",
  "params": {}
}
```

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "get-folders-request",
  "result": {
    "folders": [
      {
        "id": "000000001A447390AA...",
        "name": "Inbox",
        "full_path": "Inbox",
        "item_count": 150,
        "unread_count": 5,
        "parent_folder": "",
        "folder_type": "Mail",
        "accessible": true,
        "has_subfolders": false,
        "display_name": "Inbox"
      }
    ]
  }
}
```

---

### 📧 Function 2: list_inbox_emails
**Purpose**: Get a list of emails from the default inbox folder (simple method)

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "list-inbox-emails-request",
  "method": "list_inbox_emails",
  "params": {
    "limit": 10,
    "unread_only": false
  }
}
```

**Parameter Options:**
- **limit** (optional): Number of emails to return (default: 50, max: 1000)
- **unread_only** (optional): Filter to show only unread emails (default: false)

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "list-inbox-emails-request",
  "result": {
    "emails": [
      {
        "id": "000000001A447390AA...",
        "subject": "Meeting Tomorrow",
        "sender": "john.doe@company.com",
        "sender_name": "John Doe",
        "recipients": ["me@company.com"],
        "received_time": "2025-09-22T10:30:00Z",
        "sent_time": "2025-09-22T10:25:00Z",
        "importance": "Normal",
        "has_attachments": false,
        "is_read": false,
        "folder": "Inbox",
        "size": 2048,
        "body_preview": "Hi, just wanted to confirm our meeting..."
      }
    ],
    "total_count": 150,
    "folder": "Inbox"
  }
}
```

---

### 📧 Function 3: list_emails
**Purpose**: Get a list of emails from a specific folder by folder ID

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "list-emails-request",
  "method": "list_emails",
  "params": {
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "limit": 10,
    "unread_only": false
  }
}
```

**Parameter Options:**
- **folder_id** (required): Folder ID to list emails from (use get_folders to see available folder IDs)
- **limit** (optional): Number of emails to return (default: 50, max: 1000)
- **unread_only** (optional): Filter to show only unread emails (default: false)

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "list-emails-request",
  "result": {
    "emails": [
      {
        "id": "000000001A447390AA...",
        "subject": "Meeting Tomorrow",
        "sender": "john.doe@company.com",
        "sender_name": "John Doe",
        "recipients": ["me@company.com"],
        "cc": [],
        "bcc": [],
        "received_time": "2025-09-22T10:30:00Z",
        "sent_time": "2025-09-22T10:25:00Z",
        "importance": "Normal",
        "has_attachments": false,
        "is_read": false,
        "folder": "Inbox",
        "size": 2048,
        "body_preview": "Hi, just wanted to confirm our meeting..."
      }
    ],
    "total_count": 150,
    "folder": "Inbox"
  }
}
```

---

### 📄 Function 4: get_email
**Purpose**: Get detailed information about a specific email

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "get-email-request",
  "method": "get_email",
  "params": {
    "email_id": "000000001A447390AA447390AA447390AA447390AA447390AA447390AA447390AA447390",
    "include_body": true,
    "include_attachments": true,
    "body_format": "html"
  }
}
```

**Parameter Options:**
- **email_id** (required): The unique ID of the email (get this from list_emails)
- **include_body** (optional): Include full email body (default: true)
- **include_attachments** (optional): Include attachment information (default: true)
- **body_format** (optional): "text", "html", or "both" (default: "html")

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "get-email-request",
  "result": {
    "email": {
      "id": "000000001A447390AA...",
      "subject": "Meeting Tomorrow",
      "sender": "john.doe@company.com",
      "sender_name": "John Doe",
      "recipients": ["me@company.com"],
      "cc": [],
      "bcc": [],
      "received_time": "2025-09-22T10:30:00Z",
      "sent_time": "2025-09-22T10:25:00Z",
      "importance": "Normal",
      "has_attachments": true,
      "is_read": false,
      "folder": "Inbox",
      "size": 15360,
      "body_html": "<html><body><p>Hi,</p><p>Just wanted to confirm our meeting tomorrow at 2 PM...</p></body></html>",
      "body_text": "Hi,\n\nJust wanted to confirm our meeting tomorrow at 2 PM...",
      "attachments": [
        {
          "name": "agenda.pdf",
          "size": 245760,
          "type": "application/pdf",
          "attachment_id": "ATT001"
        }
      ],
      "headers": {
        "message-id": "<abc123@company.com>",
        "thread-topic": "Meeting Tomorrow",
        "x-priority": "3"
      }
    }
  }
}
```

---

### 🔍 Function 5: search_emails
**Purpose**: Search for emails using various criteria

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "search-emails-request",
  "method": "search_emails",
  "params": {
    "query": "subject:meeting AND from:john.doe@company.com",
    "folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000",
    "limit": 20
  }
}
```

**Parameter Options:**
- **query** (required): Search query using Outlook search syntax
- **folder_id** (optional): Folder ID to limit search to specific folder (use get_folders to see available folder IDs)
- **limit** (optional): Number of results to return (default: 50, max: 100)

**Search Query Examples:**
- `"subject:meeting"` - Emails with "meeting" in subject
- `"from:john@company.com"` - Emails from specific sender
- `"hasattachments:yes"` - Emails with attachments
- `"received:today"` - Emails received today
- `"importance:high"` - High importance emails
- `"subject:project AND from:manager@company.com"` - Combined criteria

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "search-emails-request",
  "result": {
    "emails": [
      {
        "id": "000000001A447390AA...",
        "subject": "Project Meeting Tomorrow",
        "sender": "john.doe@company.com",
        "sender_name": "John Doe",
        "recipients": ["me@company.com"],
        "received_time": "2025-09-22T10:30:00Z",
        "sent_time": "2025-09-22T10:25:00Z",
        "importance": "Normal",
        "has_attachments": false,
        "is_read": false,
        "folder": "Inbox",
        "body_preview": "Hi, just wanted to confirm our project meeting..."
      }
    ],
    "total_count": 5,
    "query": "subject:meeting AND from:john.doe@company.com",
    "search_folder_id": "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFE000006B7F6A90000"
  }
}
```

---

### 📧 Function 6: send_email
**Purpose**: Send an email through Outlook

**HTTP Request Configuration:**
- **Method**: `POST`
- **URL**: `http://127.0.0.1:8080/mcp` (or your configured URL)
- **Headers**: `Content-Type: application/json`
- **Body (JSON)**:
```json
{
  "jsonrpc": "2.0",
  "id": "send-email-request",
  "method": "send_email",
  "params": {
    "to_recipients": ["recipient@company.com", "another@example.com"],
    "subject": "Automated Report - Daily Summary",
    "body": "<html><body><h2>Daily Report</h2><p>Please find the daily summary attached.</p><p>Best regards,<br>Automation System</p></body></html>",
    "cc_recipients": ["manager@company.com"],
    "bcc_recipients": ["archive@company.com"],
    "body_format": "html",
    "importance": "normal",
    "attachments": ["C:\\reports\\daily_summary.pdf"],
    "save_to_sent_items": true
  }
}
```

**Parameter Options:**
- **to_recipients** (required): List of recipient email addresses
- **subject** (required): Email subject line
- **body** (required): Email body content
- **cc_recipients** (optional): List of CC recipients (default: none)
- **bcc_recipients** (optional): List of BCC recipients (default: none)
- **body_format** (optional): "html", "text", or "rtf" (default: "html")
- **importance** (optional): "low", "normal", or "high" (default: "normal")
- **attachments** (optional): List of file paths to attach (default: none)
- **save_to_sent_items** (optional): Save to Sent Items folder (default: true)

**Body Format Examples:**
- **HTML**: `"<html><body><h1>Title</h1><p>Content with <b>formatting</b></p></body></html>"`
- **Text**: `"Plain text email content\n\nWith line breaks"`
- **RTF**: Rich Text Format for advanced formatting

**Use Cases:**
- **Automated Reports**: Send daily/weekly reports with attachments
- **Notifications**: Alert users about system events or status changes
- **Follow-ups**: Send automated follow-up emails based on triggers
- **Data Export**: Email processed data or analysis results
- **Alerts**: Send high-priority notifications to multiple recipients

**Expected Response:**
```json
{
  "jsonrpc": "2.0",
  "id": "send-email-request",
  "result": {
    "email_id": "000000001A447390AA447390AA447390AA447390AA447390AA447390AA447390AA447390",
    "status": "sent",
    "recipients": {
      "to": ["recipient@company.com", "another@example.com"],
      "cc": ["manager@company.com"],
      "bcc": ["archive@company.com"]
    },
    "subject": "Automated Report - Daily Summary",
    "body_format": "html",
    "importance": "normal",
    "attachments_count": 1,
    "sent_time": "2025-09-22T14:30:00Z"
  }
}
```

**Error Responses:**
```json
{
  "jsonrpc": "2.0",
  "id": "send-email-request",
  "error": {
    "code": -32602,
    "message": "Invalid email address: invalid_email",
    "data": {
      "parameter": "to_recipients",
      "validation_error": "Email format validation failed"
    }
  }
}
```

---

### 💡 Quick Start Templates

**Template 1: Get All Folders**
```json
{"jsonrpc": "2.0", "id": "folders", "method": "get_folders", "params": {}}
```

**Template 2: Get Latest 5 Emails from Inbox (Simple)**
```json
{"jsonrpc": "2.0", "id": "inbox", "method": "list_inbox_emails", "params": {"limit": 5}}
```

**Template 3: Get Latest 5 Emails from Specific Folder**
```json
{"jsonrpc": "2.0", "id": "folder", "method": "list_emails", "params": {"folder_id": "YOUR_FOLDER_ID", "limit": 5}}
```

**Template 4: Search Today's Emails**
```json
{"jsonrpc": "2.0", "id": "today", "method": "search_emails", "params": {"query": "received:today", "limit": 10}}
```

**Template 5: Search in Specific Folder**
```json
{"jsonrpc": "2.0", "id": "search-folder", "method": "search_emails", "params": {"query": "subject:meeting", "folder_id": "YOUR_FOLDER_ID", "limit": 10}}
```

**Template 6: Get Email with Full Content**
```json
{"jsonrpc": "2.0", "id": "detail", "method": "get_email", "params": {"email_id": "YOUR_EMAIL_ID", "include_body": true}}
```

**Template 7: Send Simple Email**
```json
{"jsonrpc": "2.0", "id": "send", "method": "send_email", "params": {"to_recipients": ["user@example.com"], "subject": "Test Email", "body": "Hello from n8n automation!"}}
```

**Template 8: Send HTML Email with Attachments**
```json
{"jsonrpc": "2.0", "id": "send-rich", "method": "send_email", "params": {"to_recipients": ["user@example.com"], "subject": "Report", "body": "<h1>Report</h1><p>Please see attached.</p>", "body_format": "html", "attachments": ["C:\\reports\\data.xlsx"]}}
```

**Template 9: n8n Production Example (Tested & Working)**
```json
{"jsonrpc": "2.0", "id": "send-email-request", "method": "send_email", "params": {"to_recipients": ["recipient@company.com"], "subject": "Automated Report - Daily Summary", "body": "<html><body><h2>Daily Report</h2><p>Please find the daily summary attached.</p><p>Best regards,<br>Automation System</p></body></html>", "cc_recipients": [], "bcc_recipients": [], "body_format": "html", "importance": "normal", "attachments": [], "save_to_sent_items": false}}
```

---

### Step B3: Configure for Docker n8n (Required for Docker Users)

If you're running n8n in Docker on the same Windows machine:

1. **The MCP server is already accessible** at `http://127.0.0.1:8080/mcp` from your Windows host

2. **For Docker n8n to access the Windows host**, use one of these URLs in your n8n HTTP Request:

   **Option 1: Use host.docker.internal (Recommended)**
   ```
   http://host.docker.internal:8080/mcp
   ```

   **Option 2: Use your Windows machine's IP address**
   ```
   http://192.168.1.100:8080/mcp
   ```
   (Replace `192.168.1.100` with your actual Windows IP address)

3. **Find your Windows IP address** (if needed):
   ```cmd
   ipconfig | findstr IPv4
   ```

4. **Test the connection** from inside the Docker container:
   ```bash
   # In n8n Execute Command node, test connectivity:
   curl http://host.docker.internal:8080/health
   ```

### Step B4: Fix Docker Connectivity Issues (If Needed)

If none of the URLs work from Docker n8n, you need to configure the server to accept connections from Docker:

1. **Create a configuration file** `docker_config.json`:
```json
{
  "server_host": "0.0.0.0",
  "server_port": 8080,
  "log_level": "INFO"
}
```

2. **Start server with Docker-friendly configuration**:
```bash
python main.py http --config docker_config.json
```

3. **Test with your Windows IP** from n8n:
   - **URL**: `http://192.168.1.164:8080/health` (replace with your actual IP)

4. **If still not working, check Windows Firewall**:
```cmd
netsh advfirewall firewall add rule name="MCP Server Docker" dir=in action=allow protocol=TCP localport=8080
```

### Step B5: Configure for Remote Access (Optional)

If you want to access the MCP server from a completely different computer:

1. **Update server configuration** to bind to all interfaces:
```json
{
  "server_host": "0.0.0.0",
  "server_port": 8080
}
```

2. **Start server with custom config**:
```bash
python main.py http --config remote_config.json
```

3. **Configure firewall** to allow port 8080:
```bash
netsh advfirewall firewall add rule name="MCP Server" dir=in action=allow protocol=TCP localport=8080
```

**Security Warning**: When using remote access, ensure proper network security measures are in place.

## Creating Test Workflows

### Test Workflow for Method A (stdio Mode)

**Step A3.1: Create Test Workflow**

1. **Add Execute Command node** with the command from Step A2
2. **Rename to "Test MCP Connection (stdio)"**

3. **Add Code node for response processing**:
   - **Node name**: "Process stdio Response"
   - **JavaScript Code**:
   ```javascript
   // Process stdio MCP response from Execute Command node
   const commandOutput = items[0].json;
   
   // The output might be in different fields depending on n8n version
   let stdout = commandOutput.stdout || commandOutput.output || commandOutput;
   
   // If stdout is an object, try to get the string representation
   if (typeof stdout === 'object') {
     stdout = JSON.stringify(stdout);
   }
   
   try {
     // Try to parse the JSON response from the MCP server
     const response = JSON.parse(stdout.trim());
     
     if (response.result) {
       return [{
         json: {
           success: true,
           message: "MCP server connection successful (stdio)",
           folderCount: response.result.length,
           folders: response.result.map(f => f.name || f),
           mode: "stdio",
           raw_response: response
         }
       }];
     } else if (response.error) {
       return [{
         json: {
           success: false,
           message: "MCP server error",
           error: response.error.message,
           error_code: response.error.code,
           mode: "stdio"
         }
       }];
     }
   } catch (e) {
     return [{
       json: {
         success: false,
         message: "Failed to parse response",
         error: e.message,
         raw_output: stdout,
         mode: "stdio"
       }
     }];
   }
   ```

4. **Connect and test the workflow**

#### Troubleshooting n8n Execute Command Issues

If you get errors like `/bin/sh: not found` or path issues, n8n is using Unix shell instead of Windows. Here are the solutions:

**Solution 1: Force Windows Command Prompt (Recommended)**
```cmd
cmd /c "C:\Users\USER\Documents\Github Repo\EmailMCP\mcp_test.bat"
```

**Solution 2: Use PowerShell**
```powershell
powershell -Command "& 'C:\Users\USER\Documents\Github Repo\EmailMCP\mcp_test.bat'"
```

**Solution 3: PowerShell with Full Command**
```powershell
powershell -Command "Set-Location 'C:\Users\USER\Documents\Github Repo\EmailMCP'; '{\"jsonrpc\": \"2.0\", \"id\": \"test-folders\", \"method\": \"get_folders\", \"params\": {}}' | python main.py single-request"
```

**Solution 4: Windows CMD with Full Command**
```cmd
cmd /c "cd /d \"C:\Users\USER\Documents\Github Repo\EmailMCP\" && echo {\"jsonrpc\": \"2.0\", \"id\": \"test-folders\", \"method\": \"get_folders\", \"params\": {}} | python main.py single-request"
```

**For Your Specific Path:**
Your MCP server is at: `C:\Users\USER\Documents\Github Repo\EmailMCP`

The issue is that n8n tries to use Unix shell (`/bin/sh`) instead of Windows Command Prompt, and the spaces in "Github Repo" cause path issues.

### Test Workflow for Method B (HTTP Mode)

**Step B3.1: Test Docker Connectivity First (Docker Users Only)**

If you're using Docker n8n, you need to test which URL works to reach the Windows host. Since Docker n8n doesn't have `curl`, use an HTTP Request node instead:

**Test these URLs in order until one works:**

1. **Option 1: Try host.docker.internal** (HTTP Request node):
   - **Method**: `GET`
   - **URL**: `http://host.docker.internal:8080/health`

2. **Option 2: Try Windows IP address** (HTTP Request node):
   - **Method**: `GET`
   - **URL**: `http://192.168.1.164:8080/health` (replace with your actual Windows IP)

3. **Option 3: Try localhost** (HTTP Request node):
   - **Method**: `GET`
   - **URL**: `http://127.0.0.1:8080/health`

**Find Your Windows IP Address:**
```cmd
ipconfig | findstr IPv4
```
Look for the main network adapter IP (usually 192.168.x.x or 10.x.x.x).

**Expected response from working URL**:
```json
{
  "status": "healthy",
  "timestamp": "2025-09-21T12:36:17.464976",
  "server_info": {
    "name": "outlook-mcp-server",
    "version": "1.0.0",
    "capabilities": {...}
  }
}
```

**If none work**: The MCP server might not be accessible from Docker. You may need to configure the server to bind to all interfaces (see Step B4).

**Step B3.2: Create Test Workflow**

1. **Ensure HTTP server is running** (from Step B1)

2. **Add HTTP Request node** (configured as in Step B2)
3. **Rename to "Test MCP Connection (HTTP)"**

4. **Add Code node for response processing**:
   - **Node name**: "Process HTTP Response"
   - **JavaScript Code**:
   ```javascript
   // Process HTTP MCP response
   const response = items[0].json;
   
   if (response.result && response.result.folders && Array.isArray(response.result.folders)) {
     const folders = response.result.folders;
     return [{
       json: {
         success: true,
         message: "MCP server connection successful (HTTP)",
         folderCount: folders.length,
         folders: folders.map(f => f.name || f.display_name || 'Unknown'),
         folderDetails: folders.map(f => ({
           name: f.name || f.display_name,
           displayName: f.display_name,
           fullPath: f.full_path,
           itemCount: f.item_count,
           unreadCount: f.unread_count,
           folderType: f.folder_type,
           accessible: f.accessible,
           hasSubfolders: f.has_subfolders
         })),
         mode: "http",
         raw_response: response
       }
     }];
   } else if (response.error) {
     return [{
       json: {
         success: false,
         message: "MCP server error",
         error: response.error.message,
         error_code: response.error.code,
         mode: "http"
       }
     }];
   } else {
     return [{
       json: {
         success: false,
         message: "Unexpected response format",
         response: response,
         mode: "http"
       }
     }];
   }
   ```

5. **Connect and test the workflow**

### Expected Test Results

**Successful stdio Mode Response:**
```json
{
  "success": true,
  "message": "MCP server connection successful (stdio)",
  "folderCount": 5,
  "folders": ["Inbox", "Sent Items", "Drafts", "Deleted Items", "Outbox"],
  "mode": "stdio"
}
```

**Successful HTTP Mode Response:**
```json
{
  "success": true,
  "message": "MCP server connection successful (HTTP)",
  "folderCount": 9,
  "folders": ["Inbox", "Sent Items", "Drafts", "Deleted Items", "Outbox", "Calendar", "Contacts", "Journal", "Tasks"],
  "folderDetails": [
    {
      "name": "記事",
      "displayName": "記事",
      "fullPath": "記事",
      "itemCount": 0,
      "unreadCount": 0,
      "folderType": "Mail",
      "accessible": true,
      "hasSubfolders": false
    },
    {
      "name": "連絡人",
      "displayName": "連絡人", 
      "fullPath": "連絡人",
      "itemCount": 0,
      "unreadCount": 0,
      "folderType": "Mail",
      "accessible": true,
      "hasSubfolders": false
    }
  ],
  "mode": "http"
}
```

### Step 4: Configure n8n Environment Variables (Optional)

For easier management, you can set environment variables in n8n:

```bash
# In n8n settings or environment file
MCP_SERVER_URL=http://127.0.0.1:8080/mcp
MCP_SERVER_TIMEOUT=30000
MCP_REQUEST_ID_PREFIX=n8n-workflow
```

Then use in workflows:
```json
{
  "url": "{{ $env.MCP_SERVER_URL }}",
  "timeout": "{{ $env.MCP_SERVER_TIMEOUT }}"
}
```

## Connection Validation

### Validation for Method A (stdio Mode)

1. **Test MCP Server Installation**:
```bash
cd C:\path\to\outlook-mcp-server
python main.py test
```
Expected output: "✅ Installation test completed successfully!"

2. **Test stdio Communication Manually**:
```bash
cd C:\path\to\outlook-mcp-server
echo {"jsonrpc": "2.0", "id": "test", "method": "get_folders", "params": {}} | python main.py stdio
```

3. **Run the stdio test workflow** in n8n and verify results

### Validation for Method B (HTTP Mode)

1. **Start HTTP Server**:
```bash
cd C:\path\to\outlook-mcp-server
python main.py http
```
Look for: "✅ MCP HTTP server started successfully on http://127.0.0.1:8080"

2. **Test HTTP Endpoints**:
```bash
# Health check
curl http://127.0.0.1:8080/health

# MCP request
curl -X POST http://127.0.0.1:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "id": "test", "method": "get_folders", "params": {}}'
```

3. **Run the HTTP test workflow** in n8n and verify results

### Validation Checklist

**For stdio Mode:**
- [ ] `python main.py test` passes
- [ ] Manual stdio test returns JSON response
- [ ] n8n Execute Command node executes successfully
- [ ] Response contains folder list

**For HTTP Mode:**
- [ ] `python main.py http` starts without errors
- [ ] Health endpoint returns 200 OK
- [ ] MCP endpoint accepts POST requests
- [ ] n8n HTTP Request node gets valid responses
- [ ] Server stays running between requests

### Step 2: Run Connection Test

Execute the test workflow created in the previous section:

1. **Open the test workflow** in n8n
2. **Click "Execute Workflow"**
3. **Check the results**

#### Expected Successful Response

```json
{
  "success": true,
  "message": "MCP server connection successful",
  "folderCount": 5,
  "folders": [
    "Inbox",
    "Sent Items",
    "Drafts",
    "Deleted Items",
    "Outbox"
  ]
}
```

#### Common Error Responses

**Connection Refused**:
```json
{
  "success": false,
  "message": "Connection refused",
  "error": "ECONNREFUSED 127.0.0.1:8080"
}
```
*Solution*: Ensure MCP server is running on the correct port.

**Timeout Error**:
```json
{
  "success": false,
  "message": "Request timeout",
  "error": "Timeout after 30000ms"
}
```
*Solution*: Check if Outlook is responsive and increase timeout.

**JSON-RPC Error**:
```json
{
  "success": false,
  "message": "MCP server error",
  "error": "Method not found: get_folders"
}
```
*Solution*: Verify MCP server version and method names.

### Step 3: Validate All MCP Methods

Create a comprehensive validation workflow:

```json
{
  "name": "Complete MCP Validation",
  "nodes": [
    {
      "name": "Test get_folders",
      "type": "n8n-nodes-base.httpRequest",
      "parameters": {
        "method": "POST",
        "url": "http://127.0.0.1:8080/mcp",
        "body": {
          "jsonrpc": "2.0",
          "id": "test-folders",
          "method": "get_folders",
          "params": {}
        }
      }
    },
    {
      "name": "Test list_emails",
      "type": "n8n-nodes-base.httpRequest",
      "parameters": {
        "method": "POST",
        "url": "http://127.0.0.1:8080/mcp",
        "body": {
          "jsonrpc": "2.0",
          "id": "test-list",
          "method": "list_emails",
          "params": {
            "folder": "Inbox",
            "limit": 5
          }
        }
      }
    },
    {
      "name": "Test search_emails",
      "type": "n8n-nodes-base.httpRequest",
      "parameters": {
        "method": "POST",
        "url": "http://127.0.0.1:8080/mcp",
        "body": {
          "jsonrpc": "2.0",
          "id": "test-search",
          "method": "search_emails",
          "params": {
            "query": "subject:test",
            "limit": 3
          }
        }
      }
    }
  ]
}
```

### Step 4: Performance Validation

Test the integration under load:

```json
{
  "name": "Performance Test",
  "nodes": [
    {
      "name": "Multiple Requests",
      "type": "n8n-nodes-base.httpRequest",
      "parameters": {
        "method": "POST",
        "url": "http://127.0.0.1:8080/mcp",
        "body": {
          "jsonrpc": "2.0",
          "id": "perf-test-{{ $runIndex }}",
          "method": "list_emails",
          "params": {
            "folder": "Inbox",
            "limit": 10
          }
        }
      }
    }
  ],
  "settings": {
    "executionOrder": "v1"
  }
}
```

Run this workflow multiple times to test concurrent request handling.

## Troubleshooting

### Troubleshooting Method A (stdio Mode)

#### Issue: "Command Failed" or "Process Exited with Code 1"

**Symptoms**:
- n8n Execute Command node fails
- Process exits with non-zero code
- No output or error in response

**Solutions**:

1. **Verify MCP server installation**:
```bash
cd C:\path\to\outlook-mcp-server
python main.py test
```

2. **Check Python path in n8n**:
   - Use full Python path: `C:\Python\python.exe`
   - Or try `python3` instead of `python`
   - Verify Python is in system PATH

3. **Test command manually**:
```bash
cd C:\path\to\outlook-mcp-server
echo {"jsonrpc": "2.0", "id": "test", "method": "get_folders", "params": {}} | python main.py stdio
```

4. **Check working directory**:
   - Use absolute paths: `C:\path\to\outlook-mcp-server`
   - Verify directory exists and contains main.py
   - Ensure proper permissions

#### Issue: "Invalid JSON Response"

**Solutions**:
- Ensure JSON input is on a single line
- Check for extra characters or line breaks
- Verify JSON syntax with a validator

### Troubleshooting Method B (HTTP Mode)

#### Issue: "Connection Refused" (ECONNREFUSED)

**Symptoms**:
- n8n HTTP Request shows ECONNREFUSED
- Cannot reach http://127.0.0.1:8080

**Solutions**:

1. **Verify HTTP server is running**:
```bash
# Check if server is running
tasklist | findstr python

# Check port usage
netstat -an | findstr :8080
```

2. **Start HTTP server properly**:
```bash
cd C:\path\to\outlook-mcp-server
python main.py http
# Look for "Server started successfully" message
```

3. **Test server manually**:
```bash
curl http://127.0.0.1:8080/health
```

4. **Check firewall settings**:
```bash
# Add firewall rule if needed
netsh advfirewall firewall add rule name="MCP Server" dir=in action=allow protocol=TCP localport=8080
```

#### Issue: "Server Starts But Doesn't Respond"

**Solutions**:
1. Check server logs for errors
2. Verify Outlook is running and accessible
3. Test with health endpoint first: `GET /health`
4. Ensure Content-Type header is set to `application/json`

#### Issue: "Timeout Errors"

**Solutions**:
1. Increase timeout in n8n HTTP Request node
2. Check if Outlook is responding slowly
3. Monitor server logs for performance issues
4. Consider using stdio mode for better performance

#### Issue: "Outlook Not Found" Error

**Symptoms**:
- MCP server fails to start
- Error: "Outlook application not found"

**Solutions**:

1. **Verify Outlook installation**:
```powershell
Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Outlook*"}
```

2. **Check COM registration**:
```cmd
# Re-register Outlook COM
regsvr32 /s "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
```

3. **Test COM access manually**:
```powershell
$outlook = New-Object -ComObject Outlook.Application
$outlook.Quit()
```

#### Issue: "Permission Denied" Error

**Symptoms**:
- Access denied when connecting to Outlook
- COM interface errors

**Solutions**:

1. **Run as Administrator**:
```bash
# Right-click Command Prompt -> "Run as administrator"
python main.py stdio
```

2. **Check Outlook security settings**:
   - Open Outlook
   - File → Options → Trust Center → Trust Center Settings
   - Programmatic Access → Never warn me about suspicious activity

3. **Verify user permissions**:
```cmd
whoami /priv | findstr SeServiceLogonRight
```

#### Issue: "JSON-RPC Method Not Found"

**Symptoms**:
- Method not found errors
- Invalid response format

**Solutions**:

1. **Verify method names**:
   - Use `get_folders` (not `getFolders`)
   - Use `list_emails` (not `listEmails`)
   - Use `search_emails` (not `searchEmails`)
   - Use `get_email` (not `getEmail`)

2. **Check request format**:
```json
{
  "jsonrpc": "2.0",
  "id": "unique-id",
  "method": "get_folders",
  "params": {}
}
```

3. **Validate JSON syntax**:
   - Use JSON validator
   - Check for trailing commas
   - Verify quote marks

#### Issue: n8n UI Configuration Problems

**Symptoms**:
- Can't find HTTP Request node
- Configuration fields not appearing
- JSON editor not working

**Solutions**:

1. **HTTP Request Node Not Found**:
   - Make sure you're using n8n version 1.0+
   - Search for "HTTP Request" (not just "HTTP")
   - Try "Webhook" node as alternative

2. **Missing Configuration Fields**:
   - Scroll down in the node editor
   - Click "Add Header" to add header fields
   - Switch Body Content Type to "JSON"
   - Check "Options" section at bottom

3. **JSON Editor Issues**:
   - Ensure Body Content Type is set to "JSON"
   - Copy-paste JSON instead of typing
   - Use external JSON validator first
   - Check for invisible characters

4. **Node Connection Problems**:
   - Drag from the small dot on the right of first node
   - Drop on the small dot on the left of second node
   - Look for the connecting line to appear
   - Save workflow after connecting

**UI Navigation Tips**:
```
Node Editor Layout:
┌─────────────────────────────────────┐
│ [Node Name]                    [×]  │ ← Close button
├─────────────────────────────────────┤
│ Basic Configuration                 │ ← Always visible
│ ┌─ Method: [Dropdown]              │
│ └─ URL: [Text field]               │
├─────────────────────────────────────┤
│ Headers                             │ ← Click to expand
│ [+ Add Header]                      │
├─────────────────────────────────────┤
│ Body                                │ ← Scroll to see
│ Content Type: [Dropdown]            │
├─────────────────────────────────────┤
│ Options                             │ ← At bottom
│ [Advanced settings]                 │
├─────────────────────────────────────┤
│                    [Execute] [Save] │ ← Action buttons
└─────────────────────────────────────┘
```

### Performance Issues

#### Issue: Slow Response Times

**Solutions**:

1. **Enable caching**:
```json
{
  "caching": {
    "enabled": true,
    "email_cache_ttl": 300,
    "folder_cache_ttl": 600
  }
}
```

2. **Increase concurrent requests**:
```json
{
  "max_concurrent_requests": 20,
  "request_timeout": 60
}
```

3. **Optimize Outlook**:
   - Reduce sync frequency
   - Limit folder synchronization
   - Clear Outlook cache

#### Issue: Memory Usage

**Solutions**:

1. **Configure memory limits**:
```json
{
  "caching": {
    "max_cache_size_mb": 50
  },
  "performance": {
    "lazy_loading_enabled": true
  }
}
```

2. **Monitor memory usage**:
```bash
# Check memory usage
tasklist /fi "imagename eq python.exe" /fo table

# Monitor in real-time
python examples/memory_monitor.py
```

### Network Issues

#### Issue: Port Already in Use

**Solutions**:

1. **Find process using port**:
```cmd
netstat -ano | findstr :8080
taskkill /PID <process_id> /F
```

2. **Use different port**:
```json
{
  "server_port": 8081
}
```

3. **Update n8n configuration**:
```json
{
  "url": "http://127.0.0.1:8081/mcp"
}
```

### Debugging Tools

#### Enable Debug Logging

```bash
# Start with debug logging
python main.py stdio --log-level DEBUG

# Check log files
type logs\outlook_mcp_server.log | findstr ERROR
```

#### Use Interactive Mode

```bash
# Interactive testing
python main.py interactive --log-level DEBUG

# Test specific methods
> get_folders
> list_emails folder=Inbox limit=5
```

#### Network Debugging

```bash
# Test HTTP endpoint directly
curl -X POST http://127.0.0.1:8080/mcp ^
  -H "Content-Type: application/json" ^
  -d "{\"jsonrpc\":\"2.0\",\"id\":\"test\",\"method\":\"get_folders\",\"params\":{}}"

# Use PowerShell for testing
Invoke-RestMethod -Uri "http://127.0.0.1:8080/mcp" -Method POST -ContentType "application/json" -Body '{"jsonrpc":"2.0","id":"test","method":"get_folders","params":{}}'
```

## Security Considerations

### Localhost-Only Access

The integration is designed for localhost-only communication, providing inherent security benefits:

- **No external network exposure**
- **No public IP requirements**
- **Firewall protection by default**
- **Local user context security**

### Configuration Security

#### Folder Access Control

Restrict access to specific Outlook folders:

```json
{
  "security": {
    "allowed_folders": [
      "Inbox",
      "Sent Items",
      "Projects"
    ],
    "blocked_folders": [
      "Deleted Items",
      "Personal",
      "Confidential"
    ]
  }
}
```

#### Content Filtering

Configure content sanitization:

```json
{
  "security": {
    "max_email_size_mb": 25,
    "sanitize_html": true,
    "strip_attachments": false,
    "max_search_results": 500
  }
}
```

#### Rate Limiting

Prevent abuse with rate limiting:

```json
{
  "rate_limiting": {
    "enabled": true,
    "requests_per_minute": 60,
    "burst_allowance": 10
  }
}
```

### n8n Security Settings

#### Workflow Security

- **Limit workflow access** to authorized users
- **Use environment variables** for sensitive configuration
- **Implement input validation** in Code nodes
- **Log security events** for audit trails

#### Network Security

- **Keep n8n updated** to latest version
- **Use HTTPS** for n8n web interface (if exposed)
- **Implement authentication** for n8n access
- **Monitor network traffic** for anomalies

### Monitoring and Auditing

#### Enable Audit Logging

```json
{
  "logging": {
    "enable_audit_logging": true,
    "audit_log_file": "logs/audit.log",
    "log_security_events": true
  }
}
```

#### Monitor Access Patterns

```bash
# Check access logs
findstr "list_emails\|get_email" logs\audit.log

# Monitor failed requests
findstr "ERROR\|FAILED" logs\outlook_mcp_server.log
```

#### Set Up Alerts

Create monitoring workflows in n8n:

```json
{
  "name": "Security Monitor",
  "trigger": "schedule",
  "interval": "5 minutes",
  "actions": [
    "check_error_rate",
    "validate_connection_health",
    "alert_on_anomalies"
  ]
}
```

## Next Steps

### Explore Integration Methods

Now that you have a working connection, explore different integration approaches:

1. **[Integration Methods Guide](N8N_INTEGRATION_METHODS.md)** - Learn about HTTP Request, Execute Command, and custom node options

2. **[Workflow Examples](N8N_WORKFLOW_EXAMPLES.md)** - Ready-to-use workflow templates for common scenarios

3. **[API Reference](N8N_API_REFERENCE.md)** - Complete API documentation with n8n-specific examples

### Advanced Configuration

#### Production Deployment

For production use, consider:

- **Windows Service installation**
- **Automated startup configuration**
- **Enhanced monitoring and alerting**
- **Backup and recovery procedures**

#### Performance Optimization

Optimize for your specific use case:

- **Tune caching parameters**
- **Adjust concurrent request limits**
- **Configure connection pooling**
- **Implement custom retry logic**

#### Custom Development

Extend the integration:

- **Create custom n8n nodes**
- **Develop specialized workflows**
- **Implement business-specific logic**
- **Add custom authentication**

### Community Resources

- **Documentation**: Complete guides in the `docs/` directory
- **Examples**: Sample workflows in `examples/` directory
- **Support**: GitHub issues and discussions
- **Updates**: Check for new releases and features

### Getting Help

If you encounter issues not covered in this guide:

1. **Check the logs**: Review `logs/outlook_mcp_server.log`
2. **Run diagnostics**: Use `python main.py test`
3. **Review documentation**: Check other guides in `docs/`
4. **Search issues**: Look for similar problems on GitHub
5. **Create issue**: Report new problems with detailed information

---

**🎉 Congratulations!** You have successfully set up the integration between n8n and the Outlook MCP Server. You can now create powerful email automation workflows using n8n's visual interface with full access to your Outlook data.

**Next recommended reading**: [N8N_INTEGRATION_METHODS.md](N8N_INTEGRATION_METHODS.md) to learn about different ways to use the MCP server in your n8n workflows.