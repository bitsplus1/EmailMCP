# Outlook MCP Server - Setup and Configuration Guide

This guide provides comprehensive instructions for setting up, configuring, and deploying the Outlook MCP Server.

## Table of Contents

- [System Requirements](#system-requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the Server](#running-the-server)
- [MCP Client Integration](#mcp-client-integration)
- [Security Configuration](#security-configuration)
- [Performance Tuning](#performance-tuning)
- [Deployment Options](#deployment-options)
- [Monitoring and Logging](#monitoring-and-logging)

## System Requirements

### Operating System
- **Windows 10** or later (required for COM interface)
- **Windows Server 2016** or later (for server deployments)

### Software Dependencies
- **Microsoft Outlook** (2016 or later)
  - Must be installed and configured with at least one email account
  - Outlook must be accessible via COM interface
- **Python 3.8** or later
- **pip** package manager

### Hardware Requirements
- **Minimum**: 2 GB RAM, 1 GB free disk space
- **Recommended**: 4 GB RAM, 2 GB free disk space
- **Network**: Internet connection for email synchronization

### Permissions
- **User Account**: Must have permissions to run Outlook
- **COM Access**: User must be able to access Outlook COM objects
- **File System**: Read/write access to installation directory and log directory

## Installation

### Step 1: Verify Prerequisites

First, ensure Microsoft Outlook is installed and working:

```cmd
# Test Outlook COM access (PowerShell)
powershell -Command "New-Object -ComObject Outlook.Application"
```

If this command succeeds without errors, Outlook COM access is working.

### Step 2: Clone or Download the Server

```bash
# Clone from repository
git clone https://github.com/bitsplus1/EmailMCP.git
cd outlook-mcp-server

# Or download and extract ZIP file
# Then navigate to the extracted directory
```

### Step 3: Install Python Dependencies

```bash
# Install required packages
pip install -r requirements.txt

# Or install with specific versions for stability
pip install pywin32==306 mcp==1.0.0 asyncio pytest pytest-asyncio
```

### Step 4: Verify Installation

```bash
# Test the installation
python main.py test

# Expected output:
# Testing Outlook connection...
# ✓ Outlook connection successful
# ✓ Server is healthy
# ✓ Found X accessible folders
# ✓ Email access working
# Connection test completed successfully
```

### Step 5: Create Configuration File

```bash
# Generate default configuration
python main.py create-config

# This creates outlook_mcp_server_config.json
```

## Configuration

### Basic Configuration File

The server uses a JSON configuration file (`outlook_mcp_server_config.json`):

```json
{
  "log_level": "INFO",
  "log_dir": "logs",
  "max_concurrent_requests": 10,
  "request_timeout": 30,
  "outlook_connection_timeout": 10,
  "enable_performance_logging": true,
  "enable_console_output": true,
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

### Configuration Options

#### Logging Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `log_level` | string | "INFO" | Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `log_dir` | string | "logs" | Directory for log files |
| `enable_performance_logging` | boolean | true | Enable performance metrics logging |
| `enable_console_output` | boolean | true | Enable console output |
| `log_rotation_size_mb` | integer | 10 | Log file rotation size in MB |
| `log_retention_days` | integer | 30 | Number of days to retain log files |

#### Performance Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `max_concurrent_requests` | integer | 10 | Maximum concurrent requests |
| `request_timeout` | integer | 30 | Request timeout in seconds |
| `outlook_connection_timeout` | integer | 10 | Outlook connection timeout |
| `connection_pool_size` | integer | 5 | Size of Outlook connection pool |
| `lazy_loading_enabled` | boolean | true | Enable lazy loading for email content |

#### Rate Limiting Configuration

```json
{
  "rate_limiting": {
    "enabled": true,
    "requests_per_minute": 100,
    "burst_allowance": 10,
    "per_client_limits": {
      "default": 100,
      "premium": 500
    }
  }
}
```

#### Caching Configuration

```json
{
  "caching": {
    "enabled": true,
    "email_cache_ttl": 300,
    "folder_cache_ttl": 600,
    "max_cache_size_mb": 100,
    "cache_cleanup_interval": 3600
  }
}
```

#### Security Configuration

```json
{
  "security": {
    "allowed_folders": ["Inbox", "Sent Items", "Projects"],
    "blocked_folders": ["Deleted Items", "Junk Email"],
    "max_email_size_mb": 50,
    "max_search_results": 1000,
    "sanitize_html": true,
    "strip_attachments": false
  }
}
```

### Environment Variables

You can also configure the server using environment variables:

```bash
# Windows Command Prompt
set OUTLOOK_MCP_LOG_LEVEL=DEBUG
set OUTLOOK_MCP_LOG_DIR=C:\logs\outlook-mcp
set OUTLOOK_MCP_MAX_CONCURRENT=20

# Windows PowerShell
$env:OUTLOOK_MCP_LOG_LEVEL="DEBUG"
$env:OUTLOOK_MCP_LOG_DIR="C:\logs\outlook-mcp"
$env:OUTLOOK_MCP_MAX_CONCURRENT=20

# Linux/Mac (if running under WSL or similar)
export OUTLOOK_MCP_LOG_LEVEL=DEBUG
export OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp
export OUTLOOK_MCP_MAX_CONCURRENT=20
```

Environment variables take precedence over configuration file settings.

## Running the Server

### HTTP Server Mode (Recommended for Testing)

The HTTP server mode is ideal for testing and integration with web applications:

```bash
# Run HTTP server with default configuration
python main.py http

# Run HTTP server with custom configuration (recommended)
python main.py http --config docker_config.json

# Run HTTP server with custom host and port
python main.py http --host 0.0.0.0 --port 8080 --config docker_config.json

# Test the HTTP server
curl -X POST http://192.168.1.100:8080/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":"1","method":"list_inbox_emails","params":{"limit":5}}'
```

### Standard MCP Mode (For MCP Clients)

For integration with MCP clients:

```bash
# Run with default configuration
python main.py stdio

# Run with custom configuration
python main.py stdio --config my_config.json

# Run with command-line overrides
python main.py stdio --log-level DEBUG --max-concurrent 20
```

### Interactive Mode (Development/Testing)

For development and testing:

```bash
# Interactive mode with console output
python main.py interactive

# Interactive mode with debug logging
python main.py interactive --log-level DEBUG
```

### Service Mode (Windows Service)

To run as a Windows service, create a service wrapper:

```python
# service_wrapper.py
import asyncio
import sys
from src.outlook_mcp_server.main import main

if __name__ == "__main__":
    # Set stdio mode for service
    sys.argv = ["main.py", "stdio", "--no-console"]
    asyncio.run(main())
```

Install as Windows service using `nssm` or similar tools:

```cmd
# Using NSSM (Non-Sucking Service Manager)
nssm install OutlookMCPServer
nssm set OutlookMCPServer Application "C:\Python\python.exe"
nssm set OutlookMCPServer AppParameters "C:\path\to\service_wrapper.py"
nssm set OutlookMCPServer AppDirectory "C:\path\to\outlook-mcp-server"
nssm start OutlookMCPServer
```

### Command Line Options

```bash
python main.py [mode] [options]

Modes:
  http          Run as HTTP server (recommended for testing)
  stdio         Run as MCP stdio server (for MCP clients)
  interactive   Run in interactive mode with console output
  test          Test Outlook connection and exit
  create-config Create sample configuration file

HTTP Mode Options:
  --host HOST         Host to bind to (default: localhost)
  --port PORT         Port to bind to (default: 8080)
  --config PATH       Path to configuration file (JSON format)

General Options:
  --config PATH       Path to configuration file (JSON format)
  --log-level LEVEL   Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
  --log-dir PATH      Directory for log files
  --max-concurrent N  Maximum number of concurrent requests
  --no-console        Disable console output
  --test-connection   Test Outlook connection and exit
```

## MCP Client Integration

### Client Configuration

Configure your MCP client to connect to the Outlook MCP Server:

#### Example: VS Code Extension Configuration

```json
{
  "mcp.servers": {
    "outlook": {
      "command": "python",
      "args": ["C:\\path\\to\\outlook-mcp-server\\main.py", "stdio"],
      "cwd": "C:\\path\\to\\outlook-mcp-server",
      "env": {
        "OUTLOOK_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

#### Example: Custom MCP Client

```python
import asyncio
from mcp_client import MCPClient

async def connect_to_outlook():
    # Configure client
    client = MCPClient(
        server_command=["python", "main.py", "stdio"],
        server_cwd="C:\\path\\to\\outlook-mcp-server"
    )
    
    # Connect
    await client.connect()
    
    # Test connection
    response = await client.call_method("get_folders", {})
    print(f"Available folders: {len(response)}")
    
    await client.disconnect()

asyncio.run(connect_to_outlook())
```

### Server Capabilities

The server advertises the following capabilities:

```json
{
  "capabilities": {
    "tools": [
      {
        "name": "list_emails",
        "description": "List emails from specified folders with filtering options"
      },
      {
        "name": "get_email",
        "description": "Retrieve detailed information for a specific email by ID"
      },
      {
        "name": "search_emails",
        "description": "Search emails based on user-defined queries"
      },
      {
        "name": "get_folders",
        "description": "List all available email folders in Outlook"
      }
    ]
  }
}
```

## Security Configuration

### Folder Access Control

Restrict access to specific folders:

```json
{
  "security": {
    "allowed_folders": [
      "Inbox",
      "Sent Items",
      "Projects",
      "Archive"
    ],
    "blocked_folders": [
      "Deleted Items",
      "Junk Email",
      "Personal"
    ]
  }
}
```

### Content Filtering

Configure content filtering and sanitization:

```json
{
  "security": {
    "max_email_size_mb": 50,
    "max_search_results": 1000,
    "sanitize_html": true,
    "strip_attachments": false,
    "allowed_attachment_types": [".pdf", ".docx", ".xlsx", ".txt"],
    "blocked_attachment_types": [".exe", ".bat", ".cmd", ".scr"]
  }
}
```

### Network Security

For network deployments, consider:

1. **Firewall Configuration**: Only allow necessary ports
2. **VPN Access**: Require VPN for remote access
3. **Authentication**: Implement client authentication if needed
4. **Encryption**: Use TLS for network communication

### User Permissions

Ensure the service account has appropriate permissions:

```cmd
# Grant "Log on as a service" right
secpol.msc -> Local Policies -> User Rights Assignment -> Log on as a service

# Grant Outlook COM access
dcomcnfg.exe -> Component Services -> DCOM Config -> Microsoft Outlook
```

## Performance Tuning

### Memory Optimization

```json
{
  "performance": {
    "max_cache_size_mb": 200,
    "cache_cleanup_interval": 1800,
    "lazy_loading_threshold": 1000000,
    "connection_pool_size": 10
  }
}
```

### Concurrency Settings

```json
{
  "performance": {
    "max_concurrent_requests": 20,
    "request_timeout": 60,
    "outlook_connection_timeout": 15,
    "batch_processing_size": 10
  }
}
```

### Logging Optimization

```json
{
  "logging": {
    "log_level": "WARNING",
    "enable_performance_logging": false,
    "log_rotation_size_mb": 50,
    "async_logging": true
  }
}
```

## Deployment Options

### Development Deployment

For development and testing:

```bash
# Simple development setup
python main.py interactive --log-level DEBUG
```

### Production Deployment

For production environments:

1. **Windows Service**: Use NSSM or similar
2. **Docker Container**: Containerize for consistent deployment
3. **Load Balancer**: Use multiple instances behind a load balancer

#### Docker Deployment

```dockerfile
# Dockerfile
FROM python:3.9-windowsservercore

WORKDIR /app
COPY . .

RUN pip install -r requirements.txt

EXPOSE 8080

CMD ["python", "main.py", "stdio", "--no-console"]
```

```yaml
# docker-compose.yml
version: '3.8'
services:
  outlook-mcp-server:
    build: .
    volumes:
      - ./logs:/app/logs
      - ./config:/app/config
    environment:
      - OUTLOOK_MCP_CONFIG_FILE=/app/config/production.json
    restart: unless-stopped
```

### High Availability Setup

For high availability:

1. **Multiple Instances**: Run multiple server instances
2. **Health Checks**: Implement health check endpoints
3. **Failover**: Configure automatic failover
4. **Monitoring**: Set up comprehensive monitoring

## Monitoring and Logging

### Log Files

The server generates several types of log files:

- **Application Logs**: `outlook_mcp_server.log`
- **Performance Logs**: `performance.log`
- **Error Logs**: `errors.log`
- **Audit Logs**: `audit.log`

### Log Format

Logs use structured JSON format:

```json
{
  "timestamp": "2024-01-15T10:30:00.123Z",
  "level": "INFO",
  "logger": "outlook_mcp_server.services.email_service",
  "message": "Email retrieved successfully",
  "email_id": "AAMkADEx...",
  "processing_time_ms": 45,
  "client_id": "client_001"
}
```

### Performance Monitoring

Monitor key metrics:

- **Request Rate**: Requests per second
- **Response Time**: Average response time
- **Error Rate**: Percentage of failed requests
- **Memory Usage**: Current memory consumption
- **Cache Hit Rate**: Cache effectiveness

### Health Checks

The server provides health check endpoints:

```bash
# Test server health
python main.py test

# Check specific components
python -c "
import asyncio
from src.outlook_mcp_server.server import OutlookMCPServer

async def health_check():
    server = OutlookMCPServer({})
    await server.start()
    health = server.get_health_status()
    print(f'Health: {health}')
    await server.stop()

asyncio.run(health_check())
"
```

### Alerting

Set up alerts for:

- **Service Down**: Server not responding
- **High Error Rate**: Error rate above threshold
- **High Memory Usage**: Memory usage above limit
- **Outlook Connection Lost**: COM connection failures

### Log Analysis

Use log analysis tools:

```bash
# Search for errors
findstr "ERROR" logs\outlook_mcp_server.log

# Analyze performance
python scripts/analyze_performance.py logs/performance.log

# Generate reports
python scripts/generate_report.py --start-date 2024-01-01 --end-date 2024-01-31
```

## Troubleshooting

### Common Issues

1. **Outlook Not Found**
   - Ensure Outlook is installed and configured
   - Check COM registration: `regsvr32 outlook.exe`

2. **Permission Denied**
   - Run as administrator
   - Check Outlook security settings
   - Verify user has Outlook access

3. **Connection Timeout**
   - Increase `outlook_connection_timeout`
   - Check Outlook responsiveness
   - Restart Outlook application

4. **High Memory Usage**
   - Reduce `max_cache_size_mb`
   - Enable cache cleanup
   - Restart server periodically

### Debug Mode

Enable debug mode for detailed troubleshooting:

```bash
python main.py interactive --log-level DEBUG
```

This provides detailed information about:
- COM object interactions
- Request/response details
- Performance metrics
- Error stack traces

### Support Resources

- **Documentation**: Check the docs/ directory
- **Log Files**: Review application logs
- **Test Mode**: Use `python main.py test`
- **Community**: Check GitHub issues and discussions