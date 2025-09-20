# Outlook MCP Server Logging System

This module provides a comprehensive logging system with structured JSON output, performance metrics, and log rotation capabilities for the Outlook MCP Server.

## Features

- **Structured JSON Logging**: All log entries are formatted as structured JSON for easy parsing and analysis
- **Log Rotation**: Automatic log file rotation to prevent files from growing too large
- **Performance Metrics**: Built-in timing and resource usage logging
- **Multiple Log Levels**: Support for DEBUG, INFO, WARNING, ERROR, and CRITICAL levels
- **MCP Protocol Logging**: Specialized logging for MCP requests and responses
- **Outlook Operation Logging**: Dedicated logging for Outlook COM operations
- **Connection Status Logging**: Track Outlook connection status (Requirement 5.5)
- **Context Managers**: Easy timing of operations with context managers
- **Environment Configuration**: Configure logging through environment variables

## Requirements Addressed

This logging system addresses the following requirements:

- **Requirement 5.2**: Log all significant activities and errors for debugging purposes
- **Requirement 5.5**: Verify Outlook connectivity and log the connection status
- **Requirement 8.4**: Provide clear error messages that aid in troubleshooting

## Quick Start

```python
from src.outlook_mcp_server.logging import configure_logging, get_logger

# Configure the logging system
configure_logging(
    log_level="INFO",
    log_dir="logs",
    console_output=True
)

# Get a logger instance
logger = get_logger(__name__)

# Basic logging
logger.info("Application started")
logger.error("An error occurred", exc_info=True)

# MCP protocol logging
logger.log_mcp_request("req-123", "list_emails", {"folder": "inbox"})
logger.log_mcp_response("req-123", "list_emails", success=True, duration=0.5)

# Outlook operation logging
logger.log_outlook_operation("connect", success=True, duration=0.1)
logger.log_connection_status(True, "Connected successfully")

# Performance timing
with logger.time_operation("email_search"):
    # Your operation here
    pass
```

## Configuration

### Programmatic Configuration

```python
from src.outlook_mcp_server.logging import configure_logging

configure_logging(
    log_level="DEBUG",           # Log level: DEBUG, INFO, WARNING, ERROR, CRITICAL
    log_dir="logs",              # Directory for log files
    max_bytes=10*1024*1024,      # Max file size before rotation (10MB)
    backup_count=5,              # Number of backup files to keep
    console_output=True          # Also output to console
)
```

### Environment Variables

You can also configure logging using environment variables:

```bash
export LOG_LEVEL=DEBUG
export LOG_DIR=/var/log/outlook-mcp
export LOG_MAX_FILE_SIZE_MB=20
export LOG_BACKUP_COUNT=10
export LOG_CONSOLE_OUTPUT=true
export LOG_PERFORMANCE=true
export LOG_REQUEST_TIMING=true
export LOG_RESOURCE_USAGE=true
```

Then use:

```python
from src.outlook_mcp_server.logging.config import LoggingConfig
from src.outlook_mcp_server.logging import configure_logging

config = LoggingConfig.from_environment()
configure_logging(
    log_level=config.level,
    log_dir=config.log_dir,
    max_bytes=config.max_bytes,
    backup_count=config.backup_count,
    console_output=config.console_output
)
```

## Log Format

All logs are formatted as structured JSON with the following fields:

```json
{
    "timestamp": "2024-01-15T10:30:45.123Z",
    "level": "INFO",
    "logger": "outlook_mcp_server.service",
    "message": "MCP request received: list_emails",
    "module": "email_service",
    "function": "list_emails",
    "line": 42,
    "thread": 12345,
    "thread_name": "MainThread",
    "extra": {
        "mcp": {
            "request_id": "req-123",
            "method": "list_emails",
            "params": {"folder": "inbox", "limit": 10},
            "type": "request"
        }
    }
}
```

### Exception Logging

When exceptions occur, additional fields are included:

```json
{
    "timestamp": "2024-01-15T10:30:45.123Z",
    "level": "ERROR",
    "message": "Failed to connect to Outlook",
    "exception": {
        "type": "COMError",
        "message": "COM object not available",
        "traceback": "Traceback (most recent call last):\n..."
    }
}
```

## Specialized Logging Methods

### MCP Protocol Logging

```python
# Log MCP requests
logger.log_mcp_request("req-123", "list_emails", {"folder": "inbox"})

# Log MCP responses with timing
logger.log_mcp_response("req-123", "list_emails", success=True, duration=0.5)
```

### Outlook Operation Logging

```python
# Log Outlook COM operations
logger.log_outlook_operation(
    "get_folders", 
    success=True, 
    duration=0.2,
    folder_count=15
)

# Log connection status (Requirement 5.5)
logger.log_connection_status(True, "Connected to Outlook successfully")
```

### Performance Logging

```python
# Direct performance logging
logger.performance.log_request_timing("operation_name", 0.123, success=True)
logger.performance.log_resource_usage(memory_mb=45.2, cpu_percent=12.5)

# Context manager for automatic timing
with logger.time_operation("email_search"):
    # Your operation here
    search_results = search_emails(query)
```

## Log Rotation

The logging system automatically rotates log files when they reach the configured size limit:

- **Default max size**: 10MB per file
- **Default backup count**: 5 files
- **Rotation behavior**: When the main log file reaches the size limit, it's renamed with a numeric suffix, and a new log file is created

Log files are named:
- `outlook_mcp_server.log` (current log)
- `outlook_mcp_server.log.1` (most recent backup)
- `outlook_mcp_server.log.2` (older backup)
- etc.

## Performance Considerations

- **Structured JSON**: While JSON formatting adds some overhead, it provides significant benefits for log analysis
- **Async Logging**: The logging system is thread-safe and can handle concurrent requests
- **Memory Usage**: Log rotation prevents unbounded disk usage
- **Performance Metrics**: Built-in performance logging helps identify bottlenecks

## Integration with Other Components

The logging system is designed to integrate seamlessly with other Outlook MCP Server components:

```python
# In service classes
class EmailService:
    def __init__(self):
        self.logger = get_logger(__name__)
    
    def list_emails(self, folder, limit):
        with self.logger.time_operation("list_emails"):
            self.logger.info("Listing emails", folder=folder, limit=limit)
            # ... implementation
            self.logger.log_outlook_operation("list_emails", True, 0.5)

# In protocol handlers
class MCPProtocolHandler:
    def __init__(self):
        self.logger = get_logger(__name__)
    
    def handle_request(self, request):
        self.logger.log_mcp_request(request.id, request.method, request.params)
        # ... process request
        self.logger.log_mcp_response(request.id, request.method, True, duration)
```

## Testing

The logging system includes comprehensive tests covering:

- JSON formatting and structure
- Exception handling and formatting
- Performance logging functionality
- Log rotation behavior
- Configuration validation
- Integration scenarios

Run tests with:

```bash
python -m pytest tests/test_logging.py -v
```

## Example Usage

See `examples/logging_demo.py` for a complete demonstration of the logging system capabilities.

## Troubleshooting

### Common Issues

1. **Permission Errors**: Ensure the log directory is writable
2. **File Locks**: On Windows, ensure no other processes have log files open
3. **Disk Space**: Monitor disk usage, especially with high log volumes
4. **Performance**: Adjust log levels in production to balance detail vs. performance

### Debug Mode

Enable debug logging to see detailed system information:

```python
configure_logging(log_level="DEBUG")
```

This will log additional details about internal operations, method calls, and system state.