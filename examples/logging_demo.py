#!/usr/bin/env python3
"""
Demonstration of the Outlook MCP Server logging system.

This script shows how to use the comprehensive logging system with structured JSON output,
performance metrics, and different log levels.
"""

import sys
import time
from pathlib import Path

# Add the src directory to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.logging import configure_logging, get_logger


def main():
    """Demonstrate the logging system functionality."""
    
    # Configure logging with structured JSON output
    log_dir = Path("demo_logs")
    configure_logging(
        log_level="DEBUG",
        log_dir=str(log_dir),
        max_bytes=1024 * 1024,  # 1MB
        backup_count=3,
        console_output=True
    )
    
    # Get logger instances
    main_logger = get_logger("demo.main")
    service_logger = get_logger("demo.service")
    
    print("=== Logging System Demo ===")
    print(f"Logs will be written to: {log_dir.absolute()}")
    print()
    
    # Basic logging examples
    print("1. Basic logging messages:")
    main_logger.info("Application started")
    main_logger.debug("Debug information", user_id="demo-user", session_id="sess-123")
    main_logger.warning("This is a warning message")
    
    # MCP request/response logging
    print("2. MCP protocol logging:")
    main_logger.log_mcp_request("req-001", "list_emails", {"folder": "inbox", "limit": 10})
    
    # Simulate processing time
    time.sleep(0.1)
    
    main_logger.log_mcp_response("req-001", "list_emails", success=True, duration=0.1)
    
    # Outlook operation logging
    print("3. Outlook operation logging:")
    service_logger.log_outlook_operation(
        "connect", 
        success=True, 
        duration=0.05,
        outlook_version="16.0",
        profile="Default"
    )
    
    service_logger.log_outlook_operation(
        "get_folders", 
        success=True, 
        duration=0.02,
        folder_count=15
    )
    
    # Connection status logging (Requirement 5.5)
    print("4. Connection status logging:")
    service_logger.log_connection_status(True, "Successfully connected to Outlook")
    
    # Performance timing with context manager
    print("5. Performance timing:")
    with main_logger.time_operation("email_search"):
        # Simulate email search operation
        time.sleep(0.05)
        service_logger.info("Searching emails", query="important", folder="inbox")
    
    # Performance metrics
    main_logger.performance.log_request_timing("manual_operation", 0.123, success=True)
    main_logger.performance.log_resource_usage(memory_mb=45.2, cpu_percent=12.5)
    
    # Error logging with exception
    print("6. Error logging:")
    try:
        raise ValueError("Simulated error for demo")
    except ValueError as e:
        main_logger.error("An error occurred during processing", exc_info=True, 
                         operation="demo", error_code="DEMO_001")
    
    # Structured logging with custom fields
    print("7. Structured logging with custom fields:")
    service_logger.info(
        "Email operation completed",
        operation_type="list_emails",
        folder="inbox",
        email_count=25,
        processing_time_ms=150,
        user_context={
            "user_id": "demo-user",
            "permissions": ["read", "list"],
            "client_version": "1.0.0"
        }
    )
    
    print()
    print("Demo completed! Check the log files for structured JSON output.")
    print(f"Log file location: {log_dir / 'outlook_mcp_server.log'}")
    
    # Show a sample of the log content
    log_file = log_dir / "outlook_mcp_server.log"
    if log_file.exists():
        print("\nSample log entries:")
        print("-" * 50)
        with open(log_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            for i, line in enumerate(lines[-3:], 1):  # Show last 3 entries
                print(f"Entry {i}: {line.strip()}")


if __name__ == "__main__":
    main()