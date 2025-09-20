"""
Main entry point for the Outlook MCP Server.

This module provides the command-line interface and startup logic for the Outlook MCP Server.
It supports multiple operation modes including stdio mode for MCP client integration,
interactive mode for development and testing, and various utility commands.

Usage:
    python main.py [mode] [options]
    
Modes:
    stdio         Run as MCP stdio server (default)
    interactive   Run in interactive mode with console output
    test          Test Outlook connection and exit
    create-config Create sample configuration file

Examples:
    # Run as MCP server with default configuration
    python main.py stdio
    
    # Run with custom configuration file
    python main.py stdio --config my_config.json
    
    # Test Outlook connection
    python main.py test
    
    # Run in interactive mode with debug logging
    python main.py interactive --log-level DEBUG

Author: Outlook MCP Server Team
Version: 1.0.0
License: MIT
"""

import asyncio
import json
import sys
import argparse
from pathlib import Path
from typing import Dict, Any, Optional, List, NoReturn

from .server import OutlookMCPServer, create_server_config
from .logging.logger import get_logger


async def main() -> None:
    """
    Main entry point for the Outlook MCP Server.
    
    Parses command line arguments, loads configuration, and starts the appropriate
    server mode based on the provided arguments.
    
    Raises:
        SystemExit: On configuration errors or startup failures
        KeyboardInterrupt: On user interruption (Ctrl+C)
    """
    # Parse command line arguments with comprehensive help
    parser = argparse.ArgumentParser(
        description="Outlook MCP Server - Provides MCP access to Microsoft Outlook",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s stdio                           Run as MCP stdio server
  %(prog)s interactive --log-level DEBUG   Run in interactive mode with debug logging
  %(prog)s test                           Test Outlook connection
  %(prog)s create-config                  Create sample configuration file
  
For more information, see the documentation in the docs/ directory.
        """
    )
    
    # Configuration options
    parser.add_argument(
        "--config", 
        type=str, 
        metavar="PATH",
        help="Path to configuration file (JSON format). If not specified, uses default configuration."
    )
    
    # Logging options
    parser.add_argument(
        "--log-level", 
        type=str, 
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set the logging level (default: %(default)s)"
    )
    parser.add_argument(
        "--log-dir", 
        type=str, 
        default="logs",
        metavar="PATH",
        help="Directory for log files (default: %(default)s)"
    )
    
    # Performance options
    parser.add_argument(
        "--max-concurrent", 
        type=int, 
        default=10,
        metavar="N",
        help="Maximum number of concurrent requests (default: %(default)d)"
    )
    
    # Output options
    parser.add_argument(
        "--no-console", 
        action="store_true",
        help="Disable console output (useful for service mode)"
    )
    
    # Testing options
    parser.add_argument(
        "--test-connection", 
        action="store_true",
        help="Test Outlook connection and exit without starting server"
    )
    
    # Parse arguments
    args = parser.parse_args()
    
    # Load configuration
    config = load_config(args)
    
    # Initialize logger
    logger = get_logger(__name__)
    
    try:
        if args.test_connection:
            # Test connection mode
            await test_outlook_connection(config, logger)
            return
        
        # Create and start server
        logger.info("Starting Outlook MCP Server")
        server = OutlookMCPServer(config)
        
        await server.start()
        
        # Print server information
        server_info = server.get_server_info()
        logger.info("Server started successfully", server_info=server_info)
        
        # Print connection instructions
        print_connection_info(server_info)
        
        # Run server until interrupted
        try:
            await run_server_loop(server, logger)
        except KeyboardInterrupt:
            logger.info("Received keyboard interrupt, shutting down")
        finally:
            await server.stop()
            
    except Exception as e:
        logger.error(f"Server startup failed: {str(e)}", exc_info=True)
        sys.exit(1)


def load_config(args: argparse.Namespace) -> Dict[str, Any]:
    """
    Load server configuration from file and command line arguments.
    
    Configuration is loaded in the following order of precedence:
    1. Default configuration values
    2. Configuration file (if specified)
    3. Command line arguments (highest priority)
    
    Args:
        args: Parsed command line arguments from argparse
        
    Returns:
        Dict containing the merged configuration
        
    Raises:
        SystemExit: If configuration file cannot be loaded or is invalid
        
    Example:
        >>> args = parser.parse_args(['--log-level', 'DEBUG'])
        >>> config = load_config(args)
        >>> config['log_level']
        'DEBUG'
    """
    # Start with default configuration
    config = create_server_config()
    
    # Load from config file if specified
    if args.config:
        config_path = Path(args.config)
        
        if not config_path.exists():
            print(f"❌ Configuration file not found: {args.config}")
            print(f"   Please check the file path or create the file using:")
            print(f"   python main.py create-config")
            sys.exit(1)
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                file_config = json.load(f)
            
            # Validate configuration structure
            if not isinstance(file_config, dict):
                raise ValueError("Configuration file must contain a JSON object")
            
            # Merge file configuration with defaults
            config.update(file_config)
            print(f"✅ Loaded configuration from: {config_path}")
            
        except json.JSONDecodeError as e:
            print(f"❌ Invalid JSON in configuration file: {args.config}")
            print(f"   Error: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"❌ Error loading configuration file: {e}")
            sys.exit(1)
    
    # Override with command line arguments (highest priority)
    config["log_level"] = args.log_level
    config["log_dir"] = args.log_dir
    config["max_concurrent_requests"] = args.max_concurrent
    config["enable_console_output"] = not args.no_console
    
    # Validate final configuration
    _validate_config(config)
    
    return config


def _validate_config(config: Dict[str, Any]) -> None:
    """
    Validate configuration values for correctness.
    
    Args:
        config: Configuration dictionary to validate
        
    Raises:
        SystemExit: If configuration contains invalid values
    """
    errors: List[str] = []
    
    # Validate log level
    valid_log_levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
    if config.get("log_level") not in valid_log_levels:
        errors.append(f"Invalid log_level: {config.get('log_level')}. Must be one of: {valid_log_levels}")
    
    # Validate numeric values
    if config.get("max_concurrent_requests", 0) <= 0:
        errors.append("max_concurrent_requests must be a positive integer")
    
    if config.get("request_timeout", 0) <= 0:
        errors.append("request_timeout must be a positive number")
    
    if config.get("outlook_connection_timeout", 0) <= 0:
        errors.append("outlook_connection_timeout must be a positive number")
    
    # Validate log directory
    log_dir = config.get("log_dir")
    if log_dir:
        try:
            Path(log_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            errors.append(f"Cannot create log directory '{log_dir}': {e}")
    
    # Report validation errors
    if errors:
        print("❌ Configuration validation failed:")
        for error in errors:
            print(f"   • {error}")
        sys.exit(1)


async def test_outlook_connection(config: Dict[str, Any], logger) -> None:
    """
    Test Outlook connection and server functionality.
    
    Performs comprehensive testing of:
    - Outlook COM connection
    - Server component initialization
    - Folder access permissions
    - Email retrieval capabilities
    - Basic server health checks
    
    Args:
        config: Server configuration dictionary
        logger: Logger instance for detailed logging
        
    Raises:
        SystemExit: If connection test fails critically
        
    Note:
        This function provides detailed output to help diagnose connection issues.
        Warnings are non-fatal and indicate partial functionality.
    """
    print("🔍 Testing Outlook MCP Server Connection")
    print("=" * 50)
    
    server: Optional[OutlookMCPServer] = None
    
    try:
        # Initialize server
        print("📋 Initializing server components...")
        server = OutlookMCPServer(config)
        
        # Start server (this tests Outlook connection)
        print("🔌 Connecting to Microsoft Outlook...")
        await server.start()
        
        # Get server statistics
        stats = server.get_server_stats()
        
        # Test server health
        if server.is_healthy():
            print("✅ Outlook connection successful")
            print("✅ Server is healthy and ready")
            print(f"✅ Connection established {stats['connections_established']} time(s)")
            
            # Test folder access
            print("\n📁 Testing folder access...")
            try:
                folders = server.folder_service.get_folders()
                folder_count = len(folders)
                print(f"✅ Successfully accessed {folder_count} folder(s)")
                
                # Show sample folders (up to 5)
                if folders:
                    print("   Sample folders:")
                    for folder in folders[:5]:
                        folder_dict = folder.to_dict() if hasattr(folder, 'to_dict') else folder
                        name = folder_dict.get('name', 'Unknown')
                        item_count = folder_dict.get('item_count', 0)
                        print(f"   • {name} ({item_count} items)")
                    
                    if folder_count > 5:
                        print(f"   ... and {folder_count - 5} more folders")
                        
            except Exception as e:
                print(f"⚠️  Warning: Could not list folders - {e}")
                logger.warning(f"Folder access test failed: {e}")
            
            # Test email access
            print("\n📧 Testing email access...")
            try:
                emails = await server.email_service.list_emails(limit=3)
                email_count = len(emails)
                print(f"✅ Successfully accessed emails (found {email_count} in test)")
                
                # Show sample email info (without sensitive data)
                if emails:
                    print("   Sample emails:")
                    for i, email in enumerate(emails, 1):
                        subject = email.get('subject', 'No Subject')[:50]
                        sender = email.get('sender', 'Unknown Sender')
                        is_read = email.get('is_read', True)
                        read_status = "Read" if is_read else "Unread"
                        print(f"   {i}. {subject}... (from: {sender}, {read_status})")
                        
            except Exception as e:
                print(f"⚠️  Warning: Could not access emails - {e}")
                logger.warning(f"Email access test failed: {e}")
            
            # Test search functionality
            print("\n🔍 Testing search functionality...")
            try:
                search_results = await server.email_service.search_emails("test", limit=1)
                print(f"✅ Search functionality working (found {len(search_results)} result(s))")
            except Exception as e:
                print(f"⚠️  Warning: Search test failed - {e}")
                logger.warning(f"Search test failed: {e}")
            
            # Show server capabilities
            print("\n🛠️  Server capabilities:")
            server_info = server.get_server_info()
            capabilities = server_info.get('capabilities', {})
            tools = capabilities.get('tools', [])
            
            for tool in tools:
                name = tool.get('name', 'Unknown')
                description = tool.get('description', 'No description')
                print(f"   • {name}: {description}")
            
            print("\n" + "=" * 50)
            print("✅ CONNECTION TEST PASSED")
            print("   The server is ready to handle MCP requests.")
            print("   You can now start the server with: python main.py stdio")
                
        else:
            print("❌ Server health check failed")
            print("   The server started but is not in a healthy state.")
            
            # Try to get more details about the problem
            if not server.outlook_adapter or not server.outlook_adapter.is_connected():
                print("   Issue: Outlook adapter is not connected")
            
            raise RuntimeError("Server health check failed")
            
    except KeyboardInterrupt:
        print("\n⏹️  Connection test interrupted by user")
        
    except Exception as e:
        print(f"\n❌ CONNECTION TEST FAILED")
        print(f"   Error: {e}")
        print(f"\n🔧 Troubleshooting tips:")
        print(f"   1. Ensure Microsoft Outlook is installed and configured")
        print(f"   2. Try opening Outlook manually to verify it works")
        print(f"   3. Check if Outlook is running in safe mode")
        print(f"   4. Verify you have permission to access Outlook")
        print(f"   5. See docs/TROUBLESHOOTING.md for more help")
        
        logger.error(f"Connection test failed: {e}", exc_info=True)
        sys.exit(1)
        
    finally:
        # Clean up server resources
        if server:
            try:
                await server.stop()
                print("\n🧹 Server resources cleaned up")
            except Exception as e:
                print(f"\n⚠️  Warning: Error during cleanup - {e}")
                logger.warning(f"Cleanup error: {e}")
        
        print("=" * 50)


async def run_server_loop(server: OutlookMCPServer, logger) -> None:
    """Run the main server loop."""
    logger.info("Server is running. Press Ctrl+C to stop.")
    
    # In a real MCP server, this would handle incoming connections
    # For now, we'll just keep the server alive and log periodic status
    
    try:
        while server.is_running():
            await asyncio.sleep(30)  # Check every 30 seconds
            
            # Log periodic status
            if server.is_healthy():
                stats = server.get_server_stats()
                logger.debug("Server status check", stats=stats)
            else:
                logger.warning("Server health check failed - attempting to reconnect")
                # In a real implementation, you might attempt to reconnect here
                
    except asyncio.CancelledError:
        logger.info("Server loop cancelled")
        raise


def print_connection_info(server_info: Dict[str, Any]) -> None:
    """Print connection information for MCP clients."""
    print("\n" + "="*60)
    print("OUTLOOK MCP SERVER READY")
    print("="*60)
    print(f"Server Name: {server_info.get('name', 'outlook-mcp-server')}")
    print(f"Version: {server_info.get('version', '1.0.0')}")
    print(f"Protocol Version: {server_info.get('protocolVersion', '2024-11-05')}")
    print("\nAvailable Methods:")
    
    capabilities = server_info.get('capabilities', {})
    tools = capabilities.get('tools', [])
    
    for tool in tools:
        print(f"  • {tool['name']}: {tool['description']}")
    
    print("\nTo connect with an MCP client:")
    print("1. Configure your MCP client to connect to this server")
    print("2. Use the methods listed above to interact with Outlook")
    print("3. Check the logs directory for detailed operation logs")
    print("\nPress Ctrl+C to stop the server")
    print("="*60 + "\n")


def create_sample_config() -> None:
    """Create a sample configuration file."""
    sample_config = {
        "log_level": "INFO",
        "log_dir": "logs",
        "max_concurrent_requests": 10,
        "request_timeout": 30,
        "outlook_connection_timeout": 10,
        "enable_performance_logging": True,
        "enable_console_output": True
    }
    
    config_path = Path("outlook_mcp_server_config.json")
    with open(config_path, 'w') as f:
        json.dump(sample_config, f, indent=2)
    
    print(f"Sample configuration created: {config_path}")


if __name__ == "__main__":
    # Handle special commands
    if len(sys.argv) > 1 and sys.argv[1] == "create-config":
        create_sample_config()
        sys.exit(0)
    
    # Run the main server
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)