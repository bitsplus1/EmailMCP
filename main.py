#!/usr/bin/env python3
"""
Main entry point for Outlook MCP Server.

This script can run the server in different modes:
- stdio: Standard MCP stdio transport (default)
- interactive: Interactive mode with console output
- test: Test Outlook connection and exit
"""

import sys
import asyncio
import argparse
from pathlib import Path

# Add src to path so we can import the package
sys.path.insert(0, str(Path(__file__).parent / "src"))

from outlook_mcp_server.mcp_stdio_server import run_stdio_server
from outlook_mcp_server.main import main as run_interactive_server, test_outlook_connection, create_sample_config
from outlook_mcp_server.server import create_server_config


def main():
    """Main entry point with mode selection."""
    parser = argparse.ArgumentParser(description="Outlook MCP Server")
    parser.add_argument(
        "mode",
        nargs="?",
        default="stdio",
        choices=["stdio", "interactive", "test", "create-config"],
        help="Server mode (default: stdio)"
    )
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration file (JSON format)"
    )
    parser.add_argument(
        "--log-level",
        type=str,
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Logging level"
    )
    parser.add_argument(
        "--log-dir",
        type=str,
        default="logs",
        help="Directory for log files"
    )
    
    args = parser.parse_args()
    
    if args.mode == "create-config":
        create_sample_config()
        return
    
    # Create basic config
    config = create_server_config(
        log_level=args.log_level,
        log_dir=args.log_dir,
        enable_console_output=(args.mode != "stdio")
    )
    
    # Load config file if specified
    if args.config:
        import json
        config_path = Path(args.config)
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    file_config = json.load(f)
                config.update(file_config)
            except Exception as e:
                print(f"Error loading config file: {e}")
                sys.exit(1)
        else:
            print(f"Config file not found: {args.config}")
            sys.exit(1)
    
    try:
        if args.mode == "stdio":
            # Run as MCP stdio server (standard mode)
            asyncio.run(run_stdio_server(config))
        elif args.mode == "interactive":
            # Run in interactive mode with console output
            asyncio.run(run_interactive_server())
        elif args.mode == "test":
            # Test connection and exit
            from outlook_mcp_server.logging.logger import get_logger
            logger = get_logger(__name__)
            asyncio.run(test_outlook_connection(config, logger))
            
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()