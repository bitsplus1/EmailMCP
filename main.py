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
from outlook_mcp_server.http_server import run_http_server
from outlook_mcp_server.main import main as run_interactive_server, test_outlook_connection, create_sample_config
from outlook_mcp_server.server import create_server_config, OutlookMCPServer


async def handle_single_request(config):
    """Handle a single MCP request from stdin and exit."""
    import json
    import sys
    
    try:
        # Read JSON request from stdin
        input_line = sys.stdin.readline().strip()
        if not input_line:
            print('{"jsonrpc": "2.0", "id": null, "error": {"code": -32700, "message": "No input provided"}}')
            return
        
        # Parse JSON request
        try:
            request_data = json.loads(input_line)
        except json.JSONDecodeError as e:
            print(f'{{"jsonrpc": "2.0", "id": null, "error": {{"code": -32700, "message": "Parse error: {str(e)}"}}}}')
            return
        
        # Create and start server
        server = OutlookMCPServer(config)
        await server.start()
        
        try:
            # Handle the request
            response = await server.handle_request(request_data)
            
            # Output response
            print(json.dumps(response, ensure_ascii=False))
            
        finally:
            # Clean up
            await server.stop()
            
    except Exception as e:
        error_response = {
            "jsonrpc": "2.0",
            "id": request_data.get("id") if 'request_data' in locals() else None,
            "error": {
                "code": -32603,
                "message": f"Internal error: {str(e)}"
            }
        }
        print(json.dumps(error_response))


def main():
    """Main entry point with mode selection."""
    parser = argparse.ArgumentParser(description="Outlook MCP Server")
    parser.add_argument(
        "mode",
        nargs="?",
        default="stdio",
        choices=["stdio", "http", "interactive", "test", "create-config", "single-request"],
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
        elif args.mode == "http":
            # Run as MCP HTTP server (remote access mode)
            asyncio.run(run_http_server(config))
        elif args.mode == "interactive":
            # Run in interactive mode with console output
            asyncio.run(run_interactive_server())
        elif args.mode == "test":
            # Test connection and exit
            from outlook_mcp_server.logging.logger import get_logger
            logger = get_logger(__name__)
            asyncio.run(test_outlook_connection(config, logger))
        elif args.mode == "single-request":
            # Process a single request from stdin and exit
            asyncio.run(handle_single_request(config))
            
    except KeyboardInterrupt:
        print("\nServer stopped by user")
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()