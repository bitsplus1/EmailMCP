#!/usr/bin/env python3
"""
Command-line interface for EmailMCP server.
"""

import argparse
import asyncio
import logging
import sys
from pathlib import Path

from email_mcp.server import main as run_server
from email_mcp.config import EmailMCPConfig


def setup_logging(level: str = "INFO"):
    """Set up logging configuration."""
    logging.basicConfig(
        level=getattr(logging, level.upper()),
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )


def create_config_file(config_path: str):
    """Create a sample configuration file."""
    config = EmailMCPConfig()
    config.save_to_file(config_path)
    print(f"Created sample configuration file at {config_path}")
    print("Please edit the configuration file with your email settings.")


async def test_connection():
    """Test the email connection with current configuration."""
    from email_mcp.server import EmailMCPServer
    
    server = EmailMCPServer()
    try:
        if server.email_client:
            connected = await server.email_client.connect()
            if connected:
                print("✓ Email connection successful")
                folders = await server.email_client.list_folders()
                print(f"✓ Found {len(folders)} folders")
                await server.email_client.disconnect()
            else:
                print("✗ Email connection failed")
                return 1
        else:
            print("✗ No email client configured")
            return 1
    except Exception as e:
        print(f"✗ Email connection error: {e}")
        return 1
    
    return 0


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="EmailMCP - Model Context Protocol server for email",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  email-mcp                    # Start the MCP server
  email-mcp --test             # Test email connection
  email-mcp --create-config    # Create sample config file
  email-mcp --log-level DEBUG  # Start with debug logging
        """
    )
    
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="Set logging level (default: INFO)"
    )
    
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration file (JSON format)"
    )
    
    parser.add_argument(
        "--create-config",
        type=str,
        metavar="PATH",
        help="Create a sample configuration file at the specified path"
    )
    
    parser.add_argument(
        "--test",
        action="store_true",
        help="Test email connection and exit"
    )
    
    args = parser.parse_args()
    
    setup_logging(args.log_level)
    
    if args.create_config:
        create_config_file(args.create_config)
        return 0
    
    if args.test:
        return asyncio.run(test_connection())
    
    # Load custom config if provided
    if args.config:
        config_path = Path(args.config)
        if not config_path.exists():
            print(f"Error: Configuration file {args.config} not found")
            return 1
        
        # Import and override global config
        from email_mcp import config
        config.__dict__.update(EmailMCPConfig.load_from_file(args.config).__dict__)
    
    # Start the MCP server
    try:
        asyncio.run(run_server())
    except KeyboardInterrupt:
        print("\nServer stopped by user")
    except Exception as e:
        logging.error(f"Server error: {e}")
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())