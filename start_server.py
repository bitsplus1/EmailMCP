#!/usr/bin/env python3
"""
Production startup script for Outlook MCP Server.

This script provides a production-ready way to start the Outlook MCP Server
with proper configuration loading, environment variable support, health checks,
and graceful shutdown handling.

Usage:
    python start_server.py [options]
    
Environment Variables:
    OUTLOOK_MCP_CONFIG_FILE    Path to configuration file
    OUTLOOK_MCP_LOG_LEVEL      Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    OUTLOOK_MCP_LOG_DIR        Directory for log files
    OUTLOOK_MCP_MAX_CONCURRENT Maximum concurrent requests
    OUTLOOK_MCP_PORT           Port for health check endpoint (optional)
    OUTLOOK_MCP_HOST           Host for health check endpoint (default: localhost)

Examples:
    # Start with default configuration
    python start_server.py
    
    # Start with custom config file
    python start_server.py --config production_config.json
    
    # Start with environment variables
    export OUTLOOK_MCP_LOG_LEVEL=DEBUG
    export OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp
    python start_server.py
    
    # Start as a service (no console output)
    python start_server.py --service-mode
"""

import os
import sys
import json
import signal
import asyncio
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.mcp_stdio_server import run_stdio_server
from outlook_mcp_server.logging.logger import get_logger


class ProductionServer:
    """Production server wrapper with enhanced features."""
    
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.server: Optional[OutlookMCPServer] = None
        self.logger = get_logger(__name__)
        self.shutdown_requested = False
        
    async def start(self) -> None:
        """Start the production server."""
        try:
            self.logger.info("Starting Outlook MCP Server in production mode")
            
            # Setup signal handlers
            self._setup_signal_handlers()
            
            # Create and start server
            self.server = OutlookMCPServer(self.config)
            await self.server.start()
            
            # Log startup success
            server_info = self.server.get_server_info()
            self.logger.info("Server started successfully", server_info=server_info)
            
            # Run main server loop
            await self._run_server_loop()
            
        except Exception as e:
            self.logger.error(f"Failed to start server: {e}", exc_info=True)
            raise
    
    async def stop(self) -> None:
        """Stop the production server."""
        if self.server:
            self.logger.info("Stopping server")
            await self.server.stop()
            self.server = None
    
    def _setup_signal_handlers(self) -> None:
        """Setup signal handlers for graceful shutdown."""
        def signal_handler(signum, frame):
            self.logger.info(f"Received signal {signum}, initiating graceful shutdown")
            self.shutdown_requested = True
        
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)
        
        if hasattr(signal, 'SIGBREAK'):  # Windows
            signal.signal(signal.SIGBREAK, signal_handler)
    
    async def _run_server_loop(self) -> None:
        """Run the main server loop."""
        try:
            if self.config.get("server_mode") == "stdio":
                # Run as MCP stdio server
                await run_stdio_server(self.config)
            else:
                # Run as standalone server with health checks
                await self._run_standalone_loop()
                
        except asyncio.CancelledError:
            self.logger.info("Server loop cancelled")
        except Exception as e:
            self.logger.error(f"Server loop error: {e}", exc_info=True)
            raise
    
    async def _run_standalone_loop(self) -> None:
        """Run standalone server loop with periodic health checks."""
        health_check_interval = self.config.get("health_check_interval", 30)
        
        while not self.shutdown_requested and self.server and self.server.is_running():
            try:
                # Wait for shutdown or health check interval
                await asyncio.sleep(health_check_interval)
                
                # Perform health check
                if self.server and not self.server.is_healthy():
                    self.logger.warning("Server health check failed")
                    # Could implement reconnection logic here
                
            except asyncio.CancelledError:
                break
            except Exception as e:
                self.logger.error(f"Error in server loop: {e}", exc_info=True)
                break
    
    def get_health_status(self) -> Dict[str, Any]:
        """Get server health status."""
        if not self.server:
            return {
                "status": "stopped",
                "healthy": False,
                "timestamp": datetime.utcnow().isoformat()
            }
        
        stats = self.server.get_server_stats()
        return {
            "status": "running" if self.server.is_running() else "stopped",
            "healthy": self.server.is_healthy(),
            "stats": stats,
            "timestamp": datetime.utcnow().isoformat()
        }


def load_configuration() -> Dict[str, Any]:
    """Load configuration from file and environment variables."""
    # Start with default configuration
    config = create_server_config()
    
    # Load from config file (environment variable or command line)
    config_file = os.getenv("OUTLOOK_MCP_CONFIG_FILE")
    if config_file:
        config_path = Path(config_file)
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    file_config = json.load(f)
                config.update(file_config)
                print(f"✅ Loaded configuration from: {config_path}")
            except Exception as e:
                print(f"❌ Error loading config file {config_file}: {e}")
                sys.exit(1)
        else:
            print(f"❌ Config file not found: {config_file}")
            sys.exit(1)
    
    # Override with environment variables
    env_mappings = {
        "OUTLOOK_MCP_LOG_LEVEL": "log_level",
        "OUTLOOK_MCP_LOG_DIR": "log_dir",
        "OUTLOOK_MCP_MAX_CONCURRENT": "max_concurrent_requests",
        "OUTLOOK_MCP_REQUEST_TIMEOUT": "request_timeout",
        "OUTLOOK_MCP_CONNECTION_TIMEOUT": "outlook_connection_timeout",
        "OUTLOOK_MCP_PERFORMANCE_LOGGING": "enable_performance_logging",
        "OUTLOOK_MCP_CONSOLE_OUTPUT": "enable_console_output",
        "OUTLOOK_MCP_HEALTH_CHECK_INTERVAL": "health_check_interval",
        "OUTLOOK_MCP_SERVER_MODE": "server_mode"
    }
    
    for env_var, config_key in env_mappings.items():
        env_value = os.getenv(env_var)
        if env_value is not None:
            # Convert string values to appropriate types
            if config_key in ["max_concurrent_requests", "request_timeout", 
                             "outlook_connection_timeout", "health_check_interval"]:
                try:
                    config[config_key] = int(env_value)
                except ValueError:
                    print(f"❌ Invalid integer value for {env_var}: {env_value}")
                    sys.exit(1)
            elif config_key in ["enable_performance_logging", "enable_console_output"]:
                config[config_key] = env_value.lower() in ("true", "1", "yes", "on")
            else:
                config[config_key] = env_value
    
    return config


def validate_environment() -> None:
    """Validate the runtime environment."""
    errors = []
    
    # Check Python version
    if sys.version_info < (3, 8):
        errors.append("Python 3.8 or higher is required")
    
    # Check for required modules
    try:
        import win32com.client
    except ImportError:
        errors.append("pywin32 package is required (pip install pywin32)")
    
    # Check if running on Windows
    if os.name != 'nt':
        errors.append("This server requires Windows to access Outlook COM interface")
    
    if errors:
        print("❌ Environment validation failed:")
        for error in errors:
            print(f"   • {error}")
        sys.exit(1)


def create_pid_file(pid_file_path: str) -> None:
    """Create a PID file for process management."""
    try:
        pid_path = Path(pid_file_path)
        pid_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(pid_path, 'w') as f:
            f.write(str(os.getpid()))
        
        print(f"✅ PID file created: {pid_path}")
        
    except Exception as e:
        print(f"⚠️  Warning: Could not create PID file: {e}")


def remove_pid_file(pid_file_path: str) -> None:
    """Remove the PID file."""
    try:
        pid_path = Path(pid_file_path)
        if pid_path.exists():
            pid_path.unlink()
    except Exception:
        pass  # Ignore errors when removing PID file


async def main() -> None:
    """Main entry point for production server."""
    parser = argparse.ArgumentParser(
        description="Outlook MCP Server - Production Startup Script",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Environment Variables:
  OUTLOOK_MCP_CONFIG_FILE      Path to configuration file
  OUTLOOK_MCP_LOG_LEVEL        Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
  OUTLOOK_MCP_LOG_DIR          Directory for log files
  OUTLOOK_MCP_MAX_CONCURRENT   Maximum concurrent requests
  OUTLOOK_MCP_SERVER_MODE      Server mode (stdio, standalone)

Examples:
  %(prog)s                           Start with default configuration
  %(prog)s --config prod.json        Start with custom config file
  %(prog)s --service-mode             Start in service mode (no console output)
  %(prog)s --test-connection          Test Outlook connection and exit
        """
    )
    
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration file (overrides OUTLOOK_MCP_CONFIG_FILE)"
    )
    
    parser.add_argument(
        "--service-mode",
        action="store_true",
        help="Run in service mode (disable console output, create PID file)"
    )
    
    parser.add_argument(
        "--pid-file",
        type=str,
        default="outlook_mcp_server.pid",
        help="Path to PID file (default: outlook_mcp_server.pid)"
    )
    
    parser.add_argument(
        "--test-connection",
        action="store_true",
        help="Test Outlook connection and exit"
    )
    
    parser.add_argument(
        "--validate-config",
        action="store_true",
        help="Validate configuration and exit"
    )
    
    args = parser.parse_args()
    
    # Validate environment
    validate_environment()
    
    # Override config file from command line
    if args.config:
        os.environ["OUTLOOK_MCP_CONFIG_FILE"] = args.config
    
    # Load configuration
    config = load_configuration()
    
    # Apply service mode settings
    if args.service_mode:
        config["enable_console_output"] = False
        config["server_mode"] = "standalone"
    
    # Set default server mode if not specified
    if "server_mode" not in config:
        config["server_mode"] = "stdio"
    
    # Validate configuration only
    if args.validate_config:
        print("✅ Configuration validation passed")
        print(f"Configuration: {json.dumps(config, indent=2)}")
        return
    
    # Test connection only
    if args.test_connection:
        from outlook_mcp_server.main import test_outlook_connection
        logger = get_logger(__name__)
        await test_outlook_connection(config, logger)
        return
    
    # Create PID file in service mode
    pid_file_path = None
    if args.service_mode:
        pid_file_path = args.pid_file
        create_pid_file(pid_file_path)
    
    # Create and start production server
    server = ProductionServer(config)
    
    try:
        await server.start()
    except KeyboardInterrupt:
        print("\n⏹️  Server stopped by user")
    except Exception as e:
        print(f"❌ Server failed: {e}")
        sys.exit(1)
    finally:
        # Cleanup
        await server.stop()
        if pid_file_path:
            remove_pid_file(pid_file_path)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n⏹️  Startup interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"❌ Fatal error: {e}")
        sys.exit(1)