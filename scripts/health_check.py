#!/usr/bin/env python3
"""
Health check script for Outlook MCP Server monitoring.

This script provides a standardized health check interface for monitoring
systems, load balancers, and deployment tools. It returns appropriate
exit codes and JSON output for automated monitoring.

Usage:
    python scripts/health_check.py [options]

Exit Codes:
    0 - Healthy (all checks passed)
    1 - Unhealthy (critical checks failed)
    2 - Error (script execution failed)

Output Format:
    JSON object with health status information
"""

import asyncio
import json
import sys
import argparse
import time
from pathlib import Path
from typing import Dict, Any, Optional

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

try:
    from outlook_mcp_server.health import get_health_status, is_server_healthy
    from outlook_mcp_server.server import OutlookMCPServer, create_server_config
    from outlook_mcp_server.logging.logger import get_logger
except ImportError as e:
    print(json.dumps({
        "status": "error",
        "message": f"Failed to import required modules: {e}",
        "timestamp": time.time()
    }))
    sys.exit(2)


class HealthCheckRunner:
    """Health check runner for monitoring systems."""
    
    def __init__(self, config_file: Optional[str] = None, timeout: int = 30):
        self.config_file = config_file
        self.timeout = timeout
        self.logger = None
        
    async def run_health_check(self) -> Dict[str, Any]:
        """
        Run comprehensive health check.
        
        Returns:
            Dictionary with health check results
        """
        start_time = time.time()
        
        try:
            # Initialize logger
            self.logger = get_logger(__name__)
            
            # Load configuration
            config = self._load_config()
            
            # Create server instance for health checking
            server = OutlookMCPServer(config)
            
            # Perform health check with timeout
            health_status = await asyncio.wait_for(
                get_health_status(server),
                timeout=self.timeout
            )
            
            # Convert to dictionary for JSON output
            result = {
                "status": health_status.status,
                "healthy": health_status.status == "healthy",
                "timestamp": health_status.timestamp,
                "uptime_seconds": health_status.uptime_seconds,
                "outlook_connected": health_status.outlook_connected,
                "server_running": health_status.server_running,
                "check_duration": time.time() - start_time,
                "checks": health_status.checks,
                "metrics": health_status.metrics
            }
            
            return result
            
        except asyncio.TimeoutError:
            return {
                "status": "unhealthy",
                "healthy": False,
                "timestamp": time.time(),
                "error": "Health check timed out",
                "timeout_seconds": self.timeout,
                "check_duration": time.time() - start_time
            }
        except Exception as e:
            return {
                "status": "error",
                "healthy": False,
                "timestamp": time.time(),
                "error": str(e),
                "error_type": type(e).__name__,
                "check_duration": time.time() - start_time
            }
    
    async def run_quick_check(self) -> Dict[str, Any]:
        """
        Run a quick health check (just connectivity).
        
        Returns:
            Dictionary with basic health status
        """
        start_time = time.time()
        
        try:
            # Load configuration
            config = self._load_config()
            
            # Create server instance
            server = OutlookMCPServer(config)
            
            # Quick health check
            is_healthy = await asyncio.wait_for(
                is_server_healthy(server),
                timeout=self.timeout
            )
            
            return {
                "status": "healthy" if is_healthy else "unhealthy",
                "healthy": is_healthy,
                "timestamp": time.time(),
                "check_type": "quick",
                "check_duration": time.time() - start_time
            }
            
        except asyncio.TimeoutError:
            return {
                "status": "unhealthy",
                "healthy": False,
                "timestamp": time.time(),
                "error": "Quick health check timed out",
                "timeout_seconds": self.timeout,
                "check_duration": time.time() - start_time
            }
        except Exception as e:
            return {
                "status": "error",
                "healthy": False,
                "timestamp": time.time(),
                "error": str(e),
                "error_type": type(e).__name__,
                "check_duration": time.time() - start_time
            }
    
    def _load_config(self) -> Dict[str, Any]:
        """Load server configuration."""
        if self.config_file:
            config_path = Path(self.config_file)
            if not config_path.exists():
                raise FileNotFoundError(f"Configuration file not found: {self.config_file}")
            
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    file_config = json.load(f)
                
                # Merge with default config
                config = create_server_config()
                config.update(file_config)
                return config
                
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in configuration file: {e}")
        else:
            # Use default configuration
            return create_server_config()
    
    def get_exit_code(self, result: Dict[str, Any]) -> int:
        """
        Get appropriate exit code based on health check result.
        
        Args:
            result: Health check result dictionary
            
        Returns:
            Exit code (0=healthy, 1=unhealthy, 2=error)
        """
        status = result.get("status", "error")
        
        if status == "healthy":
            return 0
        elif status in ["unhealthy", "degraded"]:
            return 1
        else:  # error or unknown
            return 2


async def main():
    """Main entry point for health check script."""
    parser = argparse.ArgumentParser(
        description="Outlook MCP Server Health Check",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exit Codes:
  0    Server is healthy
  1    Server is unhealthy or degraded
  2    Health check failed or error occurred

Output:
  JSON object with health status information

Examples:
  %(prog)s                           Basic health check
  %(prog)s --config prod.json        Health check with custom config
  %(prog)s --quick                   Quick connectivity check only
  %(prog)s --timeout 60              Health check with 60 second timeout
  %(prog)s --format nagios           Nagios-compatible output format
        """
    )
    
    parser.add_argument(
        "--config",
        type=str,
        help="Path to configuration file"
    )
    
    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="Health check timeout in seconds (default: 30)"
    )
    
    parser.add_argument(
        "--quick",
        action="store_true",
        help="Perform quick health check (connectivity only)"
    )
    
    parser.add_argument(
        "--format",
        choices=["json", "nagios", "simple"],
        default="json",
        help="Output format (default: json)"
    )
    
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Include detailed check information"
    )
    
    args = parser.parse_args()
    
    # Create health check runner
    runner = HealthCheckRunner(
        config_file=args.config,
        timeout=args.timeout
    )
    
    try:
        # Run appropriate health check
        if args.quick:
            result = await runner.run_quick_check()
        else:
            result = await runner.run_health_check()
        
        # Format output
        if args.format == "json":
            # JSON output (default)
            if not args.verbose:
                # Remove detailed information for cleaner output
                result.pop("checks", None)
                result.pop("metrics", None)
            
            print(json.dumps(result, indent=2))
            
        elif args.format == "nagios":
            # Nagios-compatible output
            status = result["status"].upper()
            message = result.get("error", f"Server is {result['status']}")
            
            if result.get("outlook_connected") is not None:
                outlook_status = "connected" if result["outlook_connected"] else "disconnected"
                message += f" (Outlook: {outlook_status})"
            
            print(f"{status} - {message}")
            
        elif args.format == "simple":
            # Simple text output
            status = result["status"]
            healthy = result["healthy"]
            
            print(f"Status: {status}")
            print(f"Healthy: {'Yes' if healthy else 'No'}")
            
            if "error" in result:
                print(f"Error: {result['error']}")
            
            if "outlook_connected" in result:
                outlook_status = "Yes" if result["outlook_connected"] else "No"
                print(f"Outlook Connected: {outlook_status}")
        
        # Exit with appropriate code
        exit_code = runner.get_exit_code(result)
        sys.exit(exit_code)
        
    except KeyboardInterrupt:
        print(json.dumps({
            "status": "error",
            "message": "Health check interrupted by user",
            "timestamp": time.time()
        }))
        sys.exit(2)
    except Exception as e:
        print(json.dumps({
            "status": "error",
            "message": f"Health check script failed: {e}",
            "error_type": type(e).__name__,
            "timestamp": time.time()
        }))
        sys.exit(2)


if __name__ == "__main__":
    asyncio.run(main())