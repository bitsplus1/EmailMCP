#!/usr/bin/env python3
"""
Deployment and configuration examples for Outlook MCP Server.

This script demonstrates various deployment scenarios and configuration
options for the Outlook MCP Server in different environments.
"""

import asyncio
import json
import os
import sys
from pathlib import Path
from typing import Dict, Any

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.health import get_health_status, is_server_healthy
from outlook_mcp_server.logging.logger import get_logger


class DeploymentExamples:
    """Examples of different deployment configurations and scenarios."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
    
    def create_development_config(self) -> Dict[str, Any]:
        """Create a development configuration."""
        return {
            "log_level": "DEBUG",
            "log_dir": "logs",
            "max_concurrent_requests": 3,
            "request_timeout": 30,
            "outlook_connection_timeout": 10,
            "enable_performance_logging": True,
            "enable_console_output": True,
            "server_mode": "standalone",
            "health_check_interval": 15,
            "rate_limiting": {
                "enabled": False,
                "requests_per_minute": 60
            },
            "caching": {
                "enabled": True,
                "email_cache_ttl": 60,
                "max_cache_size_mb": 50
            },
            "monitoring": {
                "enabled": True,
                "metrics_interval": 30
            }
        }
    
    def create_production_config(self) -> Dict[str, Any]:
        """Create a production configuration."""
        return {
            "log_level": "INFO",
            "log_dir": "/var/log/outlook-mcp-server",
            "max_concurrent_requests": 20,
            "request_timeout": 45,
            "outlook_connection_timeout": 15,
            "enable_performance_logging": True,
            "enable_console_output": False,
            "server_mode": "stdio",
            "health_check_interval": 60,
            "rate_limiting": {
                "enabled": True,
                "requests_per_minute": 100,
                "burst_size": 20
            },
            "caching": {
                "enabled": True,
                "email_cache_ttl": 300,
                "max_cache_size_mb": 200
            },
            "monitoring": {
                "enabled": True,
                "metrics_interval": 300,
                "health_check_endpoint": True
            },
            "security": {
                "validate_requests": True,
                "sanitize_responses": True,
                "max_request_size": 1048576,
                "allowed_folders": ["Inbox", "Sent Items", "Drafts"]
            }
        }
    
    def create_high_performance_config(self) -> Dict[str, Any]:
        """Create a high-performance configuration for busy environments."""
        return {
            "log_level": "WARNING",  # Reduce logging overhead
            "log_dir": "/var/log/outlook-mcp-server",
            "max_concurrent_requests": 50,
            "request_timeout": 60,
            "outlook_connection_timeout": 20,
            "enable_performance_logging": True,
            "enable_console_output": False,
            "server_mode": "stdio",
            "health_check_interval": 120,
            "rate_limiting": {
                "enabled": True,
                "requests_per_minute": 500,
                "burst_size": 100
            },
            "caching": {
                "enabled": True,
                "email_cache_ttl": 600,
                "max_cache_size_mb": 500
            },
            "performance": {
                "connection_pool_size": 10,
                "lazy_loading": True,
                "memory_management": True,
                "compression": True
            },
            "monitoring": {
                "enabled": True,
                "metrics_interval": 600
            }
        }
    
    def save_config_file(self, config: Dict[str, Any], filename: str) -> None:
        """Save configuration to a file."""
        config_path = Path(filename)
        config_path.parent.mkdir(parents=True, exist_ok=True)
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
        
        print(f"✅ Configuration saved to: {config_path}")
    
    async def test_configuration(self, config: Dict[str, Any]) -> bool:
        """Test a configuration by creating and starting a server."""
        print("🔍 Testing configuration...")
        
        server = None
        try:
            # Create server with configuration
            server = OutlookMCPServer(config)
            
            # Start server
            await server.start()
            
            # Check health
            is_healthy = await is_server_healthy(server)
            
            if is_healthy:
                print("✅ Configuration test passed - server is healthy")
                
                # Get detailed health status
                health_status = await get_health_status(server)
                print(f"   Status: {health_status.status}")
                print(f"   Outlook Connected: {health_status.outlook_connected}")
                
                return True
            else:
                print("❌ Configuration test failed - server is not healthy")
                return False
                
        except Exception as e:
            print(f"❌ Configuration test failed: {e}")
            return False
        finally:
            if server:
                await server.stop()
    
    async def demonstrate_environment_loading(self) -> None:
        """Demonstrate loading configuration from environment variables."""
        print("\n📋 Environment Variable Configuration Example")
        print("=" * 50)
        
        # Set example environment variables
        env_vars = {
            "OUTLOOK_MCP_LOG_LEVEL": "DEBUG",
            "OUTLOOK_MCP_LOG_DIR": "logs/env-example",
            "OUTLOOK_MCP_MAX_CONCURRENT": "5",
            "OUTLOOK_MCP_SERVER_MODE": "standalone"
        }
        
        # Temporarily set environment variables
        original_env = {}
        for key, value in env_vars.items():
            original_env[key] = os.environ.get(key)
            os.environ[key] = value
        
        try:
            # Create configuration that will use environment variables
            config = create_server_config()
            
            # The start_server.py script would override these with env vars
            # For demonstration, we'll manually apply them
            config["log_level"] = os.environ["OUTLOOK_MCP_LOG_LEVEL"]
            config["log_dir"] = os.environ["OUTLOOK_MCP_LOG_DIR"]
            config["max_concurrent_requests"] = int(os.environ["OUTLOOK_MCP_MAX_CONCURRENT"])
            config["server_mode"] = os.environ["OUTLOOK_MCP_SERVER_MODE"]
            
            print("Environment variables set:")
            for key, value in env_vars.items():
                print(f"  {key}={value}")
            
            print(f"\nResulting configuration:")
            print(f"  log_level: {config['log_level']}")
            print(f"  log_dir: {config['log_dir']}")
            print(f"  max_concurrent_requests: {config['max_concurrent_requests']}")
            print(f"  server_mode: {config.get('server_mode', 'stdio')}")
            
        finally:
            # Restore original environment
            for key, original_value in original_env.items():
                if original_value is None:
                    os.environ.pop(key, None)
                else:
                    os.environ[key] = original_value
    
    async def demonstrate_health_monitoring(self) -> None:
        """Demonstrate health monitoring capabilities."""
        print("\n🏥 Health Monitoring Example")
        print("=" * 50)
        
        # Create a test server
        config = self.create_development_config()
        server = OutlookMCPServer(config)
        
        try:
            await server.start()
            
            # Get comprehensive health status
            health_status = await get_health_status(server)
            
            print("Health Status Report:")
            print(f"  Overall Status: {health_status.status}")
            print(f"  Timestamp: {health_status.timestamp}")
            print(f"  Uptime: {health_status.uptime_seconds:.1f} seconds")
            print(f"  Server Running: {health_status.server_running}")
            print(f"  Outlook Connected: {health_status.outlook_connected}")
            
            print("\nDetailed Checks:")
            for check_name, check_result in health_status.checks.items():
                status_icon = "✅" if check_result["status"] == "pass" else "⚠️" if check_result["status"] == "warn" else "❌"
                print(f"  {status_icon} {check_name}: {check_result['message']}")
            
            print("\nPerformance Metrics:")
            for metric_name, metric_value in health_status.metrics.items():
                if isinstance(metric_value, (int, float)):
                    print(f"  📊 {metric_name}: {metric_value}")
            
        except Exception as e:
            print(f"❌ Health monitoring example failed: {e}")
        finally:
            if server:
                await server.stop()
    
    def create_environment_file(self, environment: str = "production") -> None:
        """Create an environment file for the specified environment."""
        if environment == "development":
            env_content = """# Development Environment Configuration
OUTLOOK_MCP_LOG_LEVEL=DEBUG
OUTLOOK_MCP_LOG_DIR=logs
OUTLOOK_MCP_MAX_CONCURRENT=5
OUTLOOK_MCP_REQUEST_TIMEOUT=30
OUTLOOK_MCP_CONNECTION_TIMEOUT=10
OUTLOOK_MCP_PERFORMANCE_LOGGING=true
OUTLOOK_MCP_CONSOLE_OUTPUT=true
OUTLOOK_MCP_SERVER_MODE=standalone
OUTLOOK_MCP_HEALTH_CHECK_INTERVAL=30
"""
        elif environment == "production":
            env_content = """# Production Environment Configuration
OUTLOOK_MCP_LOG_LEVEL=INFO
OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp-server
OUTLOOK_MCP_MAX_CONCURRENT=20
OUTLOOK_MCP_REQUEST_TIMEOUT=45
OUTLOOK_MCP_CONNECTION_TIMEOUT=15
OUTLOOK_MCP_PERFORMANCE_LOGGING=true
OUTLOOK_MCP_CONSOLE_OUTPUT=false
OUTLOOK_MCP_SERVER_MODE=stdio
OUTLOOK_MCP_HEALTH_CHECK_INTERVAL=60
"""
        else:
            raise ValueError(f"Unknown environment: {environment}")
        
        filename = f".env.{environment}"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(env_content)
        
        print(f"✅ Environment file created: {filename}")
    
    async def run_all_examples(self) -> None:
        """Run all deployment examples."""
        print("🚀 Outlook MCP Server - Deployment Examples")
        print("=" * 60)
        
        # Create configuration examples
        print("\n📝 Creating Configuration Examples")
        print("-" * 40)
        
        configs = {
            "config/development.json": self.create_development_config(),
            "config/production.json": self.create_production_config(),
            "config/high-performance.json": self.create_high_performance_config()
        }
        
        for filename, config in configs.items():
            self.save_config_file(config, filename)
        
        # Create environment files
        print("\n🌍 Creating Environment Files")
        print("-" * 40)
        
        self.create_environment_file("development")
        self.create_environment_file("production")
        
        # Demonstrate environment loading
        await self.demonstrate_environment_loading()
        
        # Test a configuration
        print("\n🧪 Testing Development Configuration")
        print("-" * 40)
        
        dev_config = self.create_development_config()
        await self.test_configuration(dev_config)
        
        # Demonstrate health monitoring
        await self.demonstrate_health_monitoring()
        
        print("\n✅ All deployment examples completed!")
        print("\nNext steps:")
        print("1. Review the created configuration files in config/")
        print("2. Copy and customize .env.development or .env.production")
        print("3. Test your configuration with: python start_server.py --test-connection")
        print("4. Start the server with: python start_server.py --config config/production.json")


async def main():
    """Main function to run deployment examples."""
    examples = DeploymentExamples()
    
    try:
        await examples.run_all_examples()
    except KeyboardInterrupt:
        print("\n⏹️ Examples interrupted by user")
    except Exception as e:
        print(f"\n❌ Examples failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())