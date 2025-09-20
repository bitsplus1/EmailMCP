#!/usr/bin/env python3
"""
Windows Service installer for Outlook MCP Server.

This script helps install, configure, and manage the Outlook MCP Server
as a Windows service for production deployments.

Usage:
    python scripts/install_service.py install
    python scripts/install_service.py start
    python scripts/install_service.py stop
    python scripts/install_service.py remove

Requirements:
    - pywin32 package
    - Administrator privileges
    - Windows operating system
"""

import sys
import os
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import asyncio
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(project_root / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.logging.logger import get_logger


class OutlookMCPService(win32serviceutil.ServiceFramework):
    """Windows service wrapper for Outlook MCP Server."""
    
    _svc_name_ = "OutlookMCPServer"
    _svc_display_name_ = "Outlook MCP Server"
    _svc_description_ = "Model Context Protocol server for Microsoft Outlook integration"
    
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.server = None
        self.logger = None
        socket.setdefaulttimeout(60)
    
    def SvcStop(self):
        """Stop the service."""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        
        if self.logger:
            self.logger.info("Service stop requested")
        
        # Stop the server
        if self.server:
            try:
                asyncio.create_task(self.server.stop())
            except Exception as e:
                if self.logger:
                    self.logger.error(f"Error stopping server: {e}")
    
    def SvcDoRun(self):
        """Run the service."""
        try:
            # Initialize logging
            self.logger = get_logger(__name__)
            
            # Log service start
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )
            
            self.logger.info("Outlook MCP Service starting")
            
            # Load configuration
            config = self._load_service_config()
            
            # Create and start server
            self.server = OutlookMCPServer(config)
            
            # Run the server in async context
            asyncio.run(self._run_server())
            
        except Exception as e:
            error_msg = f"Service failed to start: {e}"
            
            if self.logger:
                self.logger.error(error_msg, exc_info=True)
            
            servicemanager.LogErrorMsg(error_msg)
            
            # Report service stopped
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
    
    def _load_service_config(self):
        """Load configuration for service mode."""
        # Default service configuration
        config = create_server_config(
            log_level="INFO",
            log_dir=r"C:\ProgramData\OutlookMCPServer\logs",
            enable_console_output=False,  # No console in service mode
            max_concurrent_requests=20
        )
        
        # Try to load config file from standard locations
        config_locations = [
            r"C:\ProgramData\OutlookMCPServer\config.json",
            Path(__file__).parent.parent / "config" / "production.json",
            Path(__file__).parent.parent / "outlook_mcp_server_config.json"
        ]
        
        for config_path in config_locations:
            if Path(config_path).exists():
                try:
                    import json
                    with open(config_path, 'r') as f:
                        file_config = json.load(f)
                    config.update(file_config)
                    
                    if self.logger:
                        self.logger.info(f"Loaded service config from: {config_path}")
                    break
                except Exception as e:
                    if self.logger:
                        self.logger.warning(f"Failed to load config from {config_path}: {e}")
        
        return config
    
    async def _run_server(self):
        """Run the server asynchronously."""
        try:
            # Start the server
            await self.server.start()
            
            self.logger.info("Outlook MCP Service started successfully")
            
            # Wait for stop signal
            while True:
                # Check for stop event (non-blocking)
                result = win32event.WaitForSingleObject(self.hWaitStop, 1000)  # 1 second timeout
                
                if result == win32event.WAIT_OBJECT_0:
                    # Stop event signaled
                    break
                
                # Check server health
                if not self.server.is_healthy():
                    self.logger.warning("Server health check failed")
                    # Could implement restart logic here
                
                await asyncio.sleep(1)
            
        except Exception as e:
            self.logger.error(f"Server error: {e}", exc_info=True)
            raise
        finally:
            # Cleanup
            if self.server:
                await self.server.stop()
            
            self.logger.info("Outlook MCP Service stopped")


def install_service():
    """Install the Windows service."""
    try:
        # Install the service
        win32serviceutil.InstallService(
            OutlookMCPService._svc_reg_class_,
            OutlookMCPService._svc_name_,
            OutlookMCPService._svc_display_name_,
            description=OutlookMCPService._svc_description_
        )
        
        print(f"✅ Service '{OutlookMCPService._svc_display_name_}' installed successfully")
        print("   You can now start it with: python scripts/install_service.py start")
        print("   Or use Windows Services manager")
        
    except Exception as e:
        print(f"❌ Failed to install service: {e}")
        return False
    
    return True


def start_service():
    """Start the Windows service."""
    try:
        win32serviceutil.StartService(OutlookMCPService._svc_name_)
        print(f"✅ Service '{OutlookMCPService._svc_display_name_}' started successfully")
        
    except Exception as e:
        print(f"❌ Failed to start service: {e}")
        return False
    
    return True


def stop_service():
    """Stop the Windows service."""
    try:
        win32serviceutil.StopService(OutlookMCPService._svc_name_)
        print(f"✅ Service '{OutlookMCPService._svc_display_name_}' stopped successfully")
        
    except Exception as e:
        print(f"❌ Failed to stop service: {e}")
        return False
    
    return True


def remove_service():
    """Remove the Windows service."""
    try:
        # Stop service first if running
        try:
            stop_service()
        except:
            pass  # Ignore errors if service is not running
        
        # Remove the service
        win32serviceutil.RemoveService(OutlookMCPService._svc_name_)
        print(f"✅ Service '{OutlookMCPService._svc_display_name_}' removed successfully")
        
    except Exception as e:
        print(f"❌ Failed to remove service: {e}")
        return False
    
    return True


def service_status():
    """Check service status."""
    try:
        status = win32serviceutil.QueryServiceStatus(OutlookMCPService._svc_name_)
        status_map = {
            win32service.SERVICE_STOPPED: "Stopped",
            win32service.SERVICE_START_PENDING: "Start Pending",
            win32service.SERVICE_STOP_PENDING: "Stop Pending",
            win32service.SERVICE_RUNNING: "Running",
            win32service.SERVICE_CONTINUE_PENDING: "Continue Pending",
            win32service.SERVICE_PAUSE_PENDING: "Pause Pending",
            win32service.SERVICE_PAUSED: "Paused"
        }
        
        status_text = status_map.get(status[1], f"Unknown ({status[1]})")
        print(f"Service Status: {status_text}")
        
        return status[1]
        
    except Exception as e:
        print(f"❌ Failed to check service status: {e}")
        return None


def main():
    """Main entry point for service management."""
    if len(sys.argv) < 2:
        print("Outlook MCP Server - Windows Service Manager")
        print("=" * 50)
        print("Usage:")
        print("  python scripts/install_service.py install    Install the service")
        print("  python scripts/install_service.py start      Start the service")
        print("  python scripts/install_service.py stop       Stop the service")
        print("  python scripts/install_service.py restart    Restart the service")
        print("  python scripts/install_service.py remove     Remove the service")
        print("  python scripts/install_service.py status     Check service status")
        print("")
        print("Note: Administrator privileges are required for install/remove operations")
        return
    
    command = sys.argv[1].lower()
    
    # Check if running as administrator for install/remove operations
    if command in ['install', 'remove']:
        try:
            import ctypes
            if not ctypes.windll.shell32.IsUserAnAdmin():
                print("❌ Administrator privileges required for this operation")
                print("   Please run this script as Administrator")
                return
        except:
            pass  # Ignore if we can't check admin status
    
    if command == 'install':
        install_service()
    elif command == 'start':
        start_service()
    elif command == 'stop':
        stop_service()
    elif command == 'restart':
        print("Restarting service...")
        stop_service()
        import time
        time.sleep(2)  # Wait a moment between stop and start
        start_service()
    elif command == 'remove':
        remove_service()
    elif command == 'status':
        service_status()
    else:
        print(f"❌ Unknown command: {command}")
        print("   Use 'install', 'start', 'stop', 'restart', 'remove', or 'status'")


if __name__ == '__main__':
    if len(sys.argv) == 1:
        # If no arguments, try to run as service
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(OutlookMCPService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Handle command line arguments
        main()