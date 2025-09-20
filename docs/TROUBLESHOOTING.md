# Outlook MCP Server - Troubleshooting Guide

This guide helps you diagnose and resolve common issues with the Outlook MCP Server.

## Table of Contents

- [Quick Diagnostics](#quick-diagnostics)
- [Common Issues](#common-issues)
- [Connection Problems](#connection-problems)
- [Performance Issues](#performance-issues)
- [Error Messages](#error-messages)
- [Debugging Tools](#debugging-tools)
- [Advanced Troubleshooting](#advanced-troubleshooting)
- [Getting Help](#getting-help)

## Quick Diagnostics

### Step 1: Basic Health Check

Run the built-in diagnostic tool:

```bash
python main.py test
```

**Expected Output:**
```
Testing Outlook connection...
✓ Outlook connection successful
✓ Server is healthy
✓ Found X accessible folders
✓ Email access working
Connection test completed successfully
```

**If this fails**, see [Connection Problems](#connection-problems).

### Step 2: Check System Requirements

Verify your system meets the requirements:

```powershell
# Check Windows version
Get-ComputerInfo | Select-Object WindowsProductName, WindowsVersion

# Check Python version
python --version

# Check if Outlook is installed
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Outlook*"}

# Test Outlook COM access
New-Object -ComObject Outlook.Application
```

### Step 3: Check Logs

Look at the most recent log entries:

```bash
# View recent logs (Windows)
type logs\outlook_mcp_server.log | findstr /C:"ERROR" /C:"WARNING"

# View recent logs (PowerShell)
Get-Content logs\outlook_mcp_server.log -Tail 50 | Where-Object {$_ -match "ERROR|WARNING"}
```

## Common Issues

### Issue 1: "Outlook application not found" or "COM object creation failed"

**Symptoms:**
- Server fails to start
- Error message about Outlook not being available
- COM object creation errors

**Causes:**
- Outlook not installed
- Outlook not properly registered
- User doesn't have permission to access Outlook
- Outlook is in safe mode

**Solutions:**

1. **Verify Outlook Installation:**
   ```bash
   # Check if Outlook is installed
   dir "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
   # or
   dir "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"
   ```

2. **Re-register Outlook COM Objects:**
   ```cmd
   # Run as Administrator
   cd "C:\Program Files\Microsoft Office\root\Office16"
   outlook.exe /regserver
   ```

3. **Check Outlook Profile:**
   ```cmd
   # Open Outlook profile configuration
   outlook.exe /safe
   ```

4. **Verify User Permissions:**
   - Ensure user can open Outlook normally
   - Check if Outlook requires administrator privileges
   - Verify no group policies block COM access

### Issue 2: "Permission denied" or "Access is denied"

**Symptoms:**
- Server starts but can't access emails
- Specific folders are inaccessible
- Intermittent permission errors

**Causes:**
- Insufficient user permissions
- Outlook security settings
- Antivirus blocking COM access
- Corporate security policies

**Solutions:**

1. **Run as Administrator:**
   ```cmd
   # Run command prompt as administrator
   runas /user:Administrator cmd
   python main.py test
   ```

2. **Check Outlook Security Settings:**
   - Open Outlook → File → Options → Trust Center → Trust Center Settings
   - Check "Programmatic Access" settings
   - Disable "Warn me about suspicious activity" temporarily

3. **Configure Antivirus Exclusions:**
   - Add Python.exe to antivirus exclusions
   - Add the server directory to exclusions
   - Temporarily disable real-time protection for testing

4. **Check Group Policies:**
   ```cmd
   # Check relevant group policies
   gpresult /r | findstr /i outlook
   ```

### Issue 3: Server starts but no emails are returned

**Symptoms:**
- Server appears healthy
- Folder list is empty or incomplete
- Email queries return no results

**Causes:**
- Outlook profile not configured
- No email accounts configured
- Folders are empty
- Synchronization issues

**Solutions:**

1. **Verify Outlook Configuration:**
   ```bash
   # Test with interactive mode
   python main.py interactive --log-level DEBUG
   ```

2. **Check Email Account Status:**
   - Open Outlook manually
   - Verify email accounts are connected
   - Check for synchronization errors
   - Ensure folders contain emails

3. **Test Specific Folders:**
   ```python
   # Test script to check folder access
   import asyncio
   from src.outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter
   
   async def test_folders():
       adapter = OutlookAdapter()
       await adapter.connect()
       
       folders = adapter.get_folders()
       print(f"Found {len(folders)} folders:")
       
       for folder in folders:
           print(f"- {folder.name}: {folder.item_count} items")
           
       await adapter.disconnect()
   
   asyncio.run(test_folders())
   ```

### Issue 4: High memory usage or performance issues

**Symptoms:**
- Server uses excessive memory
- Slow response times
- Timeouts on requests
- System becomes unresponsive

**Causes:**
- Large email attachments
- Inefficient caching
- Too many concurrent requests
- Memory leaks

**Solutions:**

1. **Optimize Configuration:**
   ```json
   {
     "max_concurrent_requests": 5,
     "request_timeout": 60,
     "caching": {
       "max_cache_size_mb": 50,
       "cache_cleanup_interval": 1800
     },
     "security": {
       "max_email_size_mb": 25,
       "strip_attachments": true
     }
   }
   ```

2. **Monitor Memory Usage:**
   ```python
   # Memory monitoring script
   import psutil
   import time
   
   def monitor_memory():
       process = psutil.Process()
       while True:
           memory_mb = process.memory_info().rss / 1024 / 1024
           print(f"Memory usage: {memory_mb:.1f} MB")
           time.sleep(30)
   
   monitor_memory()
   ```

3. **Enable Performance Logging:**
   ```json
   {
     "enable_performance_logging": true,
     "log_level": "INFO"
   }
   ```

## Connection Problems

### Outlook COM Connection Issues

**Diagnostic Steps:**

1. **Test COM Access Directly:**
   ```python
   import win32com.client
   
   try:
       outlook = win32com.client.Dispatch("Outlook.Application")
       namespace = outlook.GetNamespace("MAPI")
       print("✓ COM connection successful")
       
       folders = namespace.Folders
       print(f"✓ Found {folders.Count} top-level folders")
       
   except Exception as e:
       print(f"✗ COM connection failed: {e}")
   ```

2. **Check Outlook Process:**
   ```cmd
   # Check if Outlook is running
   tasklist | findstr outlook.exe
   
   # Kill Outlook if needed
   taskkill /f /im outlook.exe
   ```

3. **Verify MAPI Installation:**
   ```cmd
   # Check MAPI installation
   reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Messaging Subsystem" /v MAPI
   ```

### Network and Firewall Issues

**For Remote Deployments:**

1. **Check Network Connectivity:**
   ```cmd
   # Test network connectivity
   ping outlook.office365.com
   telnet outlook.office365.com 993
   ```

2. **Verify Firewall Rules:**
   ```cmd
   # Check Windows Firewall
   netsh advfirewall firewall show rule name=all | findstr python
   
   # Add firewall rule if needed
   netsh advfirewall firewall add rule name="Python MCP Server" dir=in action=allow program="C:\Python\python.exe"
   ```

## Performance Issues

### Slow Response Times

**Diagnostic Steps:**

1. **Enable Performance Logging:**
   ```json
   {
     "log_level": "DEBUG",
     "enable_performance_logging": true
   }
   ```

2. **Analyze Performance Logs:**
   ```bash
   # Find slow operations
   findstr "processing_time_ms" logs\performance.log | findstr /R "[5-9][0-9][0-9][0-9]"
   ```

3. **Profile Memory Usage:**
   ```python
   # Memory profiling script
   import tracemalloc
   import asyncio
   from src.outlook_mcp_server.server import OutlookMCPServer
   
   async def profile_memory():
       tracemalloc.start()
       
       server = OutlookMCPServer({})
       await server.start()
       
       # Simulate some operations
       for i in range(10):
           await server.email_service.list_emails(limit=50)
       
       current, peak = tracemalloc.get_traced_memory()
       print(f"Current memory usage: {current / 1024 / 1024:.1f} MB")
       print(f"Peak memory usage: {peak / 1024 / 1024:.1f} MB")
       
       await server.stop()
       tracemalloc.stop()
   
   asyncio.run(profile_memory())
   ```

### High CPU Usage

**Solutions:**

1. **Reduce Concurrent Requests:**
   ```json
   {
     "max_concurrent_requests": 3,
     "request_timeout": 30
   }
   ```

2. **Optimize Caching:**
   ```json
   {
     "caching": {
       "enabled": true,
       "email_cache_ttl": 600,
       "folder_cache_ttl": 1800
     }
   }
   ```

3. **Use Rate Limiting:**
   ```json
   {
     "rate_limiting": {
       "enabled": true,
       "requests_per_minute": 60
     }
   }
   ```

## Error Messages

### "OutlookConnectionError: Not connected to Outlook"

**Cause:** Server lost connection to Outlook COM interface.

**Solutions:**
1. Restart Outlook application
2. Restart the MCP server
3. Check if Outlook crashed or was closed
4. Verify COM registration

**Prevention:**
```json
{
  "outlook_connection_timeout": 15,
  "connection_retry_attempts": 3,
  "connection_retry_delay": 5
}
```

### "ValidationError: Invalid email ID format"

**Cause:** Malformed email ID passed to get_email method.

**Solutions:**
1. Verify email ID format (should be Outlook EntryID)
2. Check if email was deleted or moved
3. Use list_emails to get valid email IDs

**Example Valid Email ID:**
```
AAMkADExMzJmYWE3LTM4ZGYtNDk2Yy1hMjU4LWVmYzJkNzNkNzE2MwBGAAAAAAC7XK...
```

### "FolderNotFoundError: Folder 'X' not found"

**Cause:** Specified folder doesn't exist or isn't accessible.

**Solutions:**
1. Use get_folders to list available folders
2. Check folder name spelling and case
3. Verify folder permissions
4. Check if folder was deleted or renamed

### "RateLimitError: Too many requests"

**Cause:** Client exceeded rate limits.

**Solutions:**
1. Implement client-side rate limiting
2. Add delays between requests
3. Increase server rate limits if appropriate

**Client-side Solution:**
```python
import asyncio

async def rate_limited_requests():
    for i in range(10):
        # Make request
        result = await client.send_request(request)
        
        # Wait between requests
        await asyncio.sleep(1)
```

### "SearchError: Search operation failed"

**Cause:** Invalid search query or search service issues.

**Solutions:**
1. Simplify search query
2. Check search syntax
3. Verify Outlook search service is running
4. Rebuild Outlook search index

**Rebuild Search Index:**
```cmd
# Rebuild Outlook search index
outlook.exe /cleanfinders
```

## Debugging Tools

### Enable Debug Mode

```bash
# Run with maximum debugging
python main.py interactive --log-level DEBUG
```

### Custom Debug Script

```python
# debug_server.py
import asyncio
import logging
from src.outlook_mcp_server.server import OutlookMCPServer

# Configure detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

async def debug_session():
    """Interactive debugging session."""
    
    config = {
        "log_level": "DEBUG",
        "enable_performance_logging": True,
        "max_concurrent_requests": 1  # Single-threaded for debugging
    }
    
    server = OutlookMCPServer(config)
    
    try:
        print("🔧 Starting debug session...")
        await server.start()
        
        # Test basic operations
        print("\n📁 Testing folder access...")
        folders = server.folder_service.get_folders()
        print(f"Found {len(folders)} folders")
        
        print("\n📧 Testing email access...")
        emails = await server.email_service.list_emails(limit=5)
        print(f"Found {len(emails)} emails")
        
        if emails:
            print(f"\n📖 Testing email details...")
            email_details = await server.email_service.get_email(emails[0]["id"])
            print(f"Retrieved email: {email_details['subject']}")
        
        print("\n🔍 Testing search...")
        search_results = await server.email_service.search_emails("test", limit=3)
        print(f"Search found {len(search_results)} results")
        
        print("\n✅ All tests completed successfully")
        
    except Exception as e:
        print(f"\n❌ Debug session failed: {str(e)}")
        import traceback
        traceback.print_exc()
        
    finally:
        await server.stop()
        print("\n👋 Debug session ended")

if __name__ == "__main__":
    asyncio.run(debug_session())
```

### Performance Profiler

```python
# profile_server.py
import cProfile
import pstats
import asyncio
from src.outlook_mcp_server.server import OutlookMCPServer

async def profile_operations():
    """Profile server operations."""
    
    server = OutlookMCPServer({})
    await server.start()
    
    # Profile email listing
    pr = cProfile.Profile()
    pr.enable()
    
    emails = await server.email_service.list_emails(limit=100)
    
    pr.disable()
    
    # Save profile results
    pr.dump_stats('profile_results.prof')
    
    # Print top functions
    stats = pstats.Stats('profile_results.prof')
    stats.sort_stats('cumulative')
    stats.print_stats(20)
    
    await server.stop()

asyncio.run(profile_operations())
```

### Memory Leak Detector

```python
# memory_leak_detector.py
import gc
import psutil
import asyncio
import time
from src.outlook_mcp_server.server import OutlookMCPServer

async def detect_memory_leaks():
    """Detect potential memory leaks."""
    
    server = OutlookMCPServer({})
    await server.start()
    
    process = psutil.Process()
    initial_memory = process.memory_info().rss
    
    print(f"Initial memory: {initial_memory / 1024 / 1024:.1f} MB")
    
    # Perform operations in a loop
    for i in range(100):
        emails = await server.email_service.list_emails(limit=10)
        
        if i % 10 == 0:
            # Force garbage collection
            gc.collect()
            
            current_memory = process.memory_info().rss
            memory_growth = (current_memory - initial_memory) / 1024 / 1024
            
            print(f"Iteration {i}: Memory = {current_memory / 1024 / 1024:.1f} MB "
                  f"(+{memory_growth:.1f} MB)")
            
            if memory_growth > 100:  # Alert if growth > 100MB
                print("⚠️ Potential memory leak detected!")
                break
    
    await server.stop()

asyncio.run(detect_memory_leaks())
```

## Advanced Troubleshooting

### COM Object Debugging

```python
# com_debug.py
import win32com.client
import pythoncom

def debug_com_objects():
    """Debug COM object interactions."""
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Create Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        print(f"✓ Outlook version: {outlook.Version}")
        
        # Get namespace
        namespace = outlook.GetNamespace("MAPI")
        print(f"✓ MAPI namespace: {namespace}")
        
        # Test folder access
        folders = namespace.Folders
        print(f"✓ Root folders: {folders.Count}")
        
        for i in range(folders.Count):
            folder = folders.Item(i + 1)
            print(f"  - {folder.Name}: {folder.Items.Count} items")
        
        # Test default folder access
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
        print(f"✓ Inbox: {inbox.Items.Count} items")
        
    except Exception as e:
        print(f"✗ COM debugging failed: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    debug_com_objects()
```

### Registry Diagnostics

```cmd
@echo off
echo Checking Outlook COM registration...

echo.
echo Outlook Application Class:
reg query "HKEY_CLASSES_ROOT\Outlook.Application" 2>nul
if errorlevel 1 echo ✗ Outlook.Application not registered

echo.
echo MAPI Configuration:
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows Messaging Subsystem" 2>nul
if errorlevel 1 echo ✗ MAPI not properly configured

echo.
echo Office Installation:
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office" 2>nul
if errorlevel 1 echo ✗ Office not found in registry

echo.
echo Outlook Profiles:
reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Profiles" 2>nul
if errorlevel 1 echo ✗ No Outlook profiles found

pause
```

### Event Log Analysis

```powershell
# Check Windows Event Logs for Outlook/COM errors
Get-WinEvent -FilterHashtable @{LogName='Application'; ProviderName='Outlook'} -MaxEvents 50 | 
    Where-Object {$_.LevelDisplayName -eq 'Error'} |
    Format-Table TimeCreated, Id, LevelDisplayName, Message -Wrap

# Check for COM errors
Get-WinEvent -FilterHashtable @{LogName='System'; Id=10016} -MaxEvents 20 |
    Format-Table TimeCreated, Message -Wrap
```

## Getting Help

### Before Seeking Help

1. **Run Diagnostics:**
   ```bash
   python main.py test > diagnostic_output.txt 2>&1
   ```

2. **Collect Logs:**
   ```bash
   # Compress logs for sharing
   powershell Compress-Archive -Path logs\* -DestinationPath logs_backup.zip
   ```

3. **Document Your Environment:**
   ```bash
   # System information
   systeminfo > system_info.txt
   python --version > python_info.txt
   pip list > installed_packages.txt
   ```

### Information to Include

When seeking help, provide:

1. **Error Messages**: Complete error messages and stack traces
2. **System Information**: OS version, Python version, Outlook version
3. **Configuration**: Your configuration file (remove sensitive data)
4. **Log Files**: Recent log entries showing the issue
5. **Steps to Reproduce**: Exact steps that cause the problem
6. **Expected vs Actual**: What you expected vs what happened

### Support Channels

1. **GitHub Issues**: For bugs and feature requests
2. **Documentation**: Check the docs/ directory
3. **Community Forums**: Stack Overflow with tags `outlook-mcp-server`
4. **Enterprise Support**: Contact your system administrator

### Self-Help Resources

1. **Test Mode**: Always start with `python main.py test`
2. **Debug Logs**: Enable DEBUG logging for detailed information
3. **Minimal Configuration**: Test with minimal configuration first
4. **Isolation Testing**: Test components individually
5. **Documentation**: Read all documentation thoroughly

### Creating Minimal Reproduction

```python
# minimal_repro.py - Create minimal reproduction case
import asyncio
from src.outlook_mcp_server.server import OutlookMCPServer

async def minimal_test():
    """Minimal test case for reproduction."""
    
    # Use minimal configuration
    config = {
        "log_level": "DEBUG",
        "max_concurrent_requests": 1
    }
    
    server = OutlookMCPServer(config)
    
    try:
        await server.start()
        
        # Reproduce the specific issue here
        # Example: List emails that causes the problem
        emails = await server.email_service.list_emails(limit=1)
        print(f"Success: Found {len(emails)} emails")
        
    except Exception as e:
        print(f"Error reproduced: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        await server.stop()

if __name__ == "__main__":
    asyncio.run(minimal_test())
```

This minimal reproduction script helps isolate the issue and provides a clear test case for debugging.