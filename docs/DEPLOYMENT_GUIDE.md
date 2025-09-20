# Deployment Guide

This guide covers deployment options for the Outlook MCP Server in various environments, from development to enterprise production deployments.

## 📋 Table of Contents

- [Deployment Overview](#deployment-overview)
- [Environment Configuration](#environment-configuration)
- [Development Deployment](#development-deployment)
- [Production Deployment](#production-deployment)
- [Windows Service Deployment](#windows-service-deployment)
- [Docker Deployment](#docker-deployment)
- [Monitoring and Health Checks](#monitoring-and-health-checks)
- [Security Considerations](#security-considerations)
- [Troubleshooting](#troubleshooting)

## 🎯 Deployment Overview

The Outlook MCP Server supports multiple deployment modes:

| Mode | Use Case | Features |
|------|----------|----------|
| **Development** | Local development and testing | Console output, debug logging, hot reload |
| **Production** | Production environments | Service mode, health checks, monitoring |
| **Windows Service** | Enterprise deployments | Auto-start, system integration, centralized management |
| **Docker** | Containerized deployments | Isolation, scalability, cloud deployment |

## ⚙️ Environment Configuration

### Environment Variables

The server supports comprehensive environment-based configuration:

```bash
# Core Configuration
OUTLOOK_MCP_CONFIG_FILE=config/production.json
OUTLOOK_MCP_LOG_LEVEL=INFO
OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp-server

# Performance Settings
OUTLOOK_MCP_MAX_CONCURRENT=20
OUTLOOK_MCP_REQUEST_TIMEOUT=45
OUTLOOK_MCP_CONNECTION_TIMEOUT=15

# Feature Toggles
OUTLOOK_MCP_PERFORMANCE_LOGGING=true
OUTLOOK_MCP_CONSOLE_OUTPUT=false

# Server Mode
OUTLOOK_MCP_SERVER_MODE=stdio
OUTLOOK_MCP_HEALTH_CHECK_INTERVAL=60

# Security Settings
OUTLOOK_MCP_RATE_LIMITING=true
OUTLOOK_MCP_MAX_REQUEST_SIZE=1048576
```

### Configuration Files

Create environment-specific configuration files:

#### Development Configuration (`config/development.json`)

```json
{
  "log_level": "DEBUG",
  "log_dir": "logs",
  "max_concurrent_requests": 5,
  "enable_console_output": true,
  "server_mode": "standalone",
  "health_check_interval": 30,
  "rate_limiting": {
    "enabled": false
  },
  "monitoring": {
    "enabled": true,
    "metrics_interval": 60
  }
}
```

#### Production Configuration (`config/production.json`)

```json
{
  "log_level": "INFO",
  "log_dir": "/var/log/outlook-mcp-server",
  "max_concurrent_requests": 20,
  "request_timeout": 45,
  "outlook_connection_timeout": 15,
  "enable_console_output": false,
  "server_mode": "stdio",
  "health_check_interval": 60,
  "rate_limiting": {
    "enabled": true,
    "requests_per_minute": 100,
    "burst_size": 20
  },
  "monitoring": {
    "enabled": true,
    "metrics_interval": 300,
    "health_check_endpoint": true
  },
  "security": {
    "validate_requests": true,
    "sanitize_responses": true,
    "max_request_size": 1048576
  }
}
```

## 🔧 Development Deployment

### Quick Development Setup

```bash
# 1. Clone and setup
git clone https://github.com/your-org/outlook-mcp-server.git
cd outlook-mcp-server
pip install -r requirements.txt

# 2. Create development environment
cp .env.example .env
# Edit .env with your preferences

# 3. Test installation
python main.py test

# 4. Start in development mode
python main.py interactive --log-level DEBUG
```

### Development with Auto-Reload

For development with automatic reloading on code changes:

```bash
# Install development dependencies
pip install watchdog

# Start with file watching
python scripts/dev_server.py --watch
```

### Development Configuration

```bash
# Use development configuration
export OUTLOOK_MCP_CONFIG_FILE=config/development.json
python start_server.py
```

## 🚀 Production Deployment

### Production Startup Script

The production startup script (`start_server.py`) provides enhanced features:

```bash
# Basic production start
python start_server.py --config config/production.json

# Service mode (no console output, PID file)
python start_server.py --service-mode --pid-file /var/run/outlook-mcp.pid

# With environment variables
export OUTLOOK_MCP_LOG_LEVEL=INFO
export OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp-server
python start_server.py

# Test configuration before starting
python start_server.py --validate-config
python start_server.py --test-connection
```

### Production Features

The production deployment includes:

- **Graceful Shutdown**: Proper cleanup on SIGTERM/SIGINT
- **Health Checks**: Periodic health monitoring
- **PID File Management**: Process tracking for service managers
- **Environment Variable Support**: Flexible configuration
- **Error Recovery**: Automatic reconnection on failures
- **Performance Monitoring**: Resource usage tracking

### Systemd Service (Linux-like environments)

Create a systemd service file (`/etc/systemd/system/outlook-mcp-server.service`):

```ini
[Unit]
Description=Outlook MCP Server
After=network.target
Wants=network.target

[Service]
Type=simple
User=outlook-mcp
Group=outlook-mcp
WorkingDirectory=/opt/outlook-mcp-server
Environment=OUTLOOK_MCP_CONFIG_FILE=/etc/outlook-mcp-server/production.json
Environment=OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp-server
ExecStart=/opt/outlook-mcp-server/venv/bin/python start_server.py --service-mode
ExecReload=/bin/kill -HUP $MAINPID
Restart=always
RestartSec=10
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
```

Enable and start the service:

```bash
sudo systemctl daemon-reload
sudo systemctl enable outlook-mcp-server
sudo systemctl start outlook-mcp-server
sudo systemctl status outlook-mcp-server
```

## 🪟 Windows Service Deployment

### Installing as Windows Service

The server includes a Windows service installer:

```bash
# Install service (requires Administrator privileges)
python scripts/install_service.py install

# Start the service
python scripts/install_service.py start

# Check service status
python scripts/install_service.py status

# Stop the service
python scripts/install_service.py stop

# Remove the service
python scripts/install_service.py remove
```

### Service Configuration

The Windows service automatically loads configuration from:

1. `C:\ProgramData\OutlookMCPServer\config.json` (highest priority)
2. `config/production.json` (fallback)
3. `outlook_mcp_server_config.json` (default fallback)

### Service Management

```bash
# Using Windows Services Manager
services.msc

# Using PowerShell
Get-Service "OutlookMCPServer"
Start-Service "OutlookMCPServer"
Stop-Service "OutlookMCPServer"
Restart-Service "OutlookMCPServer"

# Using sc command
sc query "OutlookMCPServer"
sc start "OutlookMCPServer"
sc stop "OutlookMCPServer"
```

### Service Logs

Service logs are written to:
- **Event Log**: Windows Event Viewer → Applications and Services Logs
- **File Logs**: `C:\ProgramData\OutlookMCPServer\logs\`

## 🐳 Docker Deployment

### Dockerfile

```dockerfile
FROM python:3.11-windowsservercore

# Set working directory
WORKDIR /app

# Copy requirements and install dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create logs directory
RUN mkdir logs

# Expose health check port (if using HTTP health checks)
EXPOSE 8080

# Set environment variables
ENV OUTLOOK_MCP_LOG_LEVEL=INFO
ENV OUTLOOK_MCP_LOG_DIR=logs
ENV OUTLOOK_MCP_SERVER_MODE=stdio

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD python -c "import asyncio; from src.outlook_mcp_server.health import is_server_healthy; exit(0 if asyncio.run(is_server_healthy()) else 1)"

# Start the server
CMD ["python", "start_server.py", "--service-mode"]
```

### Docker Compose

```yaml
version: '3.8'

services:
  outlook-mcp-server:
    build: .
    container_name: outlook-mcp-server
    restart: unless-stopped
    environment:
      - OUTLOOK_MCP_LOG_LEVEL=INFO
      - OUTLOOK_MCP_MAX_CONCURRENT=20
      - OUTLOOK_MCP_SERVER_MODE=stdio
    volumes:
      - ./logs:/app/logs
      - ./config:/app/config
    healthcheck:
      test: ["CMD", "python", "-c", "import asyncio; from src.outlook_mcp_server.health import is_server_healthy; exit(0 if asyncio.run(is_server_healthy()) else 1)"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 40s
```

### Building and Running

```bash
# Build the image
docker build -t outlook-mcp-server .

# Run with docker-compose
docker-compose up -d

# Check logs
docker-compose logs -f outlook-mcp-server

# Check health
docker-compose ps
```

## 📊 Monitoring and Health Checks

### Health Check Endpoints

The server provides comprehensive health checking:

```python
# Programmatic health check
import asyncio
from src.outlook_mcp_server.health import get_health_status

async def check_health():
    status = await get_health_status()
    print(f"Status: {status.status}")
    print(f"Outlook Connected: {status.outlook_connected}")
    print(f"Uptime: {status.uptime_seconds}s")

asyncio.run(check_health())
```

### Health Check Script

Create a monitoring script (`scripts/health_check.py`):

```python
#!/usr/bin/env python3
"""Health check script for monitoring systems."""

import asyncio
import sys
import json
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.health import get_health_status

async def main():
    try:
        status = await get_health_status()
        
        # Output JSON for monitoring systems
        result = {
            "status": status.status,
            "healthy": status.status == "healthy",
            "timestamp": status.timestamp,
            "uptime": status.uptime_seconds,
            "outlook_connected": status.outlook_connected
        }
        
        print(json.dumps(result))
        
        # Exit with appropriate code
        sys.exit(0 if status.status == "healthy" else 1)
        
    except Exception as e:
        print(json.dumps({"status": "error", "message": str(e)}))
        sys.exit(2)

if __name__ == "__main__":
    asyncio.run(main())
```

### Monitoring Integration

#### Nagios/Icinga

```bash
# Check command
define command {
    command_name    check_outlook_mcp
    command_line    /usr/local/bin/python3 /opt/outlook-mcp-server/scripts/health_check.py
}

# Service definition
define service {
    use                     generic-service
    host_name               outlook-server
    service_description     Outlook MCP Server
    check_command           check_outlook_mcp
    check_interval          5
    retry_interval          1
}
```

#### Prometheus Metrics

The server can export metrics for Prometheus monitoring:

```python
# Add to your monitoring setup
from src.outlook_mcp_server.monitoring import PrometheusExporter

exporter = PrometheusExporter(port=9090)
exporter.start()
```

## 🔒 Security Considerations

### Production Security Checklist

- [ ] **Run with minimal privileges**: Use dedicated service account
- [ ] **Secure configuration files**: Restrict file permissions (600/640)
- [ ] **Enable request validation**: Validate all incoming requests
- [ ] **Configure rate limiting**: Prevent abuse and DoS attacks
- [ ] **Secure log files**: Protect logs from unauthorized access
- [ ] **Network security**: Use firewalls and network segmentation
- [ ] **Regular updates**: Keep dependencies and system updated
- [ ] **Audit logging**: Enable comprehensive audit trails

### Service Account Setup (Windows)

```powershell
# Create service account
New-LocalUser -Name "OutlookMCPService" -Description "Outlook MCP Server Service Account" -NoPassword
Add-LocalGroupMember -Group "Log on as a service" -Member "OutlookMCPService"

# Set service to run as service account
sc config "OutlookMCPServer" obj= ".\OutlookMCPService" password= ""
```

### File Permissions

```bash
# Secure configuration files
chmod 600 config/*.json
chmod 600 .env

# Secure log directory
chmod 750 logs/
chown outlook-mcp:outlook-mcp logs/

# Secure application directory
chmod 755 /opt/outlook-mcp-server
chown -R outlook-mcp:outlook-mcp /opt/outlook-mcp-server
```

## 🔧 Troubleshooting

### Common Deployment Issues

#### Service Won't Start

```bash
# Check service status
python scripts/install_service.py status

# Check Windows Event Log
eventvwr.msc
# Navigate to: Applications and Services Logs

# Check configuration
python start_server.py --validate-config

# Test Outlook connection
python start_server.py --test-connection
```

#### Permission Issues

```bash
# Run as Administrator (Windows)
# Check Outlook security settings
# Verify service account permissions
```

#### Configuration Problems

```bash
# Validate configuration syntax
python -c "import json; json.load(open('config/production.json'))"

# Test with minimal configuration
python start_server.py --config config/minimal.json --test-connection
```

#### Performance Issues

```bash
# Monitor resource usage
python scripts/monitor_performance.py

# Check logs for bottlenecks
findstr "SLOW" logs\outlook_mcp_server.log

# Adjust configuration
# - Reduce max_concurrent_requests
# - Increase timeouts
# - Enable caching
```

### Log Analysis

```bash
# Check for errors
findstr "ERROR" logs\outlook_mcp_server.log

# Monitor performance
findstr "PERFORMANCE" logs\outlook_mcp_server.log

# Check health status
findstr "HEALTH" logs\outlook_mcp_server.log
```

### Getting Help

For deployment issues:

1. **Check the logs**: Always start with the log files
2. **Validate configuration**: Use `--validate-config` flag
3. **Test connection**: Use `--test-connection` flag
4. **Check documentation**: Review [TROUBLESHOOTING.md](TROUBLESHOOTING.md)
5. **Report issues**: Create GitHub issue with logs and configuration

---

## 📚 Additional Resources

- [Setup Guide](SETUP_GUIDE.md) - Initial installation and setup
- [API Documentation](API_DOCUMENTATION.md) - Complete API reference
- [Troubleshooting Guide](TROUBLESHOOTING.md) - Problem resolution
- [Examples](EXAMPLES.md) - Usage examples and integration patterns

---

**Need help with deployment?** Contact our support team or check the [GitHub Discussions](https://github.com/your-org/outlook-mcp-server/discussions) for community support.