# Outlook MCP Server

A comprehensive Model Context Protocol (MCP) server that provides programmatic access to Microsoft Outlook email functionality. This server implements the complete MCP protocol specification and exposes four core email operations with advanced features including performance optimization, comprehensive error handling, and detailed logging.

## 🚀 Features

- **📧 Complete Email Operations**: List, retrieve, search emails and manage folders
- **🔌 MCP Protocol Compliance**: Full implementation of MCP protocol for seamless integration
- **⚡ Performance Optimized**: Connection pooling, caching, lazy loading, and rate limiting
- **🛡️ Robust Error Handling**: Comprehensive error categorization with detailed diagnostics
- **📊 Advanced Logging**: Structured JSON logging with performance metrics and audit trails
- **🔄 Concurrent Processing**: Efficient handling of multiple simultaneous requests
- **🔧 Extensive Configuration**: Flexible configuration options for all environments
- **📚 Comprehensive Documentation**: Complete API docs, examples, and troubleshooting guides

## 📋 Requirements

### System Requirements
- **Operating System**: Windows 10+ (required for COM interface)
- **Microsoft Outlook**: 2016 or later, installed and configured
- **Python**: 3.8 or later
- **Memory**: 4GB RAM recommended
- **Storage**: 2GB free disk space

### Dependencies
All dependencies are listed in `requirements.txt`:
- `pywin32>=306` - Windows COM interface
- `mcp>=1.0.0` - Model Context Protocol implementation
- `asyncio` - Asynchronous programming support
- `pytest>=7.0.0` - Testing framework
- `pytest-asyncio>=0.21.0` - Async testing support

## 🛠️ Installation

### Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/your-org/outlook-mcp-server.git
cd outlook-mcp-server

# 2. Install dependencies
pip install -r requirements.txt

# 3. Test the installation
python main.py test

# 4. Create configuration (optional)
python main.py create-config

# 5. Start the server
python main.py stdio
```

### Production Deployment

For production environments, use the enhanced startup script:

```bash
# Start with production configuration
python start_server.py --config config/production.json

# Start as a service (Windows)
python start_server.py --service-mode

# Start with environment variables
export OUTLOOK_MCP_LOG_LEVEL=INFO
export OUTLOOK_MCP_LOG_DIR=/var/log/outlook-mcp
python start_server.py

# Install as Windows Service
python scripts/install_service.py install
python scripts/install_service.py start
```

### Environment Configuration

Copy and customize the environment template:

```bash
# Copy environment template
cp .env.example .env

# Edit configuration
notepad .env  # Windows
```

### Detailed Installation

For detailed installation instructions including system setup, security configuration, and deployment options, see [docs/SETUP_GUIDE.md](docs/SETUP_GUIDE.md).

## 🎯 Usage

### Basic Usage

```bash
# Run as MCP server (default mode)
python main.py stdio

# Run with custom configuration
python main.py stdio --config my_config.json

# Interactive mode for development
python main.py interactive --log-level DEBUG

# Test Outlook connection
python main.py test
```

### Production Usage

```bash
# Production startup with health checks
python start_server.py --config config/production.json

# Service mode (no console output)
python start_server.py --service-mode --pid-file /var/run/outlook-mcp.pid

# Test connection before starting
python start_server.py --test-connection

# Validate configuration
python start_server.py --validate-config
```

### Windows Service Management

```bash
# Install service
python scripts/install_service.py install

# Start/stop service
python scripts/install_service.py start
python scripts/install_service.py stop

# Check service status
python scripts/install_service.py status

# Remove service
python scripts/install_service.py remove
```

### Health Monitoring

```bash
# Check server health
python -c "
import asyncio
from src.outlook_mcp_server.health import get_health_status
status = asyncio.run(get_health_status())
print(f'Status: {status.status}')
print(f'Outlook Connected: {status.outlook_connected}')
"
```

### Configuration

Create and customize your configuration:

```bash
# Generate default configuration file
python main.py create-config
```

Example configuration (`outlook_mcp_server_config.json`):

```json
{
  "log_level": "INFO",
  "log_dir": "logs",
  "max_concurrent_requests": 10,
  "request_timeout": 30,
  "outlook_connection_timeout": 10,
  "enable_performance_logging": true,
  "rate_limiting": {
    "enabled": true,
    "requests_per_minute": 100
  },
  "caching": {
    "enabled": true,
    "email_cache_ttl": 300,
    "max_cache_size_mb": 100
  },
  "security": {
    "allowed_folders": ["Inbox", "Sent Items"],
    "max_email_size_mb": 50
  }
}
```

## 📖 Documentation

### Complete Documentation Suite

- **[API Documentation](docs/API_DOCUMENTATION.md)** - Complete API reference with examples
- **[Setup Guide](docs/SETUP_GUIDE.md)** - Detailed installation and configuration
- **[Usage Examples](docs/EXAMPLES.md)** - Practical examples and integration patterns
- **[Troubleshooting Guide](docs/TROUBLESHOOTING.md)** - Common issues and solutions

### Quick API Reference

#### Available MCP Methods

| Method | Description | Parameters |
|--------|-------------|------------|
| `list_emails` | List emails with filtering | `folder`, `unread_only`, `limit` |
| `get_email` | Get detailed email by ID | `email_id` |
| `search_emails` | Search emails by query | `query`, `folder`, `limit` |
| `get_folders` | List all available folders | None |

#### Example Request

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "method": "list_emails",
  "params": {
    "folder": "Inbox",
    "unread_only": true,
    "limit": 10
  }
}
```

#### Example Response

```json
{
  "jsonrpc": "2.0",
  "id": "1",
  "result": [
    {
      "id": "AAMkADEx...",
      "subject": "Project Update",
      "sender": "john.doe@company.com",
      "received_time": "2024-01-15T10:30:00Z",
      "is_read": false,
      "has_attachments": true,
      "folder_name": "Inbox"
    }
  ]
}
```

## 🏗️ Architecture

### System Architecture

```
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   MCP Client    │◄──►│  MCP Protocol   │◄──►│ Request Router  │
│                 │    │    Handler      │    │                 │
└─────────────────┘    └─────────────────┘    └─────────────────┘
                                                        │
                       ┌─────────────────┐    ┌─────────────────┐
                       │ Email Service   │◄──►│ Folder Service  │
                       │                 │    │                 │
                       └─────────────────┘    └─────────────────┘
                                │                       │
                       ┌─────────────────────────────────────────┐
                       │         Outlook Adapter                 │
                       │    (COM Interface Management)           │
                       └─────────────────────────────────────────┘
                                        │
                       ┌─────────────────────────────────────────┐
                       │        Microsoft Outlook               │
                       │         (COM Objects)                   │
                       └─────────────────────────────────────────┘
```

### Key Components

- **MCP Protocol Handler**: Manages protocol compliance and message formatting
- **Request Router**: Routes and validates incoming requests
- **Service Layer**: Business logic for email and folder operations
- **Outlook Adapter**: Low-level COM interface with Outlook
- **Performance Layer**: Caching, connection pooling, and optimization
- **Error Handler**: Comprehensive error processing and recovery
- **Logging System**: Structured logging with performance metrics

## 🔧 Development

### Project Structure

```
outlook-mcp-server/
├── docs/                           # Complete documentation
│   ├── API_DOCUMENTATION.md        # API reference
│   ├── EXAMPLES.md                 # Usage examples
│   ├── SETUP_GUIDE.md             # Installation guide
│   └── TROUBLESHOOTING.md         # Problem resolution
├── src/outlook_mcp_server/         # Main source code
│   ├── adapters/                   # Outlook COM integration
│   ├── services/                   # Business logic layer
│   ├── protocol/                   # MCP protocol handling
│   ├── routing/                    # Request routing
│   ├── models/                     # Data models and exceptions
│   ├── logging/                    # Logging system
│   ├── performance/                # Performance optimizations
│   ├── server.py                   # Main server class
│   └── main.py                     # Entry point
├── tests/                          # Comprehensive test suite
├── examples/                       # Example scripts
├── main.py                         # Startup script
├── requirements.txt                # Dependencies
└── README.md                       # This file
```

### Running Tests

```bash
# Run all tests
python -m pytest tests/

# Run with coverage
python -m pytest tests/ --cov=src/outlook_mcp_server

# Run integration tests
python -m pytest tests/test_integration.py

# Run performance tests
python tests/run_integration_tests.py
```

### Development Setup

```bash
# Install development dependencies
pip install -r requirements.txt
pip install pytest pytest-cov pytest-asyncio

# Run in development mode
python main.py interactive --log-level DEBUG

# Enable performance profiling
python examples/profile_server.py
```

## 🚀 Performance Features

### Optimization Technologies

- **Connection Pooling**: Reuse Outlook COM connections
- **Intelligent Caching**: Multi-level caching with TTL
- **Lazy Loading**: Load email content on demand
- **Rate Limiting**: Prevent system overload
- **Memory Management**: Automatic cleanup and optimization
- **Concurrent Processing**: Handle multiple requests efficiently

### Performance Monitoring

```bash
# Monitor performance in real-time
python main.py interactive --log-level INFO

# Generate performance reports
python scripts/analyze_performance.py logs/performance.log

# Memory usage monitoring
python examples/memory_monitor.py
```

## 🛡️ Security & Error Handling

### Security Features

- **Folder Access Control**: Restrict access to specific folders
- **Content Filtering**: Sanitize email content and attachments
- **Rate Limiting**: Prevent abuse and DoS attacks
- **Input Validation**: Comprehensive parameter validation
- **Audit Logging**: Complete audit trail of all operations

### Error Handling

The server provides comprehensive error handling with detailed diagnostics:

- **Validation Errors**: Invalid parameters or malformed requests
- **Connection Errors**: Outlook connectivity issues
- **Permission Errors**: Access control violations
- **Resource Errors**: Missing emails or folders
- **Performance Errors**: Timeouts and rate limits

Each error includes:
- Appropriate MCP error codes
- Detailed error messages
- Contextual information for debugging
- Suggested resolution steps

## 📊 Monitoring & Logging

### Logging Features

- **Structured JSON Logging**: Machine-readable log format
- **Performance Metrics**: Request timing and resource usage
- **Audit Trails**: Complete operation history
- **Error Tracking**: Detailed error information with context
- **Log Rotation**: Automatic log file management

### Log Analysis

```bash
# View recent errors
findstr "ERROR" logs\outlook_mcp_server.log

# Analyze performance
python scripts/analyze_logs.py --performance

# Generate reports
python scripts/generate_report.py --daily
```

## 🆘 Troubleshooting

### Quick Diagnostics

```bash
# Test system health
python main.py test

# Check configuration
python main.py stdio --config my_config.json --test-connection

# Debug mode
python main.py interactive --log-level DEBUG
```

### Common Issues

| Issue | Solution |
|-------|----------|
| Outlook not found | Ensure Outlook is installed and registered |
| Permission denied | Run as administrator or check Outlook security settings |
| Connection timeout | Increase `outlook_connection_timeout` in configuration |
| High memory usage | Reduce cache size or enable memory management |

For detailed troubleshooting, see [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md).

## 🤝 Contributing

We welcome contributions! Please see our contributing guidelines:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Guidelines

- Follow PEP 8 style guidelines
- Add comprehensive tests for new features
- Update documentation for API changes
- Ensure all tests pass before submitting
- Include type hints for all functions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Support

### Getting Help

- **Documentation**: Check the [docs/](docs/) directory
- **Issues**: Report bugs on [GitHub Issues](https://github.com/your-org/outlook-mcp-server/issues)
- **Discussions**: Join [GitHub Discussions](https://github.com/your-org/outlook-mcp-server/discussions)
- **Email**: Contact support@your-org.com

### Enterprise Support

For enterprise deployments and commercial support:
- Professional installation and configuration
- Custom integration development
- Priority support and SLA
- Training and consultation

Contact: enterprise@your-org.com

---

**Made with ❤️ by the Outlook MCP Server Team**

*Empowering developers to build amazing email-integrated applications with the power of Microsoft Outlook and the Model Context Protocol.*