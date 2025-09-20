# EmailMCP

A Model Context Protocol (MCP) server to read Outlook emails using Python.

EmailMCP provides a standardized interface for AI assistants and other MCP clients to interact with email data from Outlook/Office 365 and Exchange servers.

## Features

- **MCP Protocol Support**: Fully compatible with the Model Context Protocol specification
- **Multiple Email Providers**: Support for Outlook/Office 365 and Exchange Server
- **Comprehensive Email Operations**: List emails, search, get detailed email content, and manage folders
- **Configurable**: Flexible configuration via environment variables or JSON files
- **Async/Await**: Built with Python's asyncio for efficient handling of concurrent operations
- **Mock Mode**: Test and develop without real email credentials

## Quick Start

### Installation

1. Clone this repository:
```bash
git clone https://github.com/bitsplus1/EmailMCP.git
cd EmailMCP
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. For development (optional):
```bash
pip install -r requirements-dev.txt
```

### Configuration

1. Copy the example environment file:
```bash
cp .env.example .env
```

2. Edit `.env` with your email credentials:

**For Outlook/Office 365:**
```env
OUTLOOK_CLIENT_ID=your_outlook_client_id_here
OUTLOOK_CLIENT_SECRET=your_outlook_client_secret_here
OUTLOOK_TENANT_ID=your_outlook_tenant_id_here
```

**For Exchange Server:**
```env
EXCHANGE_SERVER=your_exchange_server_here
EXCHANGE_USERNAME=your_exchange_username_here
EXCHANGE_PASSWORD=your_exchange_password_here
EXCHANGE_DOMAIN=your_exchange_domain_here
```

### Running the Server

Start the MCP server:
```bash
python -m email_mcp.server
```

Or use the CLI:
```bash
python -m email_mcp.cli
```

Test your connection:
```bash
python -m email_mcp.cli --test
```

## Available Tools

The EmailMCP server provides the following tools via the MCP protocol:

### 1. `list_emails`
List emails from a specified folder with optional filtering.

**Parameters:**
- `folder` (string, optional): Email folder to search (default: "inbox")
- `limit` (integer, optional): Maximum number of emails to return (default: 10, max: 100)
- `unread_only` (boolean, optional): Only return unread emails (default: false)

### 2. `get_email`
Get detailed information about a specific email.

**Parameters:**
- `email_id` (string, required): Unique identifier of the email

### 3. `search_emails`
Search emails by subject, sender, or content.

**Parameters:**
- `query` (string, required): Search query
- `folder` (string, optional): Email folder to search (default: "inbox")
- `limit` (integer, optional): Maximum number of results (default: 10, max: 50)

### 4. `get_folders`
List all available email folders.

**Parameters:** None

## Configuration Options

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `OUTLOOK_CLIENT_ID` | Azure App Registration Client ID | - |
| `OUTLOOK_CLIENT_SECRET` | Azure App Registration Client Secret | - |
| `OUTLOOK_TENANT_ID` | Azure Tenant ID | - |
| `OUTLOOK_REDIRECT_URI` | OAuth redirect URI | `http://localhost:8080/callback` |
| `EXCHANGE_SERVER` | Exchange server URL | - |
| `EXCHANGE_USERNAME` | Exchange username | - |
| `EXCHANGE_PASSWORD` | Exchange password | - |
| `EXCHANGE_DOMAIN` | Exchange domain | - |
| `EXCHANGE_AUTODISCOVER` | Use Exchange autodiscover | `true` |
| `LOG_LEVEL` | Logging level | `INFO` |
| `MAX_EMAILS_PER_REQUEST` | Maximum emails per request | `100` |
| `CACHE_TTL` | Cache time-to-live (seconds) | `300` |
| `TIMEOUT` | Request timeout (seconds) | `30` |

### JSON Configuration

You can also use a JSON configuration file:

```bash
python -m email_mcp.cli --create-config config.json
python -m email_mcp.cli --config config.json
```

## Setting up Outlook/Office 365

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Set the redirect URI to `http://localhost:8080/callback`
5. Under "API permissions", add:
   - `Mail.Read`
   - `Mail.ReadWrite` 
   - `User.Read`
6. Generate a client secret
7. Note down the Application (client) ID, Directory (tenant) ID, and client secret

## Development

### Project Structure

```
EmailMCP/
├── email_mcp/
│   ├── __init__.py          # Package initialization
│   ├── server.py            # Main MCP server implementation
│   ├── clients.py           # Email client implementations
│   ├── config.py            # Configuration management
│   └── cli.py               # Command-line interface
├── requirements.txt         # Production dependencies
├── requirements-dev.txt     # Development dependencies
├── pyproject.toml          # Project configuration
├── .env.example            # Example environment file
├── .gitignore              # Git ignore rules
└── README.md               # This file
```

### Adding New Email Providers

To add support for a new email provider:

1. Create a new client class in `clients.py` that inherits from `EmailClient`
2. Implement all abstract methods
3. Add configuration options in `config.py`
4. Update the client setup logic in `server.py`

### Running Tests

```bash
pytest
```

### Code Formatting

```bash
black email_mcp/
isort email_mcp/
```

### Type Checking

```bash
mypy email_mcp/
```

## Usage with MCP Clients

This server is designed to work with any MCP-compatible client. Here's how to connect:

1. Start the EmailMCP server
2. Configure your MCP client to connect to the server
3. Use the available tools to interact with your email data

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Verify your credentials in the `.env` file
2. **Connection Timeouts**: Check your network connection and server settings
3. **Permission Errors**: Ensure your Azure app has the required API permissions

### Logging

Enable debug logging to see detailed information:

```bash
python -m email_mcp.cli --log-level DEBUG
```

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## Support

For issues and questions:
- Create an issue on GitHub
- Check the troubleshooting section in this README

## Roadmap

- [ ] Complete Outlook/Office 365 integration
- [ ] Complete Exchange Server integration  
- [ ] Add email composition and sending capabilities
- [ ] Add attachment handling
- [ ] Add calendar integration
- [ ] Add contacts integration
- [ ] Performance optimizations and caching
- [ ] Enhanced search capabilities
