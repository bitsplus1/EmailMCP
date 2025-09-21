@echo off
REM Simple batch file for n8n MCP testing
cd /d "C:\Users\USER\Documents\Github Repo\EmailMCP"
echo {"jsonrpc": "2.0", "id": "n8n-test", "method": "get_folders", "params": {}} | python main.py single-request