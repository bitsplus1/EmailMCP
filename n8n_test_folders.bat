@echo off
REM Batch file for testing MCP server from n8n
REM This gets the list of Outlook folders

cd /d "C:\Users\USER\Documents\Github Repo\EmailMCP"
echo {"jsonrpc": "2.0", "id": "n8n-test", "method": "get_folders", "params": {}} | python main.py single-request