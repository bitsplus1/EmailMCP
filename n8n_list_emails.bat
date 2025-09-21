@echo off
REM Batch file for listing emails from Inbox
REM Usage: This lists the first 5 emails from Inbox

cd /d "C:\Users\USER\Documents\Github Repo\EmailMCP"
echo {"jsonrpc": "2.0", "id": "n8n-emails", "method": "list_emails", "params": {"folder": "Inbox", "limit": 5}} | python main.py single-request