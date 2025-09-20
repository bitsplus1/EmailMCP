"""Services package for Outlook MCP Server."""

from .folder_service import FolderService

# EmailService will be imported when it's implemented
# from .email_service import EmailService

__all__ = ['FolderService']