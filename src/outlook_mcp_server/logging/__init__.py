"""
Logging module for Outlook MCP Server.

This module provides structured JSON logging with rotation and performance metrics.
"""

from .logger import Logger, get_logger, configure_logging

__all__ = ['Logger', 'get_logger', 'configure_logging']