"""
Configuration management for EmailMCP.

This module handles loading and managing configuration settings for the EmailMCP server.
"""

import os
from typing import Optional, Dict, Any
from pydantic import BaseModel, Field
from dotenv import load_dotenv
import json


class OutlookConfig(BaseModel):
    """Configuration for Outlook/Office 365 connection."""
    
    client_id: Optional[str] = None
    client_secret: Optional[str] = None
    tenant_id: Optional[str] = None
    redirect_uri: str = "http://localhost:8080/callback"
    scopes: list[str] = Field(default_factory=lambda: [
        "https://graph.microsoft.com/Mail.Read",
        "https://graph.microsoft.com/Mail.ReadWrite",
        "https://graph.microsoft.com/User.Read"
    ])


class ExchangeConfig(BaseModel):
    """Configuration for Exchange server connection."""
    
    server: Optional[str] = None
    username: Optional[str] = None
    password: Optional[str] = None
    domain: Optional[str] = None
    autodiscover: bool = True


class ServerConfig(BaseModel):
    """General server configuration."""
    
    log_level: str = "INFO"
    max_emails_per_request: int = 100
    cache_ttl: int = 300  # seconds
    timeout: int = 30  # seconds


class EmailMCPConfig(BaseModel):
    """Main configuration class for EmailMCP."""
    
    outlook: OutlookConfig = Field(default_factory=OutlookConfig)
    exchange: ExchangeConfig = Field(default_factory=ExchangeConfig)
    server: ServerConfig = Field(default_factory=ServerConfig)
    
    @classmethod
    def load_from_env(cls) -> "EmailMCPConfig":
        """Load configuration from environment variables."""
        # Load .env file if it exists
        load_dotenv()
        
        outlook_config = OutlookConfig(
            client_id=os.getenv("OUTLOOK_CLIENT_ID"),
            client_secret=os.getenv("OUTLOOK_CLIENT_SECRET"),
            tenant_id=os.getenv("OUTLOOK_TENANT_ID"),
            redirect_uri=os.getenv("OUTLOOK_REDIRECT_URI", "http://localhost:8080/callback"),
        )
        
        exchange_config = ExchangeConfig(
            server=os.getenv("EXCHANGE_SERVER"),
            username=os.getenv("EXCHANGE_USERNAME"),
            password=os.getenv("EXCHANGE_PASSWORD"),
            domain=os.getenv("EXCHANGE_DOMAIN"),
            autodiscover=os.getenv("EXCHANGE_AUTODISCOVER", "true").lower() == "true"
        )
        
        server_config = ServerConfig(
            log_level=os.getenv("LOG_LEVEL", "INFO"),
            max_emails_per_request=int(os.getenv("MAX_EMAILS_PER_REQUEST", "100")),
            cache_ttl=int(os.getenv("CACHE_TTL", "300")),
            timeout=int(os.getenv("TIMEOUT", "30"))
        )
        
        return cls(
            outlook=outlook_config,
            exchange=exchange_config,
            server=server_config
        )
    
    @classmethod
    def load_from_file(cls, config_path: str) -> "EmailMCPConfig":
        """Load configuration from a JSON file."""
        with open(config_path, 'r') as f:
            config_data = json.load(f)
        return cls(**config_data)
    
    def save_to_file(self, config_path: str) -> None:
        """Save configuration to a JSON file."""
        with open(config_path, 'w') as f:
            json.dump(self.model_dump(), f, indent=2)


# Global configuration instance
config = EmailMCPConfig.load_from_env()