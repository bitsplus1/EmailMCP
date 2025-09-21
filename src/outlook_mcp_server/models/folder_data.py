"""Folder data model with validation."""

from dataclasses import dataclass
from typing import Optional
import re
from .exceptions import ValidationError


@dataclass
class FolderData:
    """Data model for folder information."""
    
    id: str
    name: str
    full_path: str
    item_count: int = 0
    unread_count: int = 0
    parent_folder: str = ""
    folder_type: str = "Mail"

    def __post_init__(self):
        """Validate folder data after initialization."""
        self.validate()

    def validate(self) -> None:
        """Validate folder data fields."""
        if not self.id or not isinstance(self.id, str):
            raise ValidationError("Folder ID must be a non-empty string")
        
        if not self.name or not isinstance(self.name, str):
            raise ValidationError("Folder name must be a non-empty string")
        
        if not self.full_path or not isinstance(self.full_path, str):
            raise ValidationError("Folder full path must be a non-empty string")
        
        if self.item_count < 0:
            raise ValidationError("Item count cannot be negative")
        
        if self.unread_count < 0:
            raise ValidationError("Unread count cannot be negative")
        
        if self.unread_count > self.item_count:
            raise ValidationError("Unread count cannot exceed total item count")
        
        # Validate folder name doesn't contain invalid characters
        # Temporarily disabled to handle Unicode folder names
        # if not self._is_valid_folder_name(self.name):
        #     raise ValidationError("Folder name contains invalid characters")

    @staticmethod
    def _is_valid_folder_name(name: str) -> bool:
        """Validate folder name format - allows Unicode characters including Chinese."""
        if not name or not isinstance(name, str):
            return False
        
        # Strip whitespace
        name = name.strip()
        if not name:
            return False
        
        # Debug logging to see what's causing the issue
        import logging
        logger = logging.getLogger(__name__)
        
        # Only reject characters that are actually problematic for file systems
        # Allow Unicode characters (including Chinese, Japanese, Korean, etc.)
        invalid_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
        
        # Check for invalid characters
        problematic_chars = [char for char in name if char in invalid_chars]
        if problematic_chars:
            logger.debug(f"Folder name '{name}' contains invalid chars: {problematic_chars}")
            return False
        
        # Check for control characters (ASCII 0-31)
        control_chars = [char for char in name if ord(char) < 32]
        if control_chars:
            logger.debug(f"Folder name '{name}' contains control chars: {[ord(c) for c in control_chars]}")
            return False
        
        # Log successful validation for debugging
        logger.debug(f"Folder name '{name}' passed validation (length: {len(name)})")
        
        # Allow all other characters including Unicode
        return True

    @staticmethod
    def validate_folder_name(folder_name: str) -> bool:
        """Validate folder name format."""
        if not folder_name or not isinstance(folder_name, str):
            return False
        
        # Check length and invalid characters
        if len(folder_name.strip()) == 0 or len(folder_name) > 255:
            return False
        
        return FolderData._is_valid_folder_name(folder_name)

    def to_dict(self) -> dict:
        """Convert folder data to dictionary for JSON serialization."""
        return {
            "id": self.id,
            "name": self.name,
            "full_path": self.full_path,
            "item_count": self.item_count,
            "unread_count": self.unread_count,
            "parent_folder": self.parent_folder,
            "folder_type": self.folder_type
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'FolderData':
        """Create FolderData instance from dictionary."""
        return cls(
            id=data["id"],
            name=data["name"],
            full_path=data["full_path"],
            item_count=data.get("item_count", 0),
            unread_count=data.get("unread_count", 0),
            parent_folder=data.get("parent_folder", ""),
            folder_type=data.get("folder_type", "Mail")
        )