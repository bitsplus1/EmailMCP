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
        if not self._is_valid_folder_name(self.name):
            raise ValidationError("Folder name contains invalid characters")

    @staticmethod
    def _is_valid_folder_name(name: str) -> bool:
        """Validate folder name format."""
        if not name or not isinstance(name, str):
            return False
        
        # Folder names should not contain certain characters
        invalid_chars = ['<', '>', ':', '"', '|', '?', '*', '\\', '/']
        return not any(char in name for char in invalid_chars)

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