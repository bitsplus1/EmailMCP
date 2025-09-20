"""Email data model with validation."""

from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional
import re
from .exceptions import ValidationError


@dataclass
class EmailData:
    """Data model for email information."""
    
    id: str
    subject: str
    sender: str
    sender_email: str
    recipients: List[str] = field(default_factory=list)
    cc_recipients: List[str] = field(default_factory=list)
    bcc_recipients: List[str] = field(default_factory=list)
    body: str = ""
    body_html: str = ""
    received_time: Optional[datetime] = None
    sent_time: Optional[datetime] = None
    is_read: bool = False
    has_attachments: bool = False
    importance: str = "Normal"
    folder_name: str = ""
    size: int = 0

    def __post_init__(self):
        """Validate email data after initialization."""
        self.validate()

    def validate(self) -> None:
        """Validate email data fields."""
        if not self.id or not isinstance(self.id, str):
            raise ValidationError("Email ID must be a non-empty string")
        
        if not isinstance(self.subject, str):
            raise ValidationError("Email subject must be a string")
        
        if not self.sender or not isinstance(self.sender, str):
            raise ValidationError("Email sender must be a non-empty string")
        
        if not self.sender_email or not self._is_valid_email(self.sender_email):
            raise ValidationError("Sender email must be a valid email address")
        
        # Validate recipient email addresses
        for recipient in self.recipients + self.cc_recipients + self.bcc_recipients:
            if not self._is_valid_email(recipient):
                raise ValidationError(f"Invalid recipient email address: {recipient}")
        
        if self.importance not in ["Low", "Normal", "High"]:
            raise ValidationError("Importance must be 'Low', 'Normal', or 'High'")
        
        if self.size < 0:
            raise ValidationError("Email size cannot be negative")

    @staticmethod
    def _is_valid_email(email: str) -> bool:
        """Validate email address format."""
        if not email or not isinstance(email, str):
            return False
        
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None

    @staticmethod
    def validate_email_id(email_id: str) -> bool:
        """Validate email ID format."""
        if not email_id or not isinstance(email_id, str):
            return False
        
        # Email IDs should be non-empty strings with reasonable length
        return len(email_id.strip()) > 0 and len(email_id) <= 255

    def to_dict(self) -> dict:
        """Convert email data to dictionary for JSON serialization."""
        return {
            "id": self.id,
            "subject": self.subject,
            "sender": self.sender,
            "sender_email": self.sender_email,
            "recipients": self.recipients,
            "cc_recipients": self.cc_recipients,
            "bcc_recipients": self.bcc_recipients,
            "body": self.body,
            "body_html": self.body_html,
            "received_time": self.received_time.isoformat() if self.received_time else None,
            "sent_time": self.sent_time.isoformat() if self.sent_time else None,
            "is_read": self.is_read,
            "has_attachments": self.has_attachments,
            "importance": self.importance,
            "folder_name": self.folder_name,
            "size": self.size
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'EmailData':
        """Create EmailData instance from dictionary."""
        # Convert datetime strings back to datetime objects
        received_time = None
        if data.get("received_time"):
            received_time = datetime.fromisoformat(data["received_time"])
        
        sent_time = None
        if data.get("sent_time"):
            sent_time = datetime.fromisoformat(data["sent_time"])
        
        return cls(
            id=data["id"],
            subject=data["subject"],
            sender=data["sender"],
            sender_email=data["sender_email"],
            recipients=data.get("recipients", []),
            cc_recipients=data.get("cc_recipients", []),
            bcc_recipients=data.get("bcc_recipients", []),
            body=data.get("body", ""),
            body_html=data.get("body_html", ""),
            received_time=received_time,
            sent_time=sent_time,
            is_read=data.get("is_read", False),
            has_attachments=data.get("has_attachments", False),
            importance=data.get("importance", "Normal"),
            folder_name=data.get("folder_name", ""),
            size=data.get("size", 0)
        )