"""Performance optimization components."""

# Import components individually to avoid circular imports
try:
    from .rate_limiter import RateLimiter, RateLimitConfig, TimeoutManager
except ImportError as e:
    print(f"Warning: Could not import rate_limiter: {e}")
    RateLimiter = RateLimitConfig = TimeoutManager = None

try:
    from .memory_manager import MemoryManager, MemoryConfig
except ImportError as e:
    print(f"Warning: Could not import memory_manager: {e}")
    MemoryManager = MemoryConfig = None

try:
    from .lazy_loader import LazyEmailLoader, LazyAttachmentLoader
except ImportError as e:
    print(f"Warning: Could not import lazy_loader: {e}")
    LazyEmailLoader = LazyAttachmentLoader = None

__all__ = [
    'RateLimiter',
    'RateLimitConfig', 
    'TimeoutManager',
    'MemoryManager',
    'MemoryConfig',
    'LazyEmailLoader',
    'LazyAttachmentLoader'
]