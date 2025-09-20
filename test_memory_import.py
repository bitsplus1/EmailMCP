#!/usr/bin/env python3

import sys
import os

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

try:
    # Try importing the memory manager directly
    from outlook_mcp_server.performance.memory_manager import MemoryManager, MemoryConfig
    print("✓ Direct import successful")
    
    # Test creating an instance
    config = MemoryConfig(max_memory_mb=64)
    manager = MemoryManager(config)
    print("✓ MemoryManager instance created successfully")
    
    # Test basic functionality
    stats = manager.get_stats()
    print(f"✓ Stats retrieved: {len(stats)} items")
    
    manager.shutdown()
    print("✓ Manager shutdown successful")
    
except Exception as e:
    print(f"✗ Error: {e}")
    import traceback
    traceback.print_exc()