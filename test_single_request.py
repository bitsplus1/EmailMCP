#!/usr/bin/env python3
"""Test script to verify single-request mode works."""

import asyncio
import json
import sys
import subprocess
import time
from pathlib import Path

async def test_single_request():
    """Test the single-request mode."""
    print("🧪 Testing single-request mode...")
    
    # Prepare test request
    test_request = {
        "jsonrpc": "2.0",
        "id": "test-single",
        "method": "get_folders",
        "params": {}
    }
    
    request_json = json.dumps(test_request)
    print(f"📤 Sending request: {request_json}")
    
    try:
        # Start the process with a timeout
        start_time = time.time()
        
        process = subprocess.Popen(
            [sys.executable, "main.py", "single-request"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=Path(__file__).parent
        )
        
        # Send the request and wait for response with timeout
        try:
            stdout, stderr = process.communicate(input=request_json, timeout=30)
            end_time = time.time()
            
            print(f"⏱️  Process completed in {end_time - start_time:.2f} seconds")
            print(f"📥 Response: {stdout}")
            
            if stderr:
                print(f"⚠️  Stderr: {stderr}")
            
            # Try to parse the response
            try:
                response = json.loads(stdout.strip())
                if "result" in response:
                    print("✅ Single-request mode working correctly!")
                    return True
                elif "error" in response:
                    print(f"❌ Server returned error: {response['error']}")
                    return False
                else:
                    print(f"❓ Unexpected response format: {response}")
                    return False
            except json.JSONDecodeError as e:
                print(f"❌ Failed to parse response as JSON: {e}")
                print(f"Raw output: {repr(stdout)}")
                return False
                
        except subprocess.TimeoutExpired:
            print("❌ Process timed out after 30 seconds - single-request mode is hanging!")
            process.kill()
            return False
            
    except Exception as e:
        print(f"❌ Error running test: {e}")
        return False

if __name__ == "__main__":
    result = asyncio.run(test_single_request())
    sys.exit(0 if result else 1)