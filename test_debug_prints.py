#!/usr/bin/env python3
"""
Simple test to see if our debug prints are working.
"""

import requests
import json

def test_debug_prints():
    """Test if our debug prints are visible."""
    server_url = "http://192.168.1.164:8080/mcp"
    
    print("🔧 TESTING DEBUG PRINTS")
    print("="*50)
    
    # Get just 1 email to see debug output
    print("📧 Getting 1 email to see debug output...")
    
    list_request = {
        "jsonrpc": "2.0",
        "id": "debug_test",
        "method": "list_inbox_emails",
        "params": {"limit": 1}
    }
    
    try:
        response = requests.post(server_url, json=list_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error: {result['error']}")
            return
        
        emails = result["result"]["emails"]
        print(f"✅ Got {len(emails)} email(s)")
        
        if emails:
            email = emails[0]
            print(f"📋 Email subject: {email.get('subject', 'No Subject')[:60]}")
            print(f"📋 Email body length: {len(email.get('body', ''))}")
            print(f"📋 Email size: {email.get('size', 0)} bytes")
            
            if not email.get('body'):
                print("❌ This email has no body content - should trigger our debug prints!")
            else:
                print("✅ This email has body content")
        
        print("\n🔍 If you don't see debug prints above, the server needs to be restarted.")
        
    except Exception as e:
        print(f"❌ Exception: {e}")

if __name__ == "__main__":
    test_debug_prints()