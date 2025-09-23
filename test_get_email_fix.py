#!/usr/bin/env python3
"""
Test script to verify the get_email fix.
"""

import requests
import json

def test_get_email_fix():
    """Test if get_email now works with the GetItemFromID fix."""
    server_url = "http://192.168.1.164:8080/mcp"
    
    print("🔧 TESTING GET_EMAIL FIX")
    print("="*50)
    
    # First get a list of emails to get a valid email ID
    print("📧 Getting email list to find a valid email ID...")
    
    list_request = {
        "jsonrpc": "2.0",
        "id": "list_test",
        "method": "list_inbox_emails",
        "params": {"limit": 3}
    }
    
    try:
        response = requests.post(server_url, json=list_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error getting email list: {result['error']}")
            return
        
        emails = result["result"]["emails"]
        print(f"✅ Got {len(emails)} emails from list")
        
        if not emails:
            print("❌ No emails found in list")
            return
        
        # Test get_email with the first email ID
        test_email = emails[0]
        email_id = test_email["id"]
        subject = test_email["subject"]
        
        print(f"🎯 Testing get_email with:")
        print(f"   ID: {email_id[:50]}...")
        print(f"   Subject: {subject[:50]}...")
        
        # Test get_email
        get_request = {
            "jsonrpc": "2.0",
            "id": "get_test",
            "method": "get_email",
            "params": {"email_id": email_id}
        }
        
        response = requests.post(server_url, json=get_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ get_email failed: {result['error']}")
            print(f"   Error code: {result['error'].get('code', 'unknown')}")
            print(f"   Error message: {result['error'].get('message', 'unknown')}")
        else:
            email_data = result["result"]
            print(f"✅ get_email SUCCESS!")
            print(f"   Subject: {email_data.get('subject', 'No Subject')[:50]}...")
            print(f"   Sender: {email_data.get('sender', 'Unknown')}")
            print(f"   Body length: {len(email_data.get('body', ''))}")
            print(f"   Size: {email_data.get('size', 0)} bytes")
            
            if email_data.get('body'):
                print(f"   Body preview: {email_data['body'][:100]}...")
            else:
                print(f"   Body: EMPTY")
        
    except Exception as e:
        print(f"❌ Exception: {e}")

if __name__ == "__main__":
    test_get_email_fix()