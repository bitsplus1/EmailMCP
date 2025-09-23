#!/usr/bin/env python3
"""
Debug test for get_email method - shows exact IDs and timing
"""

import requests
import json
import time

def test_get_email_debug():
    print("🔧 DEBUGGING GET_EMAIL METHOD")
    print("=" * 60)
    
    base_url = "http://192.168.1.164:8080/mcp"
    
    # First, get a list of emails
    print("📧 Step 1: Getting email list...")
    list_request = {
        "jsonrpc": "2.0",
        "id": "list_debug",
        "method": "list_inbox_emails",
        "params": {"limit": 1}  # Just get 1 email
    }
    
    try:
        response = requests.post(base_url, json=list_request, timeout=30)
        response.raise_for_status()
        list_result = response.json()
        
        if "error" in list_result:
            print(f"❌ List failed: {list_result['error']}")
            return
            
        emails = list_result.get("result", {}).get("emails", [])
        if not emails:
            print("❌ No emails found in list")
            return
            
        email = emails[0]
        email_id = email.get("id")
        subject = email.get("subject", "")[:50]
        
        print(f"✅ Got email from list:")
        print(f"   ID: {email_id}")
        print(f"   Subject: {subject}...")
        print(f"   ID Length: {len(email_id) if email_id else 0}")
        
        # Wait a moment to ensure server processes the list request
        print("\n⏱️  Waiting 2 seconds...")
        time.sleep(2)
        
        # Now try to get the same email by ID
        print(f"\n📧 Step 2: Getting email by ID...")
        get_request = {
            "jsonrpc": "2.0", 
            "id": "get_debug",
            "method": "get_email",
            "params": {"email_id": email_id}
        }
        
        print(f"🎯 Requesting email with exact ID: {email_id}")
        
        response = requests.post(base_url, json=get_request, timeout=30)
        response.raise_for_status()
        get_result = response.json()
        
        if "error" in get_result:
            print(f"❌ get_email failed:")
            print(f"   Error: {get_result['error']}")
            
            # Check if it's the same ID
            error_id = get_result.get('error', {}).get('data', {}).get('details', {}).get('email_id', '')
            print(f"   Requested ID: {email_id}")
            print(f"   Error ID:     {error_id}")
            print(f"   IDs match:    {email_id == error_id}")
        else:
            email_data = get_result.get("result", {})
            print(f"✅ get_email succeeded!")
            print(f"   Subject: {email_data.get('subject', '')[:50]}...")
            print(f"   Body length: {len(email_data.get('body', ''))}")
            
    except Exception as e:
        print(f"❌ Request failed: {e}")

if __name__ == "__main__":
    test_get_email_debug()