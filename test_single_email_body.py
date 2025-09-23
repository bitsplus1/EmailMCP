#!/usr/bin/env python3
"""
Test script to debug body extraction for a single email.
"""

import requests
import json
import time


def test_single_email_body():
    """Test body extraction for a single email with detailed debugging."""
    server_url = "http://192.168.1.164:8080/mcp"
    
    # First get the list of emails
    print("🔍 Getting list of emails...")
    
    list_request = {
        "jsonrpc": "2.0",
        "id": "get_emails",
        "method": "list_inbox_emails",
        "params": {"limit": 5}
    }
    
    response = requests.post(server_url, json=list_request, timeout=30)
    response.raise_for_status()
    result = response.json()
    
    if "error" in result:
        print(f"❌ Error: {result['error']}")
        return
    
    emails = result["result"]["emails"]
    print(f"✅ Got {len(emails)} emails")
    
    # Find an email with body content and one without
    email_with_body = None
    email_without_body = None
    
    for email in emails:
        body = email.get("body", "").strip()
        if body and not email_with_body:
            email_with_body = email
        elif not body and not email_without_body:
            email_without_body = email
        
        if email_with_body and email_without_body:
            break
    
    # Test both emails
    for test_email, label in [(email_with_body, "WITH BODY"), (email_without_body, "WITHOUT BODY")]:
        if not test_email:
            continue
            
        print(f"\n{'='*60}")
        print(f"🧪 TESTING EMAIL {label}")
        print(f"{'='*60}")
        print(f"ID: {test_email['id']}")
        print(f"Subject: {test_email['subject']}")
        print(f"Sender: {test_email['sender']}")
        print(f"Body from list (length): {len(test_email.get('body', ''))}")
        print(f"Body preview: {test_email.get('body', '')[:200]}...")
        
        # Get detailed email info
        print(f"\n🔍 Getting detailed email info...")
        
        get_request = {
            "jsonrpc": "2.0",
            "id": f"get_email_{test_email['id'][:10]}",
            "method": "get_email",
            "params": {"email_id": test_email['id']}
        }
        
        response = requests.post(server_url, json=get_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error getting detailed email: {result['error']}")
            continue
        
        # Try different response formats
        if "email" in result["result"]:
            detailed_email = result["result"]["email"]
        else:
            detailed_email = result["result"]
        detailed_body = detailed_email.get("body", "").strip()
        detailed_body_html = detailed_email.get("body_html", "").strip()
        
        print(f"✅ Got detailed email info")
        print(f"Body from get_email (length): {len(detailed_body)}")
        print(f"Body HTML from get_email (length): {len(detailed_body_html)}")
        print(f"Body preview: {detailed_body[:200]}...")
        print(f"Body HTML preview: {detailed_body_html[:200]}...")
        
        # Check all available fields
        print(f"\n📋 All available fields in detailed email:")
        for key, value in detailed_email.items():
            if isinstance(value, str):
                print(f"  {key}: {value[:100]}..." if len(str(value)) > 100 else f"  {key}: {value}")
            else:
                print(f"  {key}: {value}")


if __name__ == "__main__":
    test_single_email_body()