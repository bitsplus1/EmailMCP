#!/usr/bin/env python3
"""
Test script to send an email and immediately read it back to debug body extraction.
"""

import requests
import json
import time


def test_send_and_read_email():
    """Send an email and immediately read it back to debug body extraction."""
    server_url = "http://192.168.1.164:8080/mcp"
    
    print("🧪 SEND AND READ EMAIL TEST")
    print("="*50)
    
    # Step 1: Send a test email
    print("📤 Sending test email...")
    
    send_request = {
        "jsonrpc": "2.0",
        "id": "send_test",
        "method": "send_email",
        "params": {
            "to_recipients": ["JackieCF_Lin@compal.com"],  # Send to yourself
            "subject": "Body Extraction Debug Test - " + str(int(time.time())),
            "body": "This is a test email body for debugging body extraction.\n\nLine 1: Hello World\nLine 2: Testing body content\nLine 3: This should be visible in the extracted body.\n\nEnd of test email.",
            "body_format": "text"
        }
    }
    
    try:
        response = requests.post(server_url, json=send_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error sending email: {result['error']}")
            return
        
        print("✅ Email sent successfully!")
        print(f"📋 Send result: {result}")
        
    except Exception as e:
        print(f"❌ Failed to send email: {e}")
        return
    
    # Step 2: Wait a moment for the email to appear
    print("\n⏳ Waiting 90 seconds for email to appear in inbox...")
    time.sleep(90)
    
    # Step 3: Get recent emails and look for our test email
    print("📧 Fetching recent emails to find our test email...")
    
    list_request = {
        "jsonrpc": "2.0",
        "id": "list_recent",
        "method": "list_inbox_emails",
        "params": {"limit": 20}
    }
    
    try:
        response = requests.post(server_url, json=list_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error listing emails: {result['error']}")
            return
        
        emails = result["result"]["emails"]
        print(f"✅ Retrieved {len(emails)} recent emails")
        
        # Find our test email
        test_email = None
        for email in emails:
            if "Body Extraction Debug Test" in email.get("subject", ""):
                test_email = email
                break
        
        if not test_email:
            print("❌ Could not find our test email in recent emails")
            print("📋 Recent email subjects:")
            for email in emails[:10]:
                print(f"  - {email.get('subject', 'No Subject')}")
            return
        
        print(f"✅ Found our test email!")
        print(f"📋 Test Email Analysis:")
        print(f"  Subject: {test_email['subject']}")
        print(f"  Sender: {test_email['sender']}")
        print(f"  Size: {test_email.get('size', 'Unknown')}")
        print(f"  Has Body: {test_email.get('has_body', 'Unknown')}")
        print(f"  Body Length: {len(test_email.get('body', ''))}")
        print(f"  Body HTML Length: {len(test_email.get('body_html', ''))}")
        print(f"  Body Preview: '{test_email.get('body_preview', '')}'")
        
        if test_email.get('body'):
            print(f"  Body Content: '{test_email['body'][:200]}...'")
        else:
            print(f"  Body Content: EMPTY!")
        
        if test_email.get('body_html'):
            print(f"  HTML Content: '{test_email['body_html'][:200]}...'")
        else:
            print(f"  HTML Content: EMPTY!")
        
        # Step 4: Try to get the email details using get_email
        print(f"\n🔍 Getting detailed email info using get_email...")
        
        get_request = {
            "jsonrpc": "2.0",
            "id": "get_test_email",
            "method": "get_email",
            "params": {"email_id": test_email['id']}
        }
        
        try:
            response = requests.post(server_url, json=get_request, timeout=30)
            response.raise_for_status()
            result = response.json()
            
            if "error" in result:
                print(f"❌ Error getting email details: {result['error']}")
            else:
                detailed_email = result["result"]
                print(f"✅ Got detailed email info")
                print(f"📋 Detailed Email Analysis:")
                print(f"  Body Length: {len(detailed_email.get('body', ''))}")
                print(f"  Body HTML Length: {len(detailed_email.get('body_html', ''))}")
                
                if detailed_email.get('body'):
                    print(f"  Detailed Body: '{detailed_email['body'][:200]}...'")
                else:
                    print(f"  Detailed Body: EMPTY!")
                    
        except Exception as e:
            print(f"❌ Exception getting detailed email: {e}")
        
        # Step 5: Analysis and conclusions
        print(f"\n🎯 ANALYSIS:")
        
        expected_body = "This is a test email body for debugging body extraction."
        
        if test_email.get('body') and expected_body in test_email['body']:
            print("✅ SUCCESS: Body content extracted correctly from list_inbox_emails")
        else:
            print("❌ PROBLEM: Body content not extracted from list_inbox_emails")
            print("   This confirms the body extraction bug!")
            
            # Additional debugging
            print(f"\n🔧 DEBUGGING INFO:")
            print(f"  Expected body to contain: '{expected_body}'")
            print(f"  Actual body: '{test_email.get('body', 'NONE')}'")
            print(f"  Email size: {test_email.get('size', 0)} bytes")
            print(f"  Has attachments: {test_email.get('has_attachments', False)}")
            print(f"  Message accessible: {test_email.get('accessible', 'Unknown')}")
        
    except Exception as e:
        print(f"❌ Exception during email retrieval: {e}")


if __name__ == "__main__":
    test_send_and_read_email()