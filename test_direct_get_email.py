#!/usr/bin/env python3
"""
Direct test of get_email_by_id method without HTTP
"""

import sys
import os
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from outlook_mcp_server.adapters.outlook_adapter import OutlookAdapter

def test_direct_get_email():
    print("🔧 TESTING GET_EMAIL_BY_ID DIRECTLY")
    print("=" * 50)
    
    try:
        # Create adapter
        adapter = OutlookAdapter()
        
        # Connect
        print("📧 Connecting to Outlook...")
        adapter.connect()
        print("✅ Connected")
        
        # Get a list of emails first
        print("📧 Getting email list...")
        emails = adapter.list_inbox_emails(limit=1)
        
        if not emails:
            print("❌ No emails found")
            return
            
        email = emails[0]
        email_id = email.id
        subject = email.subject[:50]
        
        print(f"✅ Got email from list:")
        print(f"   ID: {email_id}")
        print(f"   Subject: {subject}...")
        
        # Now try to get the same email by ID
        print(f"\n📧 Calling get_email_by_id directly...")
        try:
            detailed_email = adapter.get_email_by_id(email_id)
            print(f"✅ get_email_by_id succeeded!")
            print(f"   Subject: {detailed_email.subject[:50]}...")
            print(f"   Body length: {len(detailed_email.body)}")
        except Exception as e:
            print(f"❌ get_email_by_id failed: {e}")
            print(f"   Error type: {type(e).__name__}")
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
        print(f"   Error type: {type(e).__name__}")
    finally:
        try:
            adapter.disconnect()
        except:
            pass

if __name__ == "__main__":
    test_direct_get_email()