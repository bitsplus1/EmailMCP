#!/usr/bin/env python3
"""
Debug script to examine a working email's body extraction process.
"""

import requests
import json


def debug_working_email():
    """Debug the body extraction for a known working email."""
    server_url = "http://192.168.1.164:8080/mcp"
    
    print("🔍 DEBUGGING WORKING EMAIL BODY EXTRACTION")
    print("="*60)
    
    # Get recent emails
    print("📧 Fetching recent emails...")
    
    list_request = {
        "jsonrpc": "2.0",
        "id": "debug_list",
        "method": "list_inbox_emails",
        "params": {"limit": 20}
    }
    
    try:
        response = requests.post(server_url, json=list_request, timeout=30)
        response.raise_for_status()
        result = response.json()
        
        if "error" in result:
            print(f"❌ Error: {result['error']}")
            return
        
        emails = result["result"]["emails"]
        print(f"✅ Retrieved {len(emails)} emails")
        
        # Find emails with body content
        working_emails = []
        empty_emails = []
        
        for email in emails:
            body_length = len(email.get('body', ''))
            if body_length > 0:
                working_emails.append(email)
            else:
                empty_emails.append(email)
        
        print(f"\n📊 ANALYSIS:")
        print(f"  Working emails (with body): {len(working_emails)}")
        print(f"  Empty emails (no body): {len(empty_emails)}")
        
        if working_emails:
            print(f"\n✅ WORKING EMAILS:")
            for i, email in enumerate(working_emails[:3], 1):
                print(f"  {i}. Subject: {email['subject'][:60]}...")
                print(f"     Sender: {email['sender']}")
                print(f"     Body Length: {len(email['body'])} chars")
                print(f"     Size: {email.get('size', 'Unknown')} bytes")
                print(f"     Has Body Flag: {email.get('has_body', 'Unknown')}")
                print(f"     Body Preview: '{email['body'][:100]}...'")
                print()
        
        if empty_emails:
            print(f"\n❌ EMPTY EMAILS:")
            for i, email in enumerate(empty_emails[:3], 1):
                print(f"  {i}. Subject: {email['subject'][:60]}...")
                print(f"     Sender: {email['sender']}")
                print(f"     Body Length: {len(email['body'])} chars")
                print(f"     Size: {email.get('size', 'Unknown')} bytes")
                print(f"     Has Body Flag: {email.get('has_body', 'Unknown')}")
                print()
        
        # Detailed analysis of one working email
        if working_emails:
            test_email = working_emails[0]
            print(f"\n🔬 DETAILED ANALYSIS OF WORKING EMAIL:")
            print(f"Subject: {test_email['subject']}")
            print(f"All fields:")
            for key, value in test_email.items():
                if key == 'body':
                    print(f"  {key}: '{str(value)[:200]}...' (length: {len(str(value))})")
                elif key == 'body_html':
                    print(f"  {key}: '{str(value)[:200]}...' (length: {len(str(value))})")
                else:
                    print(f"  {key}: {value}")
        
        # Detailed analysis of one empty email
        if empty_emails:
            test_email = empty_emails[0]
            print(f"\n🔬 DETAILED ANALYSIS OF EMPTY EMAIL:")
            print(f"Subject: {test_email['subject']}")
            print(f"All fields:")
            for key, value in test_email.items():
                if key in ['body', 'body_html']:
                    print(f"  {key}: '{str(value)}' (length: {len(str(value))})")
                else:
                    print(f"  {key}: {value}")
        
        # Pattern analysis
        print(f"\n🎯 PATTERN ANALYSIS:")
        
        # Analyze by sender
        working_senders = set(email['sender'] for email in working_emails)
        empty_senders = set(email['sender'] for email in empty_emails)
        
        print(f"Working email senders: {list(working_senders)[:5]}")
        print(f"Empty email senders: {list(empty_senders)[:5]}")
        
        # Analyze by subject patterns
        working_subjects = [email['subject'] for email in working_emails]
        empty_subjects = [email['subject'] for email in empty_emails]
        
        print(f"\nWorking email subject patterns:")
        for subject in working_subjects[:3]:
            print(f"  - {subject}")
        
        print(f"\nEmpty email subject patterns:")
        for subject in empty_subjects[:3]:
            print(f"  - {subject}")
        
        # Size analysis
        working_sizes = [email.get('size', 0) for email in working_emails]
        empty_sizes = [email.get('size', 0) for email in empty_emails]
        
        if working_sizes:
            avg_working_size = sum(working_sizes) / len(working_sizes)
            print(f"\nAverage size of working emails: {avg_working_size:.0f} bytes")
        
        if empty_sizes:
            avg_empty_size = sum(empty_sizes) / len(empty_sizes)
            print(f"Average size of empty emails: {avg_empty_size:.0f} bytes")
        
    except Exception as e:
        print(f"❌ Exception: {e}")


if __name__ == "__main__":
    debug_working_email()