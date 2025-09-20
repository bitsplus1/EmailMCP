#!/usr/bin/env python3
"""
Direct Outlook Test - Bypassing MCP Server

This script directly connects to Outlook COM interface to find your email.
"""

import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

try:
    import win32com.client
    print("✅ pywin32 is available")
except ImportError:
    print("❌ pywin32 not available. Please install: pip install pywin32")
    sys.exit(1)


def connect_to_outlook():
    """Connect directly to Outlook COM interface."""
    try:
        print("🔧 Connecting to Outlook COM interface...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        print("✅ Connected to Outlook successfully!")
        return outlook, namespace
    except Exception as e:
        print(f"❌ Failed to connect to Outlook: {e}")
        return None, None


def search_for_target_email(namespace):
    """Search for the target email directly."""
    target_subject = "On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result"
    
    print(f"\n🎯 Searching for email with subject:")
    print(f"   '{target_subject}'")
    
    try:
        # Get Inbox folder
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        messages = inbox.Items
        
        print(f"📧 Found {messages.Count} total emails in Inbox")
        print("🔍 Searching through emails...")
        
        # Search through messages
        found_email = None
        checked_count = 0
        
        # Sort by received time (most recent first)
        messages.Sort("[ReceivedTime]", True)
        
        for i in range(min(messages.Count, 200)):  # Check first 200 emails
            try:
                message = messages.Item(i + 1)  # COM collections are 1-indexed
                subject = str(message.Subject) if message.Subject else ""
                
                checked_count += 1
                
                # Show progress every 20 emails
                if checked_count % 20 == 0:
                    print(f"   Checked {checked_count} emails...")
                
                # Check if this is our target email
                if ("NPI KDP61" in subject and "DVT1.0_Build" in subject and 
                    "EQM1 Test Result" in subject and "9/20" in subject):
                    
                    print(f"\n✅ FOUND TARGET EMAIL!")
                    print(f"   Subject: {subject}")
                    print(f"   From: {message.SenderName if message.SenderName else 'Unknown'}")
                    print(f"   Received: {message.ReceivedTime if message.ReceivedTime else 'Unknown'}")
                    
                    found_email = message
                    break
                    
                # Show first few emails for debugging
                elif checked_count <= 10:
                    sender = message.SenderName if message.SenderName else "Unknown"
                    received = message.ReceivedTime if message.ReceivedTime else "Unknown"
                    print(f"   [{checked_count}] {subject[:70]}... (from: {sender})")
                    
            except Exception as e:
                print(f"   ⚠️  Error reading email {i+1}: {e}")
                continue
        
        print(f"\n📊 Search completed - checked {checked_count} emails")
        return found_email
        
    except Exception as e:
        print(f"❌ Error searching emails: {e}")
        return None


def extract_email_content(message):
    """Extract the first two lines from email body."""
    try:
        print(f"\n📧 Extracting email content...")
        
        # Get email body
        body = ""
        if hasattr(message, 'Body') and message.Body:
            body = str(message.Body)
        elif hasattr(message, 'HTMLBody') and message.HTMLBody:
            # If no plain text body, try to get text from HTML
            html_body = str(message.HTMLBody)
            # Simple HTML tag removal (basic)
            import re
            body = re.sub(r'<[^>]+>', '', html_body)
        
        if not body:
            print("❌ No email body content found")
            return None, None
        
        print(f"✅ Email body retrieved ({len(body)} characters)")
        
        # Split into lines and get first two non-empty lines
        lines = body.strip().split('\n')
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        
        first_line = non_empty_lines[0] if len(non_empty_lines) > 0 else ""
        second_line = non_empty_lines[1] if len(non_empty_lines) > 1 else ""
        
        return first_line, second_line
        
    except Exception as e:
        print(f"❌ Error extracting email content: {e}")
        return None, None


def main():
    """Main function to test direct Outlook access."""
    print("🧪 DIRECT OUTLOOK COM TEST")
    print("="*50)
    print("Testing direct COM interface connection to Outlook")
    print("Target: Find email with 'On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result'")
    print("Goal: Extract first two lines of email body")
    print("="*50)
    
    try:
        # Step 1: Connect to Outlook
        outlook, namespace = connect_to_outlook()
        if not outlook or not namespace:
            print("\n❌ Cannot proceed without Outlook connection")
            return False
        
        # Step 2: Search for target email
        target_email = search_for_target_email(namespace)
        
        if not target_email:
            print("\n❌ TARGET EMAIL NOT FOUND")
            print("\n🔧 Troubleshooting:")
            print("   1. Make sure the email is in your Inbox folder")
            print("   2. Check if the subject line is exactly as specified")
            print("   3. The email might be in Sent Items if you sent it")
            print("   4. Try searching manually in Outlook to verify it exists")
            return False
        
        # Step 3: Extract email content
        first_line, second_line = extract_email_content(target_email)
        
        if first_line is None:
            print("\n❌ FAILED TO EXTRACT EMAIL CONTENT")
            return False
        
        # Step 4: Display results
        print("\n" + "="*70)
        print("🎉 TEST SUCCESSFUL!")
        print("="*70)
        print(f"✅ Found target email in Outlook")
        print(f"✅ Successfully extracted email content")
        print(f"✅ Retrieved first two lines")
        
        print(f"\n📋 RESULT - FIRST TWO LINES:")
        print("=" * 60)
        print(f"Line 1: {first_line}")
        print(f"Line 2: {second_line}")
        print("=" * 60)
        
        print(f"\n📊 Email Details:")
        print(f"   Subject: {target_email.Subject}")
        print(f"   From: {target_email.SenderName}")
        print(f"   Received: {target_email.ReceivedTime}")
        print(f"   Size: {target_email.Size} bytes")
        
        print(f"\n🎯 MISSION ACCOMPLISHED!")
        print("Successfully found the email and extracted the first two lines!")
        
        return True
        
    except KeyboardInterrupt:
        print("\n⏹️  Test interrupted by user")
        return False
    except Exception as e:
        print(f"\n❌ Test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = main()
    
    if success:
        print("\n✅ TASK CONSIDERED SUCCESSFUL!")
        print("The direct Outlook COM interface successfully found your email")
        print("and extracted the first two lines from the email body.")
    else:
        print("\n❌ TASK INCOMPLETE")
        print("Could not find or process the target email.")
    
    input("\nPress Enter to exit...")