#!/usr/bin/env python3
"""
Direct test of first 50 emails in inbox to check body content extraction.
This bypasses the MCP server and directly uses Outlook COM interface.
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


def test_email_body_extraction(namespace, limit=50):
    """Test body extraction for first N emails in inbox."""
    print(f"\n🎯 Testing body extraction for first {limit} emails in inbox")
    
    try:
        # Get Inbox folder
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        messages = inbox.Items
        
        print(f"📧 Found {messages.Count} total emails in Inbox")
        
        # Sort by received time (most recent first)
        messages.Sort("[ReceivedTime]", True)
        
        # Test results
        total_tested = 0
        empty_body_count = 0
        error_count = 0
        successful_extractions = 0
        
        # Test first N emails
        test_limit = min(limit, messages.Count)
        
        print(f"🔍 Testing first {test_limit} emails...")
        print("=" * 80)
        
        for i in range(test_limit):
            try:
                message = messages.Item(i + 1)  # COM collections are 1-indexed
                total_tested += 1
                
                # Get basic email info
                subject = str(message.Subject) if message.Subject else "No Subject"
                sender = str(message.SenderName) if message.SenderName else "Unknown"
                received = str(message.ReceivedTime) if message.ReceivedTime else "Unknown"
                
                print(f"\n--- Email {total_tested}/{test_limit} ---")
                print(f"Subject: {subject[:60]}...")
                print(f"From: {sender}")
                print(f"Received: {received}")
                
                # Test different body extraction methods
                body_methods = []
                
                # Method 1: Plain text body
                try:
                    if hasattr(message, 'Body') and message.Body:
                        plain_body = str(message.Body).strip()
                        if plain_body:
                            body_methods.append(("Plain Text Body", len(plain_body), plain_body[:100]))
                except Exception as e:
                    body_methods.append(("Plain Text Body", 0, f"Error: {e}"))
                
                # Method 2: HTML body
                try:
                    if hasattr(message, 'HTMLBody') and message.HTMLBody:
                        html_body = str(message.HTMLBody).strip()
                        if html_body:
                            # Simple HTML tag removal
                            import re
                            text_from_html = re.sub(r'<[^>]+>', '', html_body).strip()
                            if text_from_html:
                                body_methods.append(("HTML Body (converted)", len(text_from_html), text_from_html[:100]))
                except Exception as e:
                    body_methods.append(("HTML Body", 0, f"Error: {e}"))
                
                # Method 3: RTF body (if available)
                try:
                    if hasattr(message, 'RTFBody') and message.RTFBody:
                        rtf_body = str(message.RTFBody).strip()
                        if rtf_body:
                            body_methods.append(("RTF Body", len(rtf_body), rtf_body[:100]))
                except Exception as e:
                    body_methods.append(("RTF Body", 0, f"Error: {e}"))
                
                # Check results
                has_body_content = False
                for method_name, length, preview in body_methods:
                    if length > 0 and not preview.startswith("Error:"):
                        has_body_content = True
                        print(f"✅ {method_name}: {length} chars - {preview}...")
                    else:
                        print(f"❌ {method_name}: {preview}")
                
                if has_body_content:
                    successful_extractions += 1
                else:
                    empty_body_count += 1
                    print("⚠️  NO BODY CONTENT FOUND IN ANY METHOD!")
                    
                    # Additional debugging for empty emails
                    print("🔍 Additional properties:")
                    try:
                        print(f"   Size: {message.Size} bytes")
                        print(f"   Importance: {message.Importance}")
                        print(f"   MessageClass: {message.MessageClass}")
                        print(f"   HasAttachments: {message.Attachments.Count > 0}")
                    except Exception as debug_e:
                        print(f"   Debug error: {debug_e}")
                
            except Exception as e:
                error_count += 1
                print(f"❌ Error processing email {i+1}: {e}")
        
        # Summary
        print("\n" + "=" * 80)
        print("📊 TEST SUMMARY")
        print("=" * 80)
        print(f"Total emails tested: {total_tested}")
        print(f"Successful body extractions: {successful_extractions}")
        print(f"Emails with empty body: {empty_body_count}")
        print(f"Errors encountered: {error_count}")
        print(f"Success rate: {(successful_extractions/total_tested*100):.1f}%" if total_tested > 0 else "N/A")
        
        if empty_body_count > 0:
            print(f"\n⚠️  FOUND {empty_body_count} EMAILS WITH EMPTY BODY CONTENT!")
            print("This indicates potential issues with body content extraction.")
            print("These emails might be:")
            print("- Meeting invitations or calendar items")
            print("- System notifications without body text")
            print("- Encrypted or protected emails")
            print("- Corrupted email items")
        else:
            print("\n✅ ALL EMAILS HAVE BODY CONTENT!")
            print("Body extraction is working correctly for all tested emails.")
        
        return {
            'total_tested': total_tested,
            'successful_extractions': successful_extractions,
            'empty_body_count': empty_body_count,
            'error_count': error_count
        }
        
    except Exception as e:
        print(f"❌ Error during email testing: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    """Main function to test email body extraction."""
    print("🧪 EMAIL BODY EXTRACTION TEST")
    print("="*50)
    print("Testing body content extraction for first 50 emails in inbox")
    print("="*50)
    
    try:
        # Connect to Outlook
        outlook, namespace = connect_to_outlook()
        if not outlook or not namespace:
            print("\n❌ Cannot proceed without Outlook connection")
            return False
        
        # Test email body extraction
        results = test_email_body_extraction(namespace, limit=50)
        
        if results is None:
            print("\n❌ TEST FAILED")
            return False
        
        # Determine if we need to fix anything
        if results['empty_body_count'] > 0:
            print(f"\n🔧 ISSUES DETECTED - Need to fix body extraction!")
            print(f"Found {results['empty_body_count']} emails with empty body content.")
            return False
        else:
            print(f"\n✅ TEST PASSED - All emails have body content!")
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
        print("\n✅ BODY EXTRACTION TEST SUCCESSFUL!")
    else:
        print("\n❌ BODY EXTRACTION TEST FOUND ISSUES!")
        print("Will need to investigate and fix the body extraction methods.")
    
    input("\nPress Enter to exit...")