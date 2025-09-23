#!/usr/bin/env python3
"""
Debug script to examine raw COM properties of working vs failing emails.
"""

import sys
from pathlib import Path
import pythoncom
import win32com.client

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))


def debug_com_properties():
    """Debug raw COM properties of emails."""
    print("🔍 DEBUGGING RAW COM PROPERTIES")
    print("="*50)
    
    try:
        # Initialize COM
        pythoncom.CoInitialize()
        
        # Connect to Outlook
        print("📧 Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Get inbox
        inbox = namespace.GetDefaultFolder(6)  # olFolderInbox
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)  # Newest first
        
        print(f"✅ Found {items.Count} emails in inbox")
        
        # Find a working email and a failing email
        working_email = None
        failing_email = None
        
        for i in range(min(20, items.Count)):
            try:
                item = items.Item(i + 1)
                
                if not hasattr(item, 'Class') or item.Class != 43:  # Not a mail item
                    continue
                
                subject = str(item.Subject) if hasattr(item, 'Subject') else 'No Subject'
                size = getattr(item, 'Size', 0)
                
                # Look for our test email (failing)
                if "Body Extraction Debug Test" in subject and not failing_email:
                    failing_email = item
                    print(f"📧 Found failing email: {subject}")
                
                # Look for a working email (large size, has body)
                elif size > 100000 and not working_email:  # Large email likely has content
                    try:
                        body = getattr(item, 'Body', '')
                        if body and len(str(body)) > 100:
                            working_email = item
                            print(f"📧 Found working email: {subject}")
                    except:
                        pass
                
                if working_email and failing_email:
                    break
                    
            except Exception as e:
                print(f"⚠️  Error processing item {i}: {e}")
                continue
        
        # Analyze working email
        if working_email:
            print(f"\n✅ WORKING EMAIL ANALYSIS:")
            analyze_email_properties(working_email, "WORKING")
        else:
            print(f"\n❌ No working email found")
        
        # Analyze failing email
        if failing_email:
            print(f"\n❌ FAILING EMAIL ANALYSIS:")
            analyze_email_properties(failing_email, "FAILING")
        else:
            print(f"\n❌ No failing email found")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


def analyze_email_properties(email_item, label):
    """Analyze all properties of an email item."""
    print(f"\n🔬 {label} EMAIL PROPERTIES:")
    
    # Basic properties
    basic_props = [
        'Subject', 'SenderName', 'SenderEmailAddress', 'Size', 'ReceivedTime',
        'SentOn', 'UnRead', 'Importance', 'MessageClass', 'Body', 'HTMLBody'
    ]
    
    for prop in basic_props:
        try:
            if hasattr(email_item, prop):
                value = getattr(email_item, prop)
                if prop in ['Body', 'HTMLBody']:
                    if value:
                        print(f"  {prop}: EXISTS ({len(str(value))} chars) - '{str(value)[:100]}...'")
                    else:
                        print(f"  {prop}: EMPTY or None")
                else:
                    print(f"  {prop}: {value}")
            else:
                print(f"  {prop}: NOT AVAILABLE")
        except Exception as e:
            print(f"  {prop}: ERROR - {e}")
    
    # Try different body access methods
    print(f"\n🔧 BODY ACCESS METHODS:")
    
    # Method 1: Direct Body access
    try:
        body = email_item.Body
        print(f"  Direct Body: {len(str(body)) if body else 0} chars")
    except Exception as e:
        print(f"  Direct Body: ERROR - {e}")
    
    # Method 2: HTMLBody access
    try:
        html_body = email_item.HTMLBody
        print(f"  Direct HTMLBody: {len(str(html_body)) if html_body else 0} chars")
    except Exception as e:
        print(f"  Direct HTMLBody: ERROR - {e}")
    
    # Method 3: PropertyAccessor
    try:
        prop_accessor = email_item.PropertyAccessor
        body_prop = prop_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1000001F")
        print(f"  MAPI Body Property: {len(str(body_prop)) if body_prop else 0} chars")
    except Exception as e:
        print(f"  MAPI Body Property: ERROR - {e}")
    
    # Method 4: Check if item is fully loaded
    try:
        # Force loading by accessing multiple properties
        _ = email_item.Subject
        _ = email_item.Size
        _ = email_item.ReceivedTime
        _ = email_item.MessageClass
        
        # Now try body again
        body = email_item.Body
        print(f"  Post-refresh Body: {len(str(body)) if body else 0} chars")
    except Exception as e:
        print(f"  Post-refresh Body: ERROR - {e}")
    
    # Method 5: Check message state
    try:
        print(f"\n📋 MESSAGE STATE:")
        print(f"  Class: {getattr(email_item, 'Class', 'Unknown')}")
        print(f"  MessageClass: {getattr(email_item, 'MessageClass', 'Unknown')}")
        print(f"  Size: {getattr(email_item, 'Size', 0)} bytes")
        print(f"  Saved: {getattr(email_item, 'Saved', 'Unknown')}")
        
        # Check if it's a draft or unsaved item
        if hasattr(email_item, 'Saved'):
            saved = email_item.Saved
            print(f"  Is Saved: {saved}")
            if not saved:
                print(f"  ⚠️  EMAIL IS NOT SAVED - This might explain empty body!")
        
    except Exception as e:
        print(f"  Message State: ERROR - {e}")


if __name__ == "__main__":
    debug_com_properties()