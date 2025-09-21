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


def search_for_target_email(namespace, target_title):
    """Search for the target email directly."""
    print(f"\n🎯 Searching for email with subject:")
    print(f"   '{target_title}'")
    
    try:
        # Get Inbox folder
        inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        messages = inbox.Items
        
        print(f"📧 Found {messages.Count} total emails in Inbox")
        print("🔍 Searching through emails...")
        
        # Search through messages
        found_email = None
        checked_count = 0
        similar_emails = []  # Store potentially similar emails
        
        # Sort by received time (most recent first)
        messages.Sort("[ReceivedTime]", True)
        
        for i in range(messages.Count):  # Check all emails (no limit)
            try:
                message = messages.Item(i + 1)  # COM collections are 1-indexed
                subject = str(message.Subject) if message.Subject else ""
                
                checked_count += 1
                
                # Show progress every 50 emails for unlimited search
                if checked_count % 50 == 0:
                    print(f"   Checked {checked_count} emails...")
                
                # Check if this is our target email
                # Use flexible matching - check if the target title is contained in the subject
                if target_title.lower() in subject.lower() or subject.lower() in target_title.lower():
                    # For more precise matching, also check key components
                    target_words = target_title.split()
                    subject_words = subject.split()
                    
                    # Count matching words (case insensitive)
                    matching_words = 0
                    for word in target_words:
                        if len(word) > 2:  # Skip very short words
                            for sub_word in subject_words:
                                if word.lower() in sub_word.lower() or sub_word.lower() in word.lower():
                                    matching_words += 1
                                    break
                    
                    # Consider it a match if most words match or exact substring match
                    word_match_ratio = matching_words / len([w for w in target_words if len(w) > 2])
                    exact_match = target_title.lower() in subject.lower() or subject.lower() in target_title.lower()
                    
                    if exact_match or word_match_ratio > 0.6:
                        print(f"\n✅ FOUND TARGET EMAIL!")
                        print(f"   Subject: {subject}")
                        print(f"   From: {message.SenderName if message.SenderName else 'Unknown'}")
                        print(f"   Received: {message.ReceivedTime if message.ReceivedTime else 'Unknown'}")
                        
                        if not exact_match:
                            print(f"   Match Type: Fuzzy match ({word_match_ratio:.1%} word similarity)")
                        else:
                            print(f"   Match Type: Exact match")
                        
                        found_email = message
                        break
                    elif word_match_ratio > 0.3:  # Store similar emails
                        similar_emails.append({
                            'message': message,
                            'subject': subject,
                            'similarity': word_match_ratio
                        })
                    
                # Show first few emails for debugging
                elif checked_count <= 10:
                    sender = message.SenderName if message.SenderName else "Unknown"
                    received = message.ReceivedTime if message.ReceivedTime else "Unknown"
                    print(f"   [{checked_count}] {subject[:70]}... (from: {sender})")
                    
            except Exception as e:
                print(f"   ⚠️  Error reading email {i+1}: {e}")
                continue
        
        print(f"\n📊 Search completed - checked {checked_count} emails")
        
        # If no exact match found, show similar emails
        if not found_email and similar_emails:
            print(f"\n🔍 No exact match found, but found {len(similar_emails)} similar emails:")
            # Sort by similarity
            similar_emails.sort(key=lambda x: x['similarity'], reverse=True)
            
            for i, email_info in enumerate(similar_emails[:5], 1):  # Show top 5
                subject = email_info['subject']
                similarity = email_info['similarity']
                sender = email_info['message'].SenderName if email_info['message'].SenderName else 'Unknown'
                print(f"   {i}. {subject[:70]}... ({similarity:.1%} similar, from: {sender})")
            
            print(f"\n❓ Would you like to use one of these similar emails instead?")
            print(f"   You can modify the 'target_mail_title' variable in the script.")
        
        return found_email
        
    except Exception as e:
        print(f"❌ Error searching emails: {e}")
        return None


def extract_latest_email_content(message):
    """Extract only the latest email content from email body, excluding email thread history."""
    try:
        print(f"\n📧 Extracting latest email content...")
        
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
            return None
        
        print(f"✅ Email body retrieved ({len(body)} characters)")
        
        # Extract only the latest email content
        latest_content = extract_latest_mail_from_thread(body)
        
        if not latest_content:
            print("⚠️  Could not extract latest content, using full body")
            latest_content = body
        
        print(f"✅ Latest content extracted ({len(latest_content)} characters)")
        
        # Return the complete latest content (all lines before thread separator)
        return latest_content.strip()
        
    except Exception as e:
        print(f"❌ Error extracting email content: {e}")
        return None


def extract_latest_mail_from_thread(email_body):
    """
    Extract only the latest email content from an email thread.
    
    This function identifies and extracts the most recent email content while
    excluding previous emails in the thread history. It also handles cases
    where the sender mentions inline updates.
    """
    import re
    
    # Common patterns that indicate the start of previous emails in thread
    thread_separators = [
        r'-----Original Message-----',
        r'From:.*Sent:.*To:.*Subject:',
        r'On .* wrote:',
        r'On .* at .* wrote:',
        r'________________________________',
        r'From: .*\nSent: .*\nTo: .*\nSubject:',
        r'> .*',  # Quoted text (lines starting with >)
        r'^\s*From\s*:',
        r'^\s*Sent\s*:',
        r'^\s*To\s*:',
        r'^\s*Subject\s*:',
        r'Begin forwarded message:',
        r'---------- Forwarded message ----------',
    ]
    
    # Keywords that indicate inline updates (need to include more content)
    inline_update_keywords = [
        'update inline below',
        'updated inline below',
        'update in below table',
        'updated in below table',
        'see inline',
        'inline below',
        'comments inline',
        'responses inline',
        'answers inline',
        'updated below',
        'changes below',
        'modifications below',
        'edits below',
        'revisions below',
        'please see below',
        'see my comments below',
        'my responses below'
    ]
    
    # Split email into lines
    lines = email_body.split('\n')
    
    # Check if the latest email mentions inline updates
    has_inline_updates = False
    first_50_lines = '\n'.join(lines[:50]).lower()  # Check first 50 lines for inline update mentions
    
    for keyword in inline_update_keywords:
        if keyword in first_50_lines:
            has_inline_updates = True
            print(f"🔍 Detected inline update keyword: '{keyword}'")
            break
    
    if has_inline_updates:
        print("📝 Email contains inline updates - including more content")
        # If inline updates are mentioned, be more conservative about cutting content
        # Look for stronger separators only
        strong_separators = [
            r'-----Original Message-----',
            r'From:.*Sent:.*To:.*Subject:',
            r'________________________________',
        ]
        
        for i, line in enumerate(lines):
            for separator_pattern in strong_separators:
                if re.search(separator_pattern, line, re.IGNORECASE | re.MULTILINE):
                    print(f"✂️  Found strong separator at line {i+1}: {line[:50]}...")
                    return '\n'.join(lines[:i]).strip()
        
        # If no strong separator found, return more content (first 70% of email)
        cutoff_line = int(len(lines) * 0.7)
        print(f"📏 No strong separator found, returning first {cutoff_line} lines")
        return '\n'.join(lines[:cutoff_line]).strip()
    
    else:
        print("📝 Standard email - extracting latest content only")
        # Standard processing - look for any thread separator
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            
            # Skip empty lines
            if not line_stripped:
                continue
            
            # Check against all separator patterns
            for separator_pattern in thread_separators:
                if re.search(separator_pattern, line, re.IGNORECASE | re.MULTILINE):
                    print(f"✂️  Found thread separator at line {i+1}: {line[:50]}...")
                    return '\n'.join(lines[:i]).strip()
            
            # Check for quoted text (lines starting with >)
            if line_stripped.startswith('>'):
                print(f"✂️  Found quoted text at line {i+1}")
                return '\n'.join(lines[:i]).strip()
            
            # Check for email headers in the middle of content
            if (line_stripped.startswith('From:') or 
                line_stripped.startswith('Sent:') or 
                line_stripped.startswith('To:') or 
                line_stripped.startswith('Subject:')):
                # Make sure this looks like an email header (has email address or date)
                if '@' in line or re.search(r'\d{1,2}/\d{1,2}/\d{4}', line):
                    print(f"✂️  Found email header at line {i+1}: {line[:50]}...")
                    return '\n'.join(lines[:i]).strip()
    
    # If no separator found, return the full content
    print("📄 No thread separator found, returning full content")
    return email_body.strip()


def main():
    """Main function to test direct Outlook access."""
    # ========================================
    # CONFIGURABLE TARGET EMAIL TITLE
    # ========================================
    # Change this variable to test different email titles
    #target_mail_title = "On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result"
    target_mail_title = "密碼即將在 2 天後過期"
    
    # Other examples you can test:
    # target_mail_title = "Meeting Reminder"
    # target_mail_title = "Project Update"
    # target_mail_title = "Weekly Report"
    # target_mail_title = "Test Email"
    
    print("🧪 DIRECT OUTLOOK COM TEST")
    print("="*50)
    print("Testing direct COM interface connection to Outlook")
    print(f"Target: Find email with '{target_mail_title}'")
    print("Goal: Extract first two lines of email body")
    print("="*50)
    
    try:
        # Step 1: Connect to Outlook
        outlook, namespace = connect_to_outlook()
        if not outlook or not namespace:
            print("\n❌ Cannot proceed without Outlook connection")
            return False
        
        # Step 2: Search for target email
        target_email = search_for_target_email(namespace, target_mail_title)
        
        if not target_email:
            print("\n❌ TARGET EMAIL NOT FOUND")
            print("\n🔧 Troubleshooting:")
            print("   1. Make sure the email is in your Inbox folder")
            print("   2. Check if the subject line is exactly as specified")
            print("   3. The email might be in Sent Items if you sent it")
            print("   4. Try searching manually in Outlook to verify it exists")
            return False
        
        # Step 3: Extract latest email content
        latest_email_content = extract_latest_email_content(target_email)
        
        if not latest_email_content:
            print("\n❌ FAILED TO EXTRACT EMAIL CONTENT")
            return False
        
        # Step 4: Display results
        print("\n" + "="*70)
        print("🎉 TEST SUCCESSFUL!")
        print("="*70)
        print(f"✅ Found target email in Outlook")
        print(f"✅ Successfully extracted latest email content")
        print(f"✅ Retrieved complete latest email content (excluding thread history)")
        
        print(f"\n📋 RESULT - COMPLETE LATEST EMAIL CONTENT:")
        print("=" * 80)
        print(latest_email_content)
        print("=" * 80)
        
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