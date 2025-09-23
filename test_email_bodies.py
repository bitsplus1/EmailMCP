#!/usr/bin/env python3
"""
Test script to check email body content for the first 50 emails in inbox.
This will help identify any issues with body content extraction.
"""

import sys
import json
import asyncio
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.logging.logger import get_logger


async def test_email_bodies():
    """Test email body content extraction for first 50 emails."""
    logger = get_logger(__name__)
    
    # Create server config
    config = create_server_config(
        log_level="INFO",
        enable_console_output=True
    )
    
    # Create and start server
    server = OutlookMCPServer(config)
    
    try:
        await server.start()
        logger.info("Server started successfully")
        
        # Get first 50 emails from inbox
        logger.info("Fetching first 50 emails from inbox...")
        
        list_request = {
            "jsonrpc": "2.0",
            "id": "test_list",
            "method": "tools/call",
            "params": {
                "name": "list_inbox_emails",
                "arguments": {
                    "limit": 50
                }
            }
        }
        
        list_response = await server.handle_request(list_request)
        
        if "error" in list_response:
            logger.error(f"Error listing emails: {list_response['error']}")
            return
        
        emails = json.loads(list_response["result"]["content"][0]["text"])
        logger.info(f"Found {len(emails)} emails")
        
        # Test each email's body content
        empty_body_count = 0
        error_count = 0
        
        for i, email in enumerate(emails, 1):
            email_id = email.get("id")
            subject = email.get("subject", "No Subject")[:50]
            
            print(f"\n--- Email {i}/50 ---")
            print(f"ID: {email_id}")
            print(f"Subject: {subject}")
            
            # Check if body is already in the list response
            body_from_list = email.get("body", "").strip()
            if body_from_list:
                print(f"Body from list (length: {len(body_from_list)}): {body_from_list[:100]}...")
                continue
            else:
                print("Body from list: EMPTY")
            
            # Try to get full email details
            get_request = {
                "jsonrpc": "2.0",
                "id": f"test_get_{i}",
                "method": "tools/call",
                "params": {
                    "name": "get_email",
                    "arguments": {
                        "email_id": email_id
                    }
                }
            }
            
            try:
                get_response = await server.handle_request(get_request)
                
                if "error" in get_response:
                    print(f"Error getting email: {get_response['error']}")
                    error_count += 1
                    continue
                
                email_details = json.loads(get_response["result"]["content"][0]["text"])
                body_from_get = email_details.get("body", "").strip()
                
                if body_from_get:
                    print(f"Body from get_email (length: {len(body_from_get)}): {body_from_get[:100]}...")
                else:
                    print("Body from get_email: EMPTY")
                    empty_body_count += 1
                    
                    # Additional debugging info
                    print(f"Email details keys: {list(email_details.keys())}")
                    if "raw_properties" in email_details:
                        raw_props = email_details["raw_properties"]
                        print(f"Raw properties available: {list(raw_props.keys()) if isinstance(raw_props, dict) else 'Not a dict'}")
                
            except Exception as e:
                print(f"Exception getting email: {e}")
                error_count += 1
        
        # Summary
        print(f"\n=== SUMMARY ===")
        print(f"Total emails tested: {len(emails)}")
        print(f"Emails with empty body: {empty_body_count}")
        print(f"Errors encountered: {error_count}")
        
        if empty_body_count > 0:
            print(f"\n⚠️  Found {empty_body_count} emails with empty body content!")
            print("This indicates a bug in body content extraction that needs to be fixed.")
        else:
            print("\n✅ All emails have body content!")
            
    except Exception as e:
        logger.error(f"Test failed: {e}", exc_info=True)
        
    finally:
        await server.stop()


if __name__ == "__main__":
    asyncio.run(test_email_bodies())