#!/usr/bin/env python3
"""
Real Outlook MCP Server Test

This script tests the MCP server against your actual Outlook instance
to find the specific email you mentioned and extract the first two lines.

Target Email: "On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result"
"""

import asyncio
import json
import sys
from pathlib import Path
from typing import Dict, Any, Optional

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.logging.logger import get_logger


class RealOutlookTester:
    """Tests the MCP server against real Outlook data."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
        self.server = None
        self.request_id = 1
    
    async def initialize_server(self) -> bool:
        """Initialize the MCP server with real Outlook connection."""
        try:
            print("🔧 Initializing Outlook MCP Server...")
            print("   Connecting to your Outlook instance...")
            
            # Create configuration with longer timeout for real Outlook
            config = create_server_config(
                log_level="INFO",
                enable_console_output=True,
                outlook_connection_timeout=30,  # Longer timeout for real connection
                max_concurrent_requests=5
            )
            
            self.server = OutlookMCPServer(config)
            await self.server.start()
            
            print("✅ Server connected to Outlook successfully!")
            return True
            
        except Exception as e:
            print(f"❌ Failed to connect to Outlook: {e}")
            print("\n🔧 Troubleshooting tips:")
            print("   1. Make sure Microsoft Outlook is running")
            print("   2. Check if Outlook is not in safe mode")
            print("   3. Verify Outlook is properly configured with an account")
            print("   4. Try closing and reopening Outlook")
            return False
    
    async def cleanup_server(self) -> None:
        """Cleanup server resources."""
        if self.server:
            await self.server.stop()
            self.server = None
    
    def get_next_request_id(self) -> str:
        """Get next request ID."""
        request_id = str(self.request_id)
        self.request_id += 1
        return request_id
    
    async def send_mcp_request(self, method: str, params: Dict[str, Any] = None) -> Dict[str, Any]:
        """Send MCP request to server."""
        request = {
            "jsonrpc": "2.0",
            "id": self.get_next_request_id(),
            "method": method,
            "params": params or {}
        }
        
        print(f"\n📤 Sending MCP Request:")
        print(f"   Method: {method}")
        print(f"   Request ID: {request['id']}")
        if params:
            print(f"   Parameters: {json.dumps(params, indent=6)}")
        
        try:
            response = await self.server.handle_request(request)
            
            print(f"\n📥 MCP Response received:")
            if "result" in response:
                result = response["result"]
                if isinstance(result, list):
                    print(f"   Status: SUCCESS - Found {len(result)} items")
                else:
                    print(f"   Status: SUCCESS")
            elif "error" in response:
                error = response["error"]
                print(f"   Status: ERROR - {error.get('message', 'Unknown error')}")
            
            return response
            
        except Exception as e:
            error_response = {
                "jsonrpc": "2.0",
                "id": request["id"],
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {str(e)}"
                }
            }
            
            print(f"\n❌ Request failed: {e}")
            return error_response
    
    async def search_target_email(self) -> Optional[Dict[str, Any]]:
        """Search for the specific target email."""
        print("\n" + "="*70)
        print("🎯 SEARCHING FOR TARGET EMAIL")
        print("="*70)
        print("Target: 'On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result'")
        
        # Try multiple search strategies
        search_strategies = [
            {
                "name": "Exact subject search",
                "params": {
                    "query": "On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result",
                    "limit": 10
                }
            },
            {
                "name": "Partial subject search",
                "params": {
                    "query": "NPI KDP61 DVT1.0_Build EQM1",
                    "limit": 20
                }
            },
            {
                "name": "Date and keyword search",
                "params": {
                    "query": "9/20 KDP61 Test Result",
                    "limit": 20
                }
            },
            {
                "name": "Recent inbox search",
                "method": "list_emails",
                "params": {
                    "folder": "Inbox",
                    "limit": 50
                }
            }
        ]
        
        target_email = None
        
        for strategy in search_strategies:
            print(f"\n🔍 Strategy: {strategy['name']}")
            
            method = strategy.get("method", "search_emails")
            response = await self.send_mcp_request(method, strategy["params"])
            
            if "result" in response:
                emails = response["result"]
                if method == "list_emails":
                    emails = emails  # list_emails returns emails directly
                
                print(f"   Found {len(emails)} emails to check")
                
                # Look for the target email
                for email in emails:
                    subject = email.get("subject", "").strip()
                    print(f"   📧 Checking: {subject[:60]}...")
                    
                    # Check if this matches our target
                    if "NPI KDP61 DVT1.0_Build EQM1 Test Result" in subject and "9/20" in subject:
                        print(f"   ✅ FOUND TARGET EMAIL!")
                        print(f"      Subject: {subject}")
                        print(f"      From: {email.get('sender', 'Unknown')}")
                        print(f"      Date: {email.get('received_time', 'Unknown')}")
                        target_email = email
                        break
                
                if target_email:
                    break
            else:
                error = response.get("error", {})
                print(f"   ❌ Strategy failed: {error.get('message', 'Unknown error')}")
        
        return target_email
    
    async def get_email_content(self, email_id: str) -> Optional[str]:
        """Get the full content of an email."""
        print(f"\n📧 RETRIEVING EMAIL CONTENT")
        print("-" * 40)
        
        response = await self.send_mcp_request("get_email", {"email_id": email_id})
        
        if "result" in response:
            email_details = response["result"]
            body = email_details.get("body", "")
            
            print(f"✅ Email content retrieved successfully")
            print(f"   Body length: {len(body)} characters")
            
            return body
        else:
            error = response.get("error", {})
            print(f"❌ Failed to get email content: {error.get('message', 'Unknown error')}")
            return None
    
    def extract_first_two_lines(self, email_body: str) -> tuple:
        """Extract the first two lines from email body."""
        if not email_body:
            return None, None
        
        # Clean up the body and split into lines
        lines = email_body.strip().split('\n')
        
        # Filter out empty lines and get first two non-empty lines
        non_empty_lines = [line.strip() for line in lines if line.strip()]
        
        first_line = non_empty_lines[0] if len(non_empty_lines) > 0 else ""
        second_line = non_empty_lines[1] if len(non_empty_lines) > 1 else ""
        
        return first_line, second_line
    
    async def run_test(self) -> bool:
        """Run the complete test to find and extract email content."""
        print("🧪 OUTLOOK MCP SERVER - REAL EMAIL TEST")
        print("="*60)
        print("Testing connection to your actual Outlook instance")
        print("Target: Find email with 'On 9/20 day NPI KDP61 DVT1.0_Build EQM1 Test Result'")
        print("Goal: Extract first two lines of email body")
        print("="*60)
        
        try:
            # Step 1: Initialize server
            if not await self.initialize_server():
                return False
            
            # Step 2: Search for target email
            target_email = await self.search_target_email()
            
            if not target_email:
                print("\n❌ TARGET EMAIL NOT FOUND")
                print("Possible reasons:")
                print("   • Email might be in a different folder")
                print("   • Subject might be slightly different")
                print("   • Email might have been deleted or archived")
                print("   • Search indexing might be incomplete")
                return False
            
            # Step 3: Get email content
            email_body = await self.get_email_content(target_email["id"])
            
            if not email_body:
                print("\n❌ FAILED TO RETRIEVE EMAIL CONTENT")
                return False
            
            # Step 4: Extract first two lines
            first_line, second_line = self.extract_first_two_lines(email_body)
            
            # Step 5: Display results
            print("\n" + "="*70)
            print("🎉 TEST SUCCESSFUL!")
            print("="*70)
            print(f"✅ Found target email: {target_email.get('subject', 'Unknown')}")
            print(f"✅ Retrieved email content ({len(email_body)} characters)")
            print(f"✅ Extracted first two lines")
            
            print(f"\n📋 RESULT - FIRST TWO LINES:")
            print("-" * 40)
            print(f"Line 1: {first_line}")
            print(f"Line 2: {second_line}")
            print("-" * 40)
            
            print(f"\n📊 Email Details:")
            print(f"   Subject: {target_email.get('subject', 'Unknown')}")
            print(f"   From: {target_email.get('sender', 'Unknown')}")
            print(f"   Date: {target_email.get('received_time', 'Unknown')}")
            print(f"   Email ID: {target_email.get('id', 'Unknown')}")
            
            return True
            
        except Exception as e:
            print(f"\n❌ TEST FAILED: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            await self.cleanup_server()


async def main():
    """Main function to run the real Outlook test."""
    tester = RealOutlookTester()
    
    try:
        success = await tester.run_test()
        
        if success:
            print(f"\n🎯 MISSION ACCOMPLISHED!")
            print("The MCP server successfully connected to your Outlook,")
            print("found the target email, and extracted the first two lines.")
        else:
            print(f"\n❌ TEST INCOMPLETE")
            print("The MCP server could not complete the requested task.")
            
    except KeyboardInterrupt:
        print("\n⏹️  Test interrupted by user")
    except Exception as e:
        print(f"\n❌ Test failed with error: {e}")


if __name__ == "__main__":
    asyncio.run(main())