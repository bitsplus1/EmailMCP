#!/usr/bin/env python3
"""
MCP Client Simulation for Outlook MCP Server

This script simulates a real MCP client interacting with the Outlook MCP Server
to demonstrate the protocol in action. It shows the exact JSON-RPC messages
that would be exchanged between a client and server.
"""

import asyncio
import json
import sys
from pathlib import Path
from typing import Dict, Any, List

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.logging.logger import get_logger


class MCPClientSimulator:
    """Simulates an MCP client interacting with the Outlook MCP Server."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
        self.server = None
        self.request_id = 1
    
    async def initialize_server(self) -> None:
        """Initialize the MCP server."""
        try:
            print("🔧 Initializing Outlook MCP Server...")
            
            config = create_server_config(
                log_level="INFO",
                enable_console_output=True,
                max_concurrent_requests=10
            )
            
            self.server = OutlookMCPServer(config)
            await self.server.start()
            
            print("✅ Server initialized and ready for MCP requests")
            
        except Exception as e:
            print(f"❌ Failed to initialize server: {e}")
            raise
    
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
        """
        Send an MCP request to the server and return the response.
        
        Args:
            method: MCP method name
            params: Request parameters
            
        Returns:
            MCP response dictionary
        """
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
        
        print(f"\n📋 Full JSON-RPC Request:")
        print(json.dumps(request, indent=2))
        
        try:
            # Send request to server
            response = await self.server.handle_request(request)
            
            print(f"\n📥 Received MCP Response:")
            print(json.dumps(response, indent=2))
            
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
            
            print(f"\n❌ Error Response:")
            print(json.dumps(error_response, indent=2))
            
            return error_response
    
    async def demonstrate_server_capabilities(self) -> None:
        """Demonstrate server capabilities discovery."""
        print("\n" + "="*60)
        print("🔍 DISCOVERING SERVER CAPABILITIES")
        print("="*60)
        
        # Get server info (this would typically be done during initialization)
        if self.server:
            server_info = self.server.get_server_info()
            
            print("📊 Server Information:")
            print(f"   Name: {server_info.get('name', 'Unknown')}")
            print(f"   Version: {server_info.get('version', 'Unknown')}")
            print(f"   Protocol Version: {server_info.get('protocolVersion', 'Unknown')}")
            
            capabilities = server_info.get('capabilities', {})
            tools = capabilities.get('tools', [])
            
            print(f"\n🛠️  Available Tools ({len(tools)}):")
            for i, tool in enumerate(tools, 1):
                print(f"   {i}. {tool.get('name', 'Unknown')}")
                print(f"      Description: {tool.get('description', 'No description')}")
                
                # Show input schema if available
                input_schema = tool.get('inputSchema', {})
                if input_schema:
                    properties = input_schema.get('properties', {})
                    if properties:
                        print(f"      Parameters: {', '.join(properties.keys())}")
    
    async def demonstrate_email_search(self) -> List[Dict[str, Any]]:
        """Demonstrate email search functionality."""
        print("\n" + "="*60)
        print("🔍 SEARCHING FOR AGODA INVOICE EMAILS")
        print("="*60)
        
        # Search for Agoda emails
        search_params = {
            "query": "from:Agoda invoice booking confirmation",
            "limit": 10
        }
        
        response = await self.send_mcp_request("search_emails", search_params)
        
        emails = []
        if "result" in response:
            emails = response["result"]
            print(f"\n✅ Found {len(emails)} emails matching search criteria")
            
            if emails:
                print("\n📧 Email Summary:")
                for i, email in enumerate(emails[:5], 1):  # Show first 5
                    subject = email.get('subject', 'No Subject')[:50]
                    sender = email.get('sender', 'Unknown Sender')
                    date = email.get('received_time', 'Unknown Date')
                    print(f"   {i}. {subject}...")
                    print(f"      From: {sender}")
                    print(f"      Date: {date}")
                
                if len(emails) > 5:
                    print(f"   ... and {len(emails) - 5} more emails")
        else:
            error = response.get('error', {})
            print(f"❌ Search failed: {error.get('message', 'Unknown error')}")
        
        return emails
    
    async def demonstrate_email_retrieval(self, email_id: str) -> Dict[str, Any]:
        """Demonstrate detailed email retrieval."""
        print(f"\n" + "="*60)
        print(f"📧 RETRIEVING EMAIL DETAILS")
        print("="*60)
        
        get_params = {"email_id": email_id}
        response = await self.send_mcp_request("get_email", get_params)
        
        email_details = {}
        if "result" in response:
            email_details = response["result"]
            print(f"\n✅ Successfully retrieved email details")
            
            # Display key information
            print(f"\n📋 Email Information:")
            print(f"   Subject: {email_details.get('subject', 'No Subject')}")
            print(f"   From: {email_details.get('sender', 'Unknown Sender')}")
            print(f"   To: {email_details.get('recipient', 'Unknown Recipient')}")
            print(f"   Date: {email_details.get('received_time', 'Unknown Date')}")
            print(f"   Has Attachments: {email_details.get('has_attachments', False)}")
            
            # Show body preview
            body = email_details.get('body', '')
            if body:
                preview = body[:200] + "..." if len(body) > 200 else body
                print(f"   Body Preview: {preview}")
        else:
            error = response.get('error', {})
            print(f"❌ Email retrieval failed: {error.get('message', 'Unknown error')}")
        
        return email_details
    
    async def demonstrate_folder_listing(self) -> List[Dict[str, Any]]:
        """Demonstrate folder listing functionality."""
        print(f"\n" + "="*60)
        print(f"📁 LISTING OUTLOOK FOLDERS")
        print("="*60)
        
        response = await self.send_mcp_request("get_folders")
        
        folders = []
        if "result" in response:
            result = response["result"]
            folders = result.get("folders", [])
            print(f"\n✅ Found {len(folders)} folders")
            
            if folders:
                print(f"\n📂 Available Folders:")
                for i, folder in enumerate(folders, 1):
                    name = folder.get('name', 'Unknown Folder')
                    item_count = folder.get('item_count', 0)
                    folder_type = folder.get('folder_type', 'Unknown')
                    print(f"   {i}. {name} ({item_count} items, type: {folder_type})")
        else:
            error = response.get('error', {})
            print(f"❌ Folder listing failed: {error.get('message', 'Unknown error')}")
        
        return folders
    
    async def demonstrate_advanced_search(self) -> List[Dict[str, Any]]:
        """Demonstrate advanced search with multiple criteria."""
        print(f"\n" + "="*60)
        print(f"🔍 ADVANCED EMAIL SEARCH")
        print("="*60)
        
        # Search for recent unread emails in Inbox
        search_params = {
            "folder": "Inbox",
            "unread_only": True,
            "limit": 5
        }
        
        response = await self.send_mcp_request("list_emails", search_params)
        
        emails = []
        if "result" in response:
            emails = response["result"]
            print(f"\n✅ Found {len(emails)} unread emails in Inbox")
            
            if emails:
                print(f"\n📧 Unread Emails:")
                for i, email in enumerate(emails, 1):
                    subject = email.get('subject', 'No Subject')[:40]
                    sender = email.get('sender', 'Unknown Sender')
                    importance = email.get('importance', 'Normal')
                    print(f"   {i}. {subject}...")
                    print(f"      From: {sender} (Importance: {importance})")
        else:
            error = response.get('error', {})
            print(f"❌ Advanced search failed: {error.get('message', 'Unknown error')}")
        
        return emails
    
    async def simulate_travel_expense_workflow(self) -> None:
        """Simulate the complete travel expense analysis workflow."""
        print(f"\n" + "="*80)
        print(f"🧳 COMPLETE TRAVEL EXPENSE ANALYSIS WORKFLOW")
        print("="*80)
        
        try:
            # Step 1: Search for Agoda emails
            print(f"\n📍 Step 1: Searching for Agoda invoice emails...")
            agoda_emails = await self.demonstrate_email_search()
            
            if not agoda_emails:
                print("⚠️  No Agoda emails found. This would typically mean:")
                print("   • No Agoda bookings in the email account")
                print("   • Emails might be in a different folder")
                print("   • Search criteria might need adjustment")
                return
            
            # Step 2: Get detailed information for first few emails
            print(f"\n📍 Step 2: Retrieving detailed email content...")
            detailed_emails = []
            
            for i, email in enumerate(agoda_emails[:3], 1):  # Process first 3 emails
                print(f"\n   Processing email {i}/3...")
                email_details = await self.demonstrate_email_retrieval(email['id'])
                if email_details:
                    detailed_emails.append(email_details)
            
            # Step 3: Simulate expense extraction
            print(f"\n📍 Step 3: Extracting expense information...")
            total_expenses = 0
            destinations = set()
            
            for email in detailed_emails:
                # Simulate expense extraction (in real scenario, this would parse email content)
                print(f"   📧 Processing: {email.get('subject', 'No Subject')[:50]}...")
                
                # Mock expense data extraction
                mock_amount = 250.00 + (len(email.get('subject', '')) % 300)  # Mock calculation
                mock_destination = ["Singapore", "Tokyo", "Bangkok", "Hong Kong"][len(detailed_emails) % 4]
                
                total_expenses += mock_amount
                destinations.add(mock_destination)
                
                print(f"      💰 Extracted Amount: USD {mock_amount:.2f}")
                print(f"      📍 Destination: {mock_destination}")
            
            # Step 4: Generate summary report
            print(f"\n📍 Step 4: Generating travel expense summary...")
            print(f"\n📊 TRAVEL EXPENSE SUMMARY")
            print(f"   Total Bookings Processed: {len(detailed_emails)}")
            print(f"   Total Expenses: USD {total_expenses:.2f}")
            print(f"   Average per Booking: USD {total_expenses/len(detailed_emails):.2f}")
            print(f"   Destinations: {', '.join(destinations)}")
            
            print(f"\n✅ Travel expense analysis workflow completed successfully!")
            
        except Exception as e:
            print(f"❌ Workflow failed: {e}")
    
    async def run_complete_demo(self) -> None:
        """Run the complete MCP client simulation demo."""
        try:
            print("🚀 Outlook MCP Server - Client Simulation Demo")
            print("="*60)
            print("This demo simulates a real MCP client interacting with the server")
            print("to demonstrate the complete protocol and workflow.")
            print("="*60)
            
            # Initialize server
            await self.initialize_server()
            
            # Demonstrate server capabilities
            await self.demonstrate_server_capabilities()
            
            # Demonstrate folder listing
            await self.demonstrate_folder_listing()
            
            # Demonstrate advanced search
            await self.demonstrate_advanced_search()
            
            # Simulate complete travel expense workflow
            await self.simulate_travel_expense_workflow()
            
            print(f"\n" + "="*80)
            print("✅ MCP CLIENT SIMULATION COMPLETED SUCCESSFULLY!")
            print("="*80)
            print("This demo showed:")
            print("• Server capability discovery")
            print("• Folder listing operations")
            print("• Email search with various parameters")
            print("• Detailed email retrieval")
            print("• Complete business workflow simulation")
            print("• Proper JSON-RPC message formatting")
            print("• Error handling and response processing")
            
        except Exception as e:
            print(f"\n❌ Demo failed: {e}")
            import traceback
            traceback.print_exc()
        finally:
            await self.cleanup_server()


async def main():
    """Main function to run the MCP client simulation."""
    simulator = MCPClientSimulator()
    
    try:
        await simulator.run_complete_demo()
    except KeyboardInterrupt:
        print("\n⏹️  Demo interrupted by user")
    except Exception as e:
        print(f"\n❌ Demo failed: {e}")


if __name__ == "__main__":
    asyncio.run(main())