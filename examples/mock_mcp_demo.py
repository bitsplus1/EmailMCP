#!/usr/bin/env python3
"""
Mock MCP Demo - Agoda Travel Expense Analysis

This script demonstrates the Outlook MCP Server functionality using mock data
to simulate the complete workflow of searching for Agoda invoice emails and
generating a travel expense report.

This shows exactly how the MCP protocol would work in a real environment.
"""

import asyncio
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Any, List
from decimal import Decimal

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.protocol.mcp_protocol_handler import MCPProtocolHandler
from outlook_mcp_server.models.mcp_models import MCPRequest, MCPResponse
from outlook_mcp_server.logging.logger import get_logger


class MockOutlookMCPServer:
    """Mock MCP server that simulates Outlook functionality with sample data."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
        self.protocol_handler = MCPProtocolHandler()
        
        # Mock email data - simulates Agoda booking confirmation emails
        self.mock_emails = [
            {
                "id": "AAMkADExMzJmNzg4LWE4YjYtNGQ4Zi1iMzA5LTQ2ZjI3ZjE4ZjE4ZgBGAAAAAAC7xKjIH",
                "subject": "Booking Confirmation - Grand Hyatt Singapore - Booking ID: AGD123456789",
                "sender": "noreply@agoda.com",
                "recipient": "john.doe@company.com",
                "received_time": "2024-01-10T09:30:00Z",
                "is_read": True,
                "has_attachments": False,
                "folder_name": "Inbox",
                "importance": "Normal",
                "body": """
Dear John Doe,

Thank you for booking with Agoda! Your booking has been confirmed.

BOOKING DETAILS
Booking Reference: AGD123456789
Hotel: Grand Hyatt Singapore
Location: Singapore, Singapore
Address: 10 Scotts Road, Singapore 228211

CHECK-IN & CHECK-OUT
Check-in: 15 January 2024
Check-out: 18 January 2024
Nights: 3

GUEST INFORMATION
Guest Name: John Doe
Number of Guests: 1 Adult

PAYMENT SUMMARY
Room Rate: USD 420.00
Taxes & Fees: USD 30.00
Total Amount: USD 450.00

Your booking is confirmed and guaranteed. Please present this confirmation at check-in.

Best regards,
Agoda Team
                """.strip()
            },
            {
                "id": "AAMkADExMzJmNzg4LWE4YjYtNGQ4Zi1iMzA5LTQ2ZjI3ZjE4ZjE4ZgBGAAAAAAC7xKjIB",
                "subject": "Your Agoda Booking Confirmation - Park Hyatt Tokyo - AGD987654321",
                "sender": "bookings@agoda.com",
                "recipient": "john.doe@company.com", 
                "received_time": "2024-01-18T14:15:00Z",
                "is_read": True,
                "has_attachments": True,
                "folder_name": "Inbox",
                "importance": "Normal",
                "body": """
Hello John,

Your hotel booking with Agoda is confirmed!

BOOKING CONFIRMATION
Booking Reference: AGD987654321
Hotel: Park Hyatt Tokyo
Location: Tokyo, Japan
Address: 3-7-1-2 Nishi-Shinjuku, Shinjuku City, Tokyo 163-1055, Japan

STAY DETAILS
Check-in: 22 January 2024
Check-out: 25 January 2024
Duration: 3 nights

GUEST DETAILS
Guest Name: John Doe
Room Type: Deluxe King Room

FINANCIAL BREAKDOWN
Room Charges: USD 620.00
Service Charge: USD 35.00
City Tax: USD 25.00
Total Amount: USD 680.00

Thank you for choosing Agoda for your travel needs.

Warm regards,
Agoda Customer Service
                """.strip()
            },
            {
                "id": "AAMkADExMzJmNzg4LWE4YjYtNGQ4Zi1iMzA5LTQ2ZjI3ZjE4ZjE4ZgBGAAAAAAC7xKjIC",
                "subject": "Agoda Invoice - Marina Bay Sands Singapore - Ref: AGD456789123",
                "sender": "invoice@agoda.com",
                "recipient": "john.doe@company.com",
                "received_time": "2024-01-25T11:45:00Z",
                "is_read": False,
                "has_attachments": False,
                "folder_name": "Inbox",
                "importance": "High",
                "body": """
Dear Valued Customer,

This is your official invoice for your recent booking with Agoda.

INVOICE DETAILS
Booking Reference: AGD456789123
Invoice Date: 25 January 2024
Hotel: Marina Bay Sands
Location: Singapore, Singapore

RESERVATION INFORMATION
Check-in: 28 January 2024
Check-out: 30 January 2024
Nights: 2
Guest Name: John Doe

CHARGES
Accommodation: USD 280.00
Resort Fee: USD 25.00
Taxes: USD 15.00
Total Amount: USD 320.00

Payment Status: Confirmed
Payment Method: Credit Card ending in 1234

For any queries, please contact our customer service.

Best wishes,
Agoda Finance Team
                """.strip()
            },
            {
                "id": "AAMkADExMzJmNzg4LWE4YjYtNGQ4Zi1iMzA5LTQ2ZjI3ZjE4ZjE4ZgBGAAAAAAC7xKjID",
                "subject": "Booking Confirmed: The Ritz-Carlton Bangkok - AGD789123456",
                "sender": "confirmations@agoda.com",
                "recipient": "john.doe@company.com",
                "received_time": "2024-02-01T16:20:00Z",
                "is_read": True,
                "has_attachments": False,
                "folder_name": "Inbox",
                "importance": "Normal",
                "body": """
Greetings John Doe,

Your luxury hotel booking is confirmed with Agoda!

CONFIRMATION DETAILS
Booking Reference: AGD789123456
Property: The Ritz-Carlton Bangkok
Location: Bangkok, Thailand
Full Address: 181 Wireless Road, Lumpini, Pathumwan, Bangkok 10330, Thailand

TRAVEL DATES
Check-in: 05 February 2024
Check-out: 08 February 2024
Stay Duration: 3 nights

BOOKING SUMMARY
Guest Name: John Doe
Room: Executive Suite
Occupancy: 1 Adult

PAYMENT INFORMATION
Base Rate: USD 240.00
Service Charges: USD 20.00
Government Tax: USD 20.00
Total Amount: USD 280.00

Your reservation is guaranteed. Enjoy your stay!

Sincerely,
Agoda Reservations
                """.strip()
            },
            {
                "id": "AAMkADExMzJmNzg4LWE4YjYtNGQ4Zi1iMzA5LTQ2ZjI3ZjE4ZjE4ZgBGAAAAAAC7xKjIE",
                "subject": "Agoda Booking Invoice - Conrad Hong Kong - Reference AGD321654987",
                "sender": "billing@agoda.com",
                "recipient": "john.doe@company.com",
                "received_time": "2024-02-08T10:30:00Z",
                "is_read": True,
                "has_attachments": True,
                "folder_name": "Inbox",
                "importance": "Normal",
                "body": """
Dear Mr. John Doe,

Please find below your booking invoice from Agoda.

INVOICE INFORMATION
Booking Reference: AGD321654987
Invoice Number: INV-2024-0208-001
Hotel: Conrad Hong Kong
Location: Hong Kong, Hong Kong

STAY INFORMATION
Check-in: 12 February 2024
Check-out: 15 February 2024
Number of Nights: 3
Guest Name: John Doe

COST BREAKDOWN
Room Rate (3 nights): USD 480.00
Government Tax: USD 24.00
Service Fee: USD 16.00
Total Amount: USD 520.00

This invoice serves as your official receipt.

Thank you for your business,
Agoda Billing Department
                """.strip()
            }
        ]
        
        # Mock folder data
        self.mock_folders = [
            {"name": "Inbox", "item_count": 127, "folder_type": "Mail"},
            {"name": "Sent Items", "item_count": 45, "folder_type": "Mail"},
            {"name": "Drafts", "item_count": 3, "folder_type": "Mail"},
            {"name": "Deleted Items", "item_count": 12, "folder_type": "Mail"},
            {"name": "Junk Email", "item_count": 8, "folder_type": "Mail"},
            {"name": "Archive", "item_count": 234, "folder_type": "Mail"}
        ]
    
    async def handle_request(self, request_data: Dict[str, Any]) -> Dict[str, Any]:
        """Handle MCP request with mock data."""
        try:
            method = request_data.get("method")
            params = request_data.get("params", {})
            request_id = request_data.get("id")
            
            if method == "search_emails":
                return await self._handle_search_emails(request_id, params)
            elif method == "get_email":
                return await self._handle_get_email(request_id, params)
            elif method == "list_emails":
                return await self._handle_list_emails(request_id, params)
            elif method == "get_folders":
                return await self._handle_get_folders(request_id, params)
            else:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "error": {
                        "code": -32601,
                        "message": f"Method not found: {method}"
                    }
                }
                
        except Exception as e:
            return {
                "jsonrpc": "2.0",
                "id": request_data.get("id"),
                "error": {
                    "code": -32603,
                    "message": f"Internal error: {str(e)}"
                }
            }
    
    async def _handle_search_emails(self, request_id: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle search_emails request."""
        query = params.get("query", "").lower()
        limit = params.get("limit", 50)
        
        # Filter emails based on query
        matching_emails = []
        for email in self.mock_emails:
            email_text = (email["subject"] + " " + email["sender"] + " " + email["body"]).lower()
            
            # Simple search logic
            if "agoda" in query and "agoda" in email_text:
                if "invoice" in query and "invoice" in email_text:
                    matching_emails.append({
                        "id": email["id"],
                        "subject": email["subject"],
                        "sender": email["sender"],
                        "received_time": email["received_time"],
                        "is_read": email["is_read"],
                        "has_attachments": email["has_attachments"],
                        "folder_name": email["folder_name"]
                    })
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": matching_emails[:limit]
        }
    
    async def _handle_get_email(self, request_id: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle get_email request."""
        email_id = params.get("email_id")
        
        # Find email by ID
        for email in self.mock_emails:
            if email["id"] == email_id:
                return {
                    "jsonrpc": "2.0",
                    "id": request_id,
                    "result": email
                }
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "error": {
                "code": -32602,
                "message": f"Email not found: {email_id}"
            }
        }
    
    async def _handle_list_emails(self, request_id: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle list_emails request."""
        folder = params.get("folder", "Inbox")
        unread_only = params.get("unread_only", False)
        limit = params.get("limit", 50)
        
        # Filter emails
        filtered_emails = []
        for email in self.mock_emails:
            if email["folder_name"] == folder:
                if not unread_only or not email["is_read"]:
                    filtered_emails.append({
                        "id": email["id"],
                        "subject": email["subject"],
                        "sender": email["sender"],
                        "received_time": email["received_time"],
                        "is_read": email["is_read"],
                        "has_attachments": email["has_attachments"],
                        "folder_name": email["folder_name"],
                        "importance": email["importance"]
                    })
        
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": filtered_emails[:limit]
        }
    
    async def _handle_get_folders(self, request_id: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle get_folders request."""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "folders": self.mock_folders
            }
        }
    
    def get_server_info(self) -> Dict[str, Any]:
        """Get server information."""
        return {
            "name": "outlook-mcp-server",
            "version": "1.0.0",
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": [
                    {
                        "name": "search_emails",
                        "description": "Search emails by query string",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "query": {"type": "string"},
                                "limit": {"type": "integer"}
                            }
                        }
                    },
                    {
                        "name": "get_email",
                        "description": "Get detailed email by ID",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "email_id": {"type": "string"}
                            }
                        }
                    },
                    {
                        "name": "list_emails",
                        "description": "List emails with filtering options",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "folder": {"type": "string"},
                                "unread_only": {"type": "boolean"},
                                "limit": {"type": "integer"}
                            }
                        }
                    },
                    {
                        "name": "get_folders",
                        "description": "Get list of available folders",
                        "inputSchema": {
                            "type": "object",
                            "properties": {}
                        }
                    }
                ]
            }
        }


class TravelExpenseAnalyzer:
    """Analyzes Agoda emails to generate travel expense reports."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
        self.server = MockOutlookMCPServer()
        self.request_id = 1
    
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
        
        print(f"\n📤 MCP Request:")
        print(f"   Method: {method}")
        print(f"   ID: {request['id']}")
        if params:
            print(f"   Params: {json.dumps(params, indent=6)}")
        
        response = await self.server.handle_request(request)
        
        print(f"\n📥 MCP Response:")
        if "result" in response:
            result = response["result"]
            if isinstance(result, list):
                print(f"   Found {len(result)} items")
                if result and len(result) <= 3:
                    print(f"   Result: {json.dumps(result, indent=6)}")
                elif result:
                    print(f"   First item: {json.dumps(result[0], indent=6)}")
                    print(f"   ... and {len(result)-1} more items")
            else:
                print(f"   Result: {json.dumps(result, indent=6)}")
        elif "error" in response:
            print(f"   Error: {response['error']}")
        
        return response
    
    def extract_expense_from_email(self, email: Dict[str, Any]) -> Dict[str, Any]:
        """Extract expense information from email."""
        import re
        
        body = email.get("body", "")
        subject = email.get("subject", "")
        
        # Extract booking reference
        booking_ref_match = re.search(r'(?:Booking Reference|Booking ID|Reference):\s*([A-Z0-9]+)', body + subject)
        booking_ref = booking_ref_match.group(1) if booking_ref_match else "Unknown"
        
        # Extract hotel name
        hotel_match = re.search(r'Hotel:\s*(.+?)(?:\n|$)', body)
        if not hotel_match:
            hotel_match = re.search(r'(?:Property|Hotel Name):\s*(.+?)(?:\n|$)', body)
        if not hotel_match:
            # Try to extract from subject
            hotel_match = re.search(r'(?:Booking Confirmation|Invoice)\s*-\s*(.+?)\s*-', subject)
        hotel_name = hotel_match.group(1).strip() if hotel_match else "Unknown Hotel"
        
        # Extract location
        location_match = re.search(r'Location:\s*(.+?)(?:\n|$)', body)
        if not location_match:
            location_match = re.search(r'Address:\s*(.+?)(?:\n|,)', body)
        location = location_match.group(1).strip() if location_match else "Unknown Location"
        
        # Extract dates
        checkin_match = re.search(r'Check-in:\s*(\d{1,2}\s+\w+\s+\d{4})', body)
        checkout_match = re.search(r'Check-out:\s*(\d{1,2}\s+\w+\s+\d{4})', body)
        
        checkin_date = checkin_match.group(1) if checkin_match else "Unknown"
        checkout_date = checkout_match.group(1) if checkout_match else "Unknown"
        
        # Calculate nights
        nights = 0
        if checkin_match and checkout_match:
            try:
                checkin = datetime.strptime(checkin_date, '%d %B %Y')
                checkout = datetime.strptime(checkout_date, '%d %B %Y')
                nights = (checkout - checkin).days
            except:
                nights_match = re.search(r'(?:Nights?|Duration):\s*(\d+)', body)
                nights = int(nights_match.group(1)) if nights_match else 1
        
        # Extract total amount
        amount_match = re.search(r'Total Amount:\s*USD\s*([\d,]+\.?\d*)', body)
        total_amount = float(amount_match.group(1).replace(',', '')) if amount_match else 0.0
        
        # Extract guest name
        guest_match = re.search(r'Guest Name:\s*(.+?)(?:\n|$)', body)
        guest_name = guest_match.group(1).strip() if guest_match else "Unknown Guest"
        
        return {
            "booking_reference": booking_ref,
            "hotel_name": hotel_name,
            "location": location,
            "check_in_date": checkin_date,
            "check_out_date": checkout_date,
            "nights": nights,
            "total_amount": total_amount,
            "currency": "USD",
            "guest_name": guest_name,
            "email_date": email.get("received_time", ""),
            "email_id": email.get("id", "")
        }
    
    def generate_travel_report(self, expenses: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate comprehensive travel report."""
        if not expenses:
            return {
                "summary": "No expenses found",
                "total_expenses": 0,
                "total_bookings": 0,
                "destinations": []
            }
        
        total_amount = sum(exp["total_amount"] for exp in expenses)
        total_nights = sum(exp["nights"] for exp in expenses)
        destinations = list(set(exp["location"] for exp in expenses))
        
        # Group by destination
        by_destination = {}
        for exp in expenses:
            dest = exp["location"]
            if dest not in by_destination:
                by_destination[dest] = {
                    "total_amount": 0,
                    "nights": 0,
                    "bookings": 0,
                    "hotels": []
                }
            by_destination[dest]["total_amount"] += exp["total_amount"]
            by_destination[dest]["nights"] += exp["nights"]
            by_destination[dest]["bookings"] += 1
            by_destination[dest]["hotels"].append(exp["hotel_name"])
        
        # Remove duplicate hotels
        for dest in by_destination:
            by_destination[dest]["hotels"] = list(set(by_destination[dest]["hotels"]))
        
        return {
            "report_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_expenses": total_amount,
            "currency": "USD",
            "total_bookings": len(expenses),
            "total_nights": total_nights,
            "average_per_night": total_amount / total_nights if total_nights > 0 else 0,
            "destinations": destinations,
            "by_destination": by_destination,
            "expenses": expenses
        }
    
    def print_travel_report(self, report: Dict[str, Any]) -> None:
        """Print formatted travel report."""
        print("\n" + "="*80)
        print("🧳 AGODA TRAVEL EXPENSE REPORT")
        print("="*80)
        
        print(f"📅 Report Generated: {report['report_date']}")
        
        print(f"\n💰 FINANCIAL SUMMARY")
        print(f"   Total Expenses: {report['currency']} {report['total_expenses']:,.2f}")
        print(f"   Total Bookings: {report['total_bookings']}")
        print(f"   Total Nights: {report['total_nights']}")
        print(f"   Average per Night: {report['currency']} {report['average_per_night']:,.2f}")
        
        print(f"\n🌍 DESTINATIONS ({len(report['destinations'])})")
        for i, dest in enumerate(report['destinations'], 1):
            dest_info = report['by_destination'][dest]
            print(f"   {i}. {dest}")
            print(f"      💰 Amount: {report['currency']} {dest_info['total_amount']:,.2f}")
            print(f"      🌙 Nights: {dest_info['nights']}")
            print(f"      🏨 Hotels: {', '.join(dest_info['hotels'])}")
        
        print(f"\n📋 DETAILED BOOKINGS")
        for i, exp in enumerate(report['expenses'], 1):
            print(f"\n   {i}. {exp['hotel_name']}")
            print(f"      📍 {exp['location']}")
            print(f"      📅 {exp['check_in_date']} to {exp['check_out_date']} ({exp['nights']} nights)")
            print(f"      💰 {exp['currency']} {exp['total_amount']:,.2f}")
            print(f"      🎫 Ref: {exp['booking_reference']}")
        
        print("\n" + "="*80)
    
    async def run_complete_analysis(self) -> None:
        """Run complete travel expense analysis."""
        print("🚀 OUTLOOK MCP SERVER - AGODA TRAVEL EXPENSE ANALYSIS")
        print("="*70)
        print("This demo shows the complete workflow:")
        print("1. Server capability discovery")
        print("2. Email search for Agoda invoices")
        print("3. Detailed email content retrieval")
        print("4. Expense data extraction")
        print("5. Travel report generation")
        print("="*70)
        
        try:
            # Step 1: Show server capabilities
            print(f"\n📍 STEP 1: Server Capability Discovery")
            print("-" * 50)
            server_info = self.server.get_server_info()
            print(f"Server: {server_info['name']} v{server_info['version']}")
            print(f"Protocol: {server_info['protocolVersion']}")
            print(f"Available Tools: {len(server_info['capabilities']['tools'])}")
            for tool in server_info['capabilities']['tools']:
                print(f"  • {tool['name']}: {tool['description']}")
            
            # Step 2: Search for Agoda emails
            print(f"\n📍 STEP 2: Search for Agoda Invoice Emails")
            print("-" * 50)
            search_response = await self.send_mcp_request("search_emails", {
                "query": "from:Agoda invoice booking confirmation",
                "limit": 10
            })
            
            if "result" not in search_response:
                print("❌ Search failed")
                return
            
            emails = search_response["result"]
            print(f"✅ Found {len(emails)} Agoda invoice emails")
            
            # Step 3: Get detailed email content
            print(f"\n📍 STEP 3: Retrieve Detailed Email Content")
            print("-" * 50)
            
            expenses = []
            for i, email in enumerate(emails, 1):
                print(f"\n   Processing email {i}/{len(emails)}: {email['subject'][:60]}...")
                
                # Get full email content
                email_response = await self.send_mcp_request("get_email", {
                    "email_id": email["id"]
                })
                
                if "result" in email_response:
                    full_email = email_response["result"]
                    
                    # Extract expense data
                    expense = self.extract_expense_from_email(full_email)
                    expenses.append(expense)
                    
                    print(f"   ✅ Extracted: {expense['hotel_name']} - ${expense['total_amount']:.2f}")
            
            # Step 4: Generate travel report
            print(f"\n📍 STEP 4: Generate Travel Expense Report")
            print("-" * 50)
            
            report = self.generate_travel_report(expenses)
            self.print_travel_report(report)
            
            # Step 5: Save report
            print(f"\n📍 STEP 5: Save Report to File")
            print("-" * 50)
            
            report_file = "agoda_travel_report.json"
            with open(report_file, 'w') as f:
                json.dump(report, f, indent=2, default=str)
            print(f"💾 Report saved to: {report_file}")
            
            print(f"\n✅ ANALYSIS COMPLETED SUCCESSFULLY!")
            print(f"📊 Processed {len(expenses)} bookings totaling ${report['total_expenses']:,.2f}")
            print(f"🌍 Visited {len(report['destinations'])} destinations")
            print(f"🏨 Stayed {report['total_nights']} nights across {len(expenses)} hotels")
            
        except Exception as e:
            print(f"❌ Analysis failed: {e}")
            import traceback
            traceback.print_exc()


async def main():
    """Main function."""
    analyzer = TravelExpenseAnalyzer()
    
    try:
        await analyzer.run_complete_analysis()
    except KeyboardInterrupt:
        print("\n⏹️  Analysis interrupted by user")
    except Exception as e:
        print(f"\n❌ Analysis failed: {e}")


if __name__ == "__main__":
    asyncio.run(main())