#!/usr/bin/env python3
"""
Travel Expense Analyzer - Agoda Invoice Processing Demo

This script demonstrates how to use the Outlook MCP Server to:
1. Search for Agoda invoice emails
2. Extract expense and travel information
3. Generate a comprehensive travel expense report

This simulates real MCP client interactions with the server.
"""

import asyncio
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, asdict
from decimal import Decimal

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from outlook_mcp_server.server import OutlookMCPServer, create_server_config
from outlook_mcp_server.models.mcp_models import MCPRequest, MCPResponse
from outlook_mcp_server.logging.logger import get_logger


@dataclass
class TravelExpense:
    """Travel expense data structure."""
    booking_reference: str
    hotel_name: str
    location: str
    check_in_date: str
    check_out_date: str
    nights: int
    total_amount: Decimal
    currency: str
    guest_name: str
    email_date: str
    email_id: str


@dataclass
class TravelReport:
    """Travel expense report data structure."""
    report_generated_at: str
    total_expenses: Decimal
    currency: str
    total_bookings: int
    total_nights: int
    destinations: List[str]
    date_range: Dict[str, str]
    expenses: List[TravelExpense]
    summary_by_destination: Dict[str, Dict[str, Any]]


class TravelExpenseAnalyzer:
    """Analyzes Agoda invoice emails to generate travel expense reports."""
    
    def __init__(self):
        self.logger = get_logger(__name__)
        self.server: Optional[OutlookMCPServer] = None
        
        # Patterns for extracting information from email content
        self.patterns = {
            'booking_reference': r'Booking\s+(?:Reference|ID|Number):\s*([A-Z0-9]+)',
            'hotel_name': r'Hotel:\s*(.+?)(?:\n|$)',
            'location': r'(?:Location|Address):\s*(.+?)(?:\n|,)',
            'check_in': r'Check-in:\s*(\d{1,2}\s+\w+\s+\d{4})',
            'check_out': r'Check-out:\s*(\d{1,2}\s+\w+\s+\d{4})',
            'total_amount': r'Total\s+Amount:\s*([A-Z]{3})\s*([\d,]+\.?\d*)',
            'guest_name': r'Guest\s+Name:\s*(.+?)(?:\n|$)',
        }
    
    async def initialize_server(self) -> None:
        """Initialize the MCP server for testing."""
        try:
            self.logger.info("Initializing Outlook MCP Server for demo")
            
            # Create server configuration
            config = create_server_config(
                log_level="INFO",
                enable_console_output=True,
                max_concurrent_requests=5
            )
            
            # Create and start server
            self.server = OutlookMCPServer(config)
            await self.server.start()
            
            self.logger.info("Server initialized successfully")
            
        except Exception as e:
            self.logger.error(f"Failed to initialize server: {e}")
            raise
    
    async def cleanup_server(self) -> None:
        """Cleanup server resources."""
        if self.server:
            await self.server.stop()
            self.server = None
    
    async def search_agoda_invoices(self) -> List[Dict[str, Any]]:
        """
        Search for Agoda invoice emails using MCP requests.
        
        Returns:
            List of email dictionaries containing Agoda invoices
        """
        self.logger.info("Searching for Agoda invoice emails")
        
        try:
            # Create MCP request to search for Agoda emails
            search_request = MCPRequest(
                id="search_agoda_1",
                method="search_emails",
                params={
                    "query": "from:Agoda invoice booking confirmation",
                    "limit": 50
                }
            )
            
            # Simulate MCP request processing
            self.logger.info(f"Sending MCP request: {search_request.method}")
            self.logger.debug(f"Request params: {search_request.params}")
            
            # Process request through server
            request_data = {
                "jsonrpc": "2.0",
                "id": search_request.id,
                "method": search_request.method,
                "params": search_request.params
            }
            
            response_data = await self.server.handle_request(request_data)
            
            # Extract results
            if "result" in response_data:
                emails = response_data["result"]
                self.logger.info(f"Found {len(emails)} Agoda emails")
                return emails
            else:
                self.logger.warning(f"Search request failed: {response_data.get('error', 'Unknown error')}")
                return []
                
        except Exception as e:
            self.logger.error(f"Error searching for Agoda emails: {e}")
            return []
    
    async def get_email_details(self, email_id: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed email content using MCP request.
        
        Args:
            email_id: Email ID to retrieve
            
        Returns:
            Email details dictionary or None if failed
        """
        try:
            # Create MCP request to get email details
            get_request = MCPRequest(
                id=f"get_email_{email_id[:8]}",
                method="get_email",
                params={"email_id": email_id}
            )
            
            # Process request
            request_data = {
                "jsonrpc": "2.0",
                "id": get_request.id,
                "method": get_request.method,
                "params": get_request.params
            }
            
            response_data = await self.server.handle_request(request_data)
            
            if "result" in response_data:
                return response_data["result"]
            else:
                self.logger.warning(f"Failed to get email {email_id}: {response_data.get('error')}")
                return None
                
        except Exception as e:
            self.logger.error(f"Error getting email details for {email_id}: {e}")
            return None
    
    def extract_expense_data(self, email: Dict[str, Any]) -> Optional[TravelExpense]:
        """
        Extract travel expense data from email content.
        
        Args:
            email: Email dictionary with content
            
        Returns:
            TravelExpense object or None if extraction failed
        """
        try:
            content = email.get("body", "") + " " + email.get("subject", "")
            
            # Extract information using regex patterns
            extracted = {}
            for key, pattern in self.patterns.items():
                match = re.search(pattern, content, re.IGNORECASE | re.MULTILINE)
                if match:
                    if key == 'total_amount':
                        extracted['currency'] = match.group(1)
                        extracted['total_amount'] = Decimal(match.group(2).replace(',', ''))
                    else:
                        extracted[key] = match.group(1).strip()
            
            # Calculate nights if we have check-in and check-out dates
            nights = 0
            if 'check_in' in extracted and 'check_out' in extracted:
                try:
                    check_in = datetime.strptime(extracted['check_in'], '%d %B %Y')
                    check_out = datetime.strptime(extracted['check_out'], '%d %B %Y')
                    nights = (check_out - check_in).days
                except ValueError:
                    nights = 1  # Default to 1 night if parsing fails
            
            # Create TravelExpense object
            expense = TravelExpense(
                booking_reference=extracted.get('booking_reference', 'Unknown'),
                hotel_name=extracted.get('hotel_name', 'Unknown Hotel'),
                location=extracted.get('location', 'Unknown Location'),
                check_in_date=extracted.get('check_in', 'Unknown'),
                check_out_date=extracted.get('check_out', 'Unknown'),
                nights=nights,
                total_amount=extracted.get('total_amount', Decimal('0')),
                currency=extracted.get('currency', 'USD'),
                guest_name=extracted.get('guest_name', 'Unknown Guest'),
                email_date=email.get('received_time', ''),
                email_id=email.get('id', '')
            )
            
            return expense
            
        except Exception as e:
            self.logger.error(f"Error extracting expense data: {e}")
            return None
    
    def generate_mock_expense_data(self) -> List[TravelExpense]:
        """
        Generate mock expense data for demonstration purposes.
        This simulates what would be extracted from real Agoda emails.
        """
        mock_expenses = [
            TravelExpense(
                booking_reference="AGD123456789",
                hotel_name="Grand Hyatt Singapore",
                location="Singapore, Singapore",
                check_in_date="15 January 2024",
                check_out_date="18 January 2024",
                nights=3,
                total_amount=Decimal("450.00"),
                currency="USD",
                guest_name="John Doe",
                email_date="2024-01-10T09:30:00Z",
                email_id="mock_email_1"
            ),
            TravelExpense(
                booking_reference="AGD987654321",
                hotel_name="Park Hyatt Tokyo",
                location="Tokyo, Japan",
                check_in_date="22 January 2024",
                check_out_date="25 January 2024",
                nights=3,
                total_amount=Decimal("680.00"),
                currency="USD",
                guest_name="John Doe",
                email_date="2024-01-18T14:15:00Z",
                email_id="mock_email_2"
            ),
            TravelExpense(
                booking_reference="AGD456789123",
                hotel_name="Marina Bay Sands",
                location="Singapore, Singapore",
                check_in_date="28 January 2024",
                check_out_date="30 January 2024",
                nights=2,
                total_amount=Decimal("320.00"),
                currency="USD",
                guest_name="John Doe",
                email_date="2024-01-25T11:45:00Z",
                email_id="mock_email_3"
            ),
            TravelExpense(
                booking_reference="AGD789123456",
                hotel_name="The Ritz-Carlton Bangkok",
                location="Bangkok, Thailand",
                check_in_date="05 February 2024",
                check_out_date="08 February 2024",
                nights=3,
                total_amount=Decimal("280.00"),
                currency="USD",
                guest_name="John Doe",
                email_date="2024-02-01T16:20:00Z",
                email_id="mock_email_4"
            ),
            TravelExpense(
                booking_reference="AGD321654987",
                hotel_name="Conrad Hong Kong",
                location="Hong Kong, Hong Kong",
                check_in_date="12 February 2024",
                check_out_date="15 February 2024",
                nights=3,
                total_amount=Decimal("520.00"),
                currency="USD",
                guest_name="John Doe",
                email_date="2024-02-08T10:30:00Z",
                email_id="mock_email_5"
            )
        ]
        
        return mock_expenses
    
    def generate_travel_report(self, expenses: List[TravelExpense]) -> TravelReport:
        """
        Generate a comprehensive travel expense report.
        
        Args:
            expenses: List of travel expenses
            
        Returns:
            TravelReport object
        """
        if not expenses:
            return TravelReport(
                report_generated_at=datetime.now().isoformat(),
                total_expenses=Decimal('0'),
                currency="USD",
                total_bookings=0,
                total_nights=0,
                destinations=[],
                date_range={"start": "", "end": ""},
                expenses=[],
                summary_by_destination={}
            )
        
        # Calculate totals
        total_expenses = sum(expense.total_amount for expense in expenses)
        total_bookings = len(expenses)
        total_nights = sum(expense.nights for expense in expenses)
        
        # Get unique destinations
        destinations = list(set(expense.location for expense in expenses))
        
        # Calculate date range
        dates = [datetime.fromisoformat(expense.email_date.replace('Z', '+00:00')) 
                for expense in expenses if expense.email_date]
        date_range = {
            "start": min(dates).strftime('%Y-%m-%d') if dates else "",
            "end": max(dates).strftime('%Y-%m-%d') if dates else ""
        }
        
        # Summary by destination
        summary_by_destination = {}
        for expense in expenses:
            location = expense.location
            if location not in summary_by_destination:
                summary_by_destination[location] = {
                    "total_amount": Decimal('0'),
                    "total_nights": 0,
                    "bookings": 0,
                    "hotels": []
                }
            
            summary_by_destination[location]["total_amount"] += expense.total_amount
            summary_by_destination[location]["total_nights"] += expense.nights
            summary_by_destination[location]["bookings"] += 1
            summary_by_destination[location]["hotels"].append(expense.hotel_name)
        
        # Remove duplicate hotels
        for location in summary_by_destination:
            summary_by_destination[location]["hotels"] = list(set(
                summary_by_destination[location]["hotels"]
            ))
        
        return TravelReport(
            report_generated_at=datetime.now().isoformat(),
            total_expenses=total_expenses,
            currency=expenses[0].currency if expenses else "USD",
            total_bookings=total_bookings,
            total_nights=total_nights,
            destinations=destinations,
            date_range=date_range,
            expenses=expenses,
            summary_by_destination=summary_by_destination
        )
    
    def print_travel_report(self, report: TravelReport) -> None:
        """Print a formatted travel expense report."""
        print("\n" + "="*80)
        print("🧳 TRAVEL EXPENSE REPORT - AGODA BOOKINGS")
        print("="*80)
        
        print(f"📅 Report Generated: {report.report_generated_at}")
        print(f"📊 Analysis Period: {report.date_range['start']} to {report.date_range['end']}")
        
        print(f"\n💰 FINANCIAL SUMMARY")
        print(f"   Total Expenses: {report.currency} {report.total_expenses:,.2f}")
        print(f"   Total Bookings: {report.total_bookings}")
        print(f"   Total Nights: {report.total_nights}")
        print(f"   Average per Night: {report.currency} {(report.total_expenses / report.total_nights):,.2f}")
        
        print(f"\n🌍 DESTINATIONS VISITED")
        for i, destination in enumerate(report.destinations, 1):
            print(f"   {i}. {destination}")
        
        print(f"\n📍 BREAKDOWN BY DESTINATION")
        for location, summary in report.summary_by_destination.items():
            print(f"\n   🏨 {location}")
            print(f"      Amount: {report.currency} {summary['total_amount']:,.2f}")
            print(f"      Nights: {summary['total_nights']}")
            print(f"      Bookings: {summary['bookings']}")
            print(f"      Hotels: {', '.join(summary['hotels'])}")
        
        print(f"\n📋 DETAILED BOOKING HISTORY")
        for i, expense in enumerate(report.expenses, 1):
            print(f"\n   {i}. {expense.hotel_name}")
            print(f"      📍 Location: {expense.location}")
            print(f"      📅 Dates: {expense.check_in_date} to {expense.check_out_date}")
            print(f"      🌙 Nights: {expense.nights}")
            print(f"      💰 Amount: {expense.currency} {expense.total_amount:,.2f}")
            print(f"      🎫 Booking Ref: {expense.booking_reference}")
            print(f"      👤 Guest: {expense.guest_name}")
        
        print("\n" + "="*80)
        print("📈 TRAVEL INSIGHTS")
        print("="*80)
        
        # Calculate insights
        avg_booking_amount = report.total_expenses / report.total_bookings
        most_expensive = max(report.expenses, key=lambda x: x.total_amount)
        longest_stay = max(report.expenses, key=lambda x: x.nights)
        
        print(f"💡 Average booking amount: {report.currency} {avg_booking_amount:,.2f}")
        print(f"💎 Most expensive stay: {most_expensive.hotel_name} ({report.currency} {most_expensive.total_amount:,.2f})")
        print(f"⏰ Longest stay: {longest_stay.hotel_name} ({longest_stay.nights} nights)")
        print(f"🏆 Most visited destination: {max(report.summary_by_destination.keys(), key=lambda x: report.summary_by_destination[x]['bookings'])}")
        
        print("\n" + "="*80)
    
    async def run_analysis(self) -> TravelReport:
        """
        Run the complete travel expense analysis workflow.
        
        Returns:
            TravelReport with analysis results
        """
        try:
            print("🚀 Starting Travel Expense Analysis")
            print("="*50)
            
            # Initialize server
            print("📡 Initializing Outlook MCP Server...")
            await self.initialize_server()
            
            # Search for Agoda emails
            print("🔍 Searching for Agoda invoice emails...")
            emails = await self.search_agoda_invoices()
            
            expenses = []
            
            if emails:
                print(f"📧 Processing {len(emails)} emails...")
                
                # Process each email to extract expense data
                for i, email in enumerate(emails, 1):
                    print(f"   Processing email {i}/{len(emails)}: {email.get('subject', 'No Subject')[:50]}...")
                    
                    # Get detailed email content
                    email_details = await self.get_email_details(email['id'])
                    if email_details:
                        # Extract expense data
                        expense = self.extract_expense_data(email_details)
                        if expense:
                            expenses.append(expense)
            else:
                print("⚠️  No Agoda emails found. Using mock data for demonstration...")
                expenses = self.generate_mock_expense_data()
            
            print(f"✅ Successfully processed {len(expenses)} travel expenses")
            
            # Generate report
            print("📊 Generating travel expense report...")
            report = self.generate_travel_report(expenses)
            
            return report
            
        except Exception as e:
            self.logger.error(f"Analysis failed: {e}", exc_info=True)
            raise
        finally:
            # Cleanup
            await self.cleanup_server()
    
    def save_report_json(self, report: TravelReport, filename: str = "travel_expense_report.json") -> None:
        """Save the report as JSON file."""
        try:
            # Convert Decimal objects to float for JSON serialization
            report_dict = asdict(report)
            
            def decimal_to_float(obj):
                if isinstance(obj, dict):
                    return {k: decimal_to_float(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [decimal_to_float(item) for item in obj]
                elif isinstance(obj, Decimal):
                    return float(obj)
                else:
                    return obj
            
            report_dict = decimal_to_float(report_dict)
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(report_dict, f, indent=2, default=str)
            
            print(f"💾 Report saved to: {filename}")
            
        except Exception as e:
            self.logger.error(f"Failed to save report: {e}")


async def main():
    """Main function to run the travel expense analysis demo."""
    analyzer = TravelExpenseAnalyzer()
    
    try:
        print("🧳 Outlook MCP Server - Travel Expense Analyzer Demo")
        print("="*60)
        print("This demo simulates MCP requests to:")
        print("1. Search for Agoda invoice emails")
        print("2. Extract travel expense information")
        print("3. Generate a comprehensive travel report")
        print("="*60)
        
        # Run the analysis
        report = await analyzer.run_analysis()
        
        # Display the report
        analyzer.print_travel_report(report)
        
        # Save report to file
        analyzer.save_report_json(report)
        
        print("\n✅ Travel expense analysis completed successfully!")
        print("\n💡 This demo shows how the Outlook MCP Server can be used to:")
        print("   • Search emails with complex queries")
        print("   • Extract structured data from email content")
        print("   • Process multiple emails efficiently")
        print("   • Generate business intelligence reports")
        
    except KeyboardInterrupt:
        print("\n⏹️  Analysis interrupted by user")
    except Exception as e:
        print(f"\n❌ Analysis failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())