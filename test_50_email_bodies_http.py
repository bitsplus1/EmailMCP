#!/usr/bin/env python3
"""
Test script to check email body content for the first 50 emails in inbox using HTTP requests.
This will help identify any issues with body content extraction.
"""

import requests
import json
import time
from typing import List, Dict, Any


class EmailBodyTester:
    """Test email body content extraction via HTTP MCP server."""
    
    def __init__(self, server_url: str = "http://192.168.1.164:8080/mcp"):
        self.server_url = server_url
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        })
    
    def make_mcp_request(self, method: str, params: Dict[str, Any], request_id: str = None) -> Dict[str, Any]:
        """Make an MCP request to the server."""
        if request_id is None:
            request_id = f"test_{int(time.time() * 1000)}"
        
        request_data = {
            "jsonrpc": "2.0",
            "id": request_id,
            "method": method,
            "params": params
        }
        
        try:
            response = self.session.post(self.server_url, json=request_data, timeout=30)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            return {
                "error": {
                    "code": -32603,
                    "message": f"HTTP request failed: {str(e)}"
                }
            }
    
    def test_server_connection(self) -> bool:
        """Test if the MCP server is responding."""
        print("🔧 Testing server connection...")
        
        # Test with a simple get_folders request
        response = self.make_mcp_request(
            "get_folders",
            {},
            "test_connection"
        )
        
        if "error" in response:
            print(f"❌ Server connection failed: {response['error']}")
            return False
        
        print("✅ Server connection successful!")
        return True
    
    def get_inbox_emails(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Get emails from inbox."""
        print(f"📧 Fetching first {limit} emails from inbox...")
        
        response = self.make_mcp_request(
            "list_inbox_emails",
            {
                "limit": limit
            },
            "get_inbox_emails"
        )
        
        if "error" in response:
            print(f"❌ Error fetching emails: {response['error']}")
            return []
        
        try:
            # Debug: Print the full response structure
            print(f"📋 Response keys: {list(response.keys())}")
            if "result" in response:
                print(f"📋 Result keys: {list(response['result'].keys())}")
                
                # Try different response formats
                if "emails" in response["result"]:
                    emails = response["result"]["emails"]
                elif "content" in response["result"]:
                    content = response["result"]["content"][0]["text"]
                    emails = json.loads(content)
                elif "data" in response["result"]:
                    emails = response["result"]["data"]
                else:
                    # Try direct result
                    emails = response["result"]
                
                print(f"✅ Successfully fetched {len(emails)} emails")
                
                # Debug: Check the type and structure of emails
                if emails:
                    print(f"📋 First email type: {type(emails[0])}")
                    if isinstance(emails[0], str):
                        print(f"📋 First email (string): {emails[0][:200]}...")
                        # Try to parse as JSON
                        try:
                            emails = [json.loads(email) if isinstance(email, str) else email for email in emails]
                            print(f"📋 Successfully parsed emails as JSON")
                        except json.JSONDecodeError:
                            print(f"📋 Emails are not JSON strings")
                    else:
                        print(f"📋 First email keys: {list(emails[0].keys()) if hasattr(emails[0], 'keys') else 'No keys'}")
                
                return emails
            else:
                print(f"❌ No 'result' in response: {response}")
                return []
        except (KeyError, json.JSONDecodeError, IndexError) as e:
            print(f"❌ Error parsing email response: {e}")
            print(f"📋 Full response: {json.dumps(response, indent=2)}")
            return []
    
    def get_email_details(self, email_id: str) -> Dict[str, Any]:
        """Get detailed email information."""
        response = self.make_mcp_request(
            "get_email",
            {
                "email_id": email_id
            },
            f"get_email_{email_id}"
        )
        
        if "error" in response:
            return {"error": response["error"]}
        
        try:
            # Try different response formats
            if "email" in response["result"]:
                return response["result"]["email"]
            elif "content" in response["result"]:
                content = response["result"]["content"][0]["text"]
                return json.loads(content)
            elif "data" in response["result"]:
                return response["result"]["data"]
            else:
                return response["result"]
        except (KeyError, json.JSONDecodeError, IndexError) as e:
            return {"error": f"Error parsing response: {e}"}
    
    def analyze_body_content(self, body: str) -> Dict[str, Any]:
        """Analyze body content and provide insights."""
        if not body or not body.strip():
            return {
                "status": "empty",
                "length": 0,
                "has_content": False,
                "preview": ""
            }
        
        body = body.strip()
        lines = body.split('\n')
        non_empty_lines = [line for line in lines if line.strip()]
        
        return {
            "status": "has_content",
            "length": len(body),
            "has_content": True,
            "total_lines": len(lines),
            "non_empty_lines": len(non_empty_lines),
            "preview": body[:200] + "..." if len(body) > 200 else body,
            "first_line": non_empty_lines[0] if non_empty_lines else "",
            "is_html": "<html" in body.lower() or "<div" in body.lower() or "<p>" in body.lower()
        }
    
    def test_all_email_bodies(self) -> Dict[str, Any]:
        """Test body content for all emails and provide detailed analysis."""
        print("\n" + "="*80)
        print("🧪 TESTING EMAIL BODY CONTENT EXTRACTION")
        print("="*80)
        
        # Test server connection first
        if not self.test_server_connection():
            return {"error": "Server connection failed"}
        
        # Get inbox emails
        emails = self.get_inbox_emails(50)
        if not emails:
            return {"error": "No emails retrieved"}
        
        # Test results
        results = {
            "total_emails": len(emails),
            "emails_with_body_in_list": 0,
            "emails_with_empty_body_in_list": 0,
            "emails_with_body_in_get": 0,
            "emails_with_empty_body_in_get": 0,
            "get_email_errors": 0,
            "detailed_results": [],
            "problematic_emails": []
        }
        
        print(f"\n📊 Testing {len(emails)} emails...")
        
        for i, email in enumerate(emails, 1):
            email_id = email.get("id", "unknown")
            subject = email.get("subject", "No Subject")[:60]
            sender = email.get("sender", "Unknown")
            
            print(f"\n--- Email {i}/{len(emails)} ---")
            print(f"ID: {email_id}")
            print(f"Subject: {subject}")
            print(f"Sender: {sender}")
            
            # Analyze body from list response
            body_from_list = email.get("body", "")
            list_analysis = self.analyze_body_content(body_from_list)
            
            if list_analysis["has_content"]:
                results["emails_with_body_in_list"] += 1
                print(f"✅ Body from list: {list_analysis['length']} chars")
                print(f"   Preview: {list_analysis['preview'][:100]}...")
            else:
                results["emails_with_empty_body_in_list"] += 1
                print("❌ Body from list: EMPTY")
            
            # Get detailed email information
            print("🔍 Getting detailed email info...")
            email_details = self.get_email_details(email_id)
            
            if "error" in email_details:
                results["get_email_errors"] += 1
                print(f"❌ Error getting email details: {email_details['error']}")
                get_analysis = {"has_content": False, "status": "error"}
            else:
                body_from_get = email_details.get("body", "")
                get_analysis = self.analyze_body_content(body_from_get)
                
                if get_analysis["has_content"]:
                    results["emails_with_body_in_get"] += 1
                    print(f"✅ Body from get_email: {get_analysis['length']} chars")
                    if get_analysis["length"] != list_analysis["length"]:
                        print(f"⚠️  Length difference: list={list_analysis['length']}, get={get_analysis['length']}")
                else:
                    results["emails_with_empty_body_in_get"] += 1
                    print("❌ Body from get_email: EMPTY")
            
            # Store detailed results
            email_result = {
                "index": i,
                "email_id": email_id,
                "subject": subject,
                "sender": sender,
                "list_body_analysis": list_analysis,
                "get_body_analysis": get_analysis,
                "has_body_issue": not list_analysis["has_content"] or not get_analysis["has_content"]
            }
            
            results["detailed_results"].append(email_result)
            
            # Track problematic emails
            if email_result["has_body_issue"]:
                results["problematic_emails"].append(email_result)
            
            # Show progress every 10 emails
            if i % 10 == 0:
                print(f"\n📈 Progress: {i}/{len(emails)} emails tested")
        
        return results
    
    def print_summary(self, results: Dict[str, Any]) -> None:
        """Print test summary."""
        if "error" in results:
            print(f"\n❌ Test failed: {results['error']}")
            return
        
        print("\n" + "="*80)
        print("📊 TEST SUMMARY")
        print("="*80)
        
        total = results["total_emails"]
        list_with_body = results["emails_with_body_in_list"]
        list_empty = results["emails_with_empty_body_in_list"]
        get_with_body = results["emails_with_body_in_get"]
        get_empty = results["emails_with_empty_body_in_get"]
        get_errors = results["get_email_errors"]
        
        print(f"Total emails tested: {total}")
        print(f"")
        print(f"LIST_INBOX_EMAILS results:")
        print(f"  ✅ Emails with body content: {list_with_body} ({list_with_body/total*100:.1f}%)")
        print(f"  ❌ Emails with empty body: {list_empty} ({list_empty/total*100:.1f}%)")
        print(f"")
        print(f"GET_EMAIL results:")
        print(f"  ✅ Emails with body content: {get_with_body} ({get_with_body/total*100:.1f}%)")
        print(f"  ❌ Emails with empty body: {get_empty} ({get_empty/total*100:.1f}%)")
        print(f"  ⚠️  Errors getting email: {get_errors} ({get_errors/total*100:.1f}%)")
        
        # Show problematic emails
        problematic = results["problematic_emails"]
        if problematic:
            print(f"\n🚨 PROBLEMATIC EMAILS ({len(problematic)} found):")
            print("-" * 80)
            
            for email in problematic[:10]:  # Show first 10
                print(f"Email {email['index']}: {email['subject']}")
                print(f"  ID: {email['email_id']}")
                print(f"  List body: {email['list_body_analysis']['status']}")
                print(f"  Get body: {email['get_body_analysis']['status']}")
                print()
            
            if len(problematic) > 10:
                print(f"... and {len(problematic) - 10} more problematic emails")
        
        # Overall assessment
        print("\n🎯 OVERALL ASSESSMENT:")
        if list_empty == 0 and get_empty == 0 and get_errors == 0:
            print("✅ ALL TESTS PASSED! All emails have body content.")
        else:
            print("❌ ISSUES FOUND! Some emails have missing body content.")
            print("\n🔧 ISSUES TO FIX:")
            if list_empty > 0:
                print(f"  • {list_empty} emails have empty body in list_inbox_emails")
            if get_empty > 0:
                print(f"  • {get_empty} emails have empty body in get_email")
            if get_errors > 0:
                print(f"  • {get_errors} emails could not be retrieved with get_email")


def main():
    """Main function to run the email body test."""
    print("🧪 EMAIL BODY CONTENT TESTER")
    print("="*50)
    print("Testing email body extraction via HTTP MCP server")
    print("Server URL: http://192.168.1.164:8080/mcp")
    print("="*50)
    
    tester = EmailBodyTester()
    
    try:
        # Run the comprehensive test
        results = tester.test_all_email_bodies()
        
        # Print summary
        tester.print_summary(results)
        
        # Save detailed results to file
        with open("email_body_test_results.json", "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        print(f"\n💾 Detailed results saved to: email_body_test_results.json")
        
    except KeyboardInterrupt:
        print("\n⏹️  Test interrupted by user")
    except Exception as e:
        print(f"\n❌ Test failed with exception: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()