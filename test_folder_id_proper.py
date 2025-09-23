#!/usr/bin/env python3
"""
Proper test script for folder ID functionality with correct parameters.
"""

import json
import requests
import sys

def test_search_emails_with_folder_id():
    """Test search_emails with folder ID parameter."""
    
    # Your inbox folder ID
    inbox_id = "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFFE000006B7F6A90000"
    
    request = {
        "jsonrpc": "2.0",
        "id": "search-with-folder-id",
        "method": "search_emails",
        "params": {
            "query": "subject:Summary AND from:JackieCF_Lin@compal.com",
            "folder": inbox_id,  # THIS IS THE CORRECT WAY - using folder ID
            "limit": 10
        }
    }
    
    try:
        print(f"🔍 Testing search_emails with folder ID: {inbox_id[:20]}...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "error" in result:
                print(f"❌ Error: {result['error']['message']}")
                return False
            elif "result" in result:
                emails = result["result"].get("emails", [])
                print(f"✅ Success! Found {len(emails)} emails using folder ID")
                if len(emails) > 0:
                    print(f"   First email: {emails[0].get('subject', 'No subject')}")
                return True
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def test_list_emails_with_folder_id():
    """Test list_emails with folder ID parameter."""
    
    # Your inbox folder ID
    inbox_id = "00000000DB2820C5F3F8204492F273035529BA6801009D535D0407B9BD4E92EA5AC6287D1BFFE000006B7F6A90000"
    
    request = {
        "jsonrpc": "2.0",
        "id": "list-with-folder-id",
        "method": "list_emails",
        "params": {
            "folder": inbox_id,  # THIS IS THE CORRECT WAY - using folder ID
            "limit": 5
        }
    }
    
    try:
        print(f"📧 Testing list_emails with folder ID: {inbox_id[:20]}...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "error" in result:
                print(f"❌ Error: {result['error']['message']}")
                return False
            elif "result" in result:
                emails = result["result"].get("emails", [])
                print(f"✅ Success! Found {len(emails)} emails using folder ID")
                if len(emails) > 0:
                    print(f"   First email: {emails[0].get('subject', 'No subject')}")
                return True
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    """Main test function."""
    print("🚀 Testing folder ID functionality with PROPER parameters...")
    
    # Test search_emails with folder ID
    search_success = test_search_emails_with_folder_id()
    
    # Test list_emails with folder ID  
    list_success = test_list_emails_with_folder_id()
    
    if search_success and list_success:
        print("\n🎉 All folder ID tests passed!")
        print("✅ You can now use folder IDs in your n8n requests")
    elif search_success:
        print("\n⚠️ search_emails with folder ID works!")
        print("❌ list_emails with folder ID has issues")
    elif list_success:
        print("\n⚠️ list_emails with folder ID works!")
        print("❌ search_emails with folder ID has issues")
    else:
        print("\n❌ Both tests failed - folder ID functionality needs more work")

if __name__ == "__main__":
    main()