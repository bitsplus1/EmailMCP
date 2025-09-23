#!/usr/bin/env python3
"""
Test script for folder ID functionality.
"""

import json
import requests
import sys

def test_get_folders():
    """Get folders and their IDs."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "get-folders-test",
        "method": "get_folders",
        "params": {}
    }
    
    try:
        print("🔍 Getting all folders with IDs...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "result" in result and "folders" in result["result"]:
                folders = result["result"]["folders"]
                print(f"✅ Found {len(folders)} folders:")
                
                inbox_folder = None
                for folder in folders:
                    name = folder.get("name", "Unknown")
                    folder_id = folder.get("id", "No ID")
                    item_count = folder.get("item_count", 0)
                    print(f"  📁 {name}")
                    print(f"     ID: {folder_id}")
                    print(f"     Items: {item_count}")
                    print()
                    
                    # Look for inbox-like folder
                    if "收件" in name or "inbox" in name.lower():
                        inbox_folder = folder
                
                return inbox_folder
            else:
                print("❌ Unexpected response format")
                return None
        else:
            print(f"❌ Request failed with status {response.status_code}")
            print(response.text)
            return None
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def test_list_emails_by_id(folder_id):
    """Test listing emails using folder ID."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "list-by-id-test",
        "method": "list_emails",
        "params": {
            "folder": folder_id,
            "limit": 5
        }
    }
    
    try:
        print(f"📧 Testing list_emails with folder ID: {folder_id[:20]}...")
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

def test_search_emails_by_id(folder_id):
    """Test searching emails using folder ID."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "search-by-id-test",
        "method": "search_emails",
        "params": {
            "query": "received:thisweek",
            "folder": folder_id,
            "limit": 5
        }
    }
    
    try:
        print(f"🔍 Testing search_emails with folder ID: {folder_id[:20]}...")
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
                print(f"✅ Success! Found {len(emails)} emails using folder ID search")
                return True
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    """Main test function."""
    print("🚀 Testing folder ID functionality...")
    
    # Get folders and find inbox
    inbox_folder = test_get_folders()
    
    if not inbox_folder:
        print("❌ Could not find inbox folder")
        return
    
    folder_id = inbox_folder.get("id")
    folder_name = inbox_folder.get("name")
    
    if not folder_id:
        print("❌ Inbox folder has no ID")
        return
    
    print(f"📁 Using folder: {folder_name} (ID: {folder_id[:20]}...)")
    
    # Test list_emails with folder ID
    list_success = test_list_emails_by_id(folder_id)
    
    # Test search_emails with folder ID
    search_success = test_search_emails_by_id(folder_id)
    
    if list_success and search_success:
        print("\n🎉 All folder ID tests passed!")
        print(f"✅ You can now use folder ID '{folder_id}' in your n8n requests")
        print("✅ This avoids all encoding issues with Chinese folder names")
    else:
        print("\n⚠️ Some tests failed, but folder ID functionality may still work")

if __name__ == "__main__":
    main()