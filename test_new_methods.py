#!/usr/bin/env python3
"""
Test script for the new list_inbox_emails and list_emails methods.
"""

import json
import requests

def test_list_inbox_emails():
    """Test the new list_inbox_emails method."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "test-list-inbox",
        "method": "list_inbox_emails",
        "params": {
            "limit": 3
        }
    }
    
    try:
        print("📧 Testing list_inbox_emails...")
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
                print(f"✅ Success! Found {len(emails)} emails in inbox")
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
    """Test the new list_emails method with folder_id."""
    
    # Get folder ID first
    folders_request = {
        "jsonrpc": "2.0",
        "id": "get-folders",
        "method": "get_folders",
        "params": {}
    }
    
    try:
        print("\n🔍 Getting folder ID...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=folders_request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "result" in result and "folders" in result["result"]:
                folders = result["result"]["folders"]
                
                # Find inbox folder
                inbox_folder = None
                for folder in folders:
                    if "收件" in folder.get("name", "") or "inbox" in folder.get("name", "").lower():
                        inbox_folder = folder
                        break
                
                if not inbox_folder:
                    print("❌ Could not find inbox folder")
                    return False
                
                folder_id = inbox_folder.get("id")
                folder_name = inbox_folder.get("name")
                
                print(f"📁 Found inbox: {folder_name}")
                print(f"   ID: {folder_id[:20]}...")
                
                # Now test list_emails with folder_id
                list_request = {
                    "jsonrpc": "2.0",
                    "id": "test-list-emails-folder-id",
                    "method": "list_emails",
                    "params": {
                        "folder_id": folder_id,
                        "limit": 3
                    }
                }
                
                print(f"\n📧 Testing list_emails with folder_id...")
                response = requests.post(
                    "http://127.0.0.1:8080/mcp",
                    json=list_request,
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
                        print(f"✅ Success! Found {len(emails)} emails using folder_id")
                        if len(emails) > 0:
                            print(f"   First email: {emails[0].get('subject', 'No subject')}")
                        return True
                else:
                    print(f"❌ Request failed with status {response.status_code}")
                    return False
            else:
                print("❌ Could not get folders")
                return False
        else:
            print(f"❌ Get folders failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    """Main test function."""
    print("🚀 Testing new email listing methods...")
    
    # Test list_inbox_emails
    inbox_success = test_list_inbox_emails()
    
    # Test list_emails with folder_id
    folder_id_success = test_list_emails_with_folder_id()
    
    print(f"\n📊 Results:")
    print(f"   list_inbox_emails: {'✅' if inbox_success else '❌'}")
    print(f"   list_emails (folder_id): {'✅' if folder_id_success else '❌'}")
    
    if inbox_success and folder_id_success:
        print("\n🎉 Both methods are working! Ready for n8n integration.")
    elif inbox_success:
        print("\n⚠️ list_inbox_emails works, but list_emails with folder_id needs work.")
    else:
        print("\n❌ Both methods need debugging.")

if __name__ == "__main__":
    main()