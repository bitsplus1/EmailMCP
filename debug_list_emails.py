#!/usr/bin/env python3
"""
Debug script for list_emails with folder ID.
"""

import json
import requests
import sys

def get_inbox_folder_id():
    """Get the inbox folder ID first."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "get-folders-debug",
        "method": "get_folders",
        "params": {}
    }
    
    try:
        print("🔍 Getting folders to find inbox ID...")
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
                
                # Find inbox folder
                for folder in folders:
                    name = folder.get("name", "")
                    if "收件" in name or "inbox" in name.lower():
                        folder_id = folder.get("id", "")
                        item_count = folder.get("item_count", 0)
                        print(f"📁 Found inbox: {name}")
                        print(f"   ID: {folder_id}")
                        print(f"   Items: {item_count}")
                        return folder_id, name, item_count
                
                print("❌ Could not find inbox folder")
                return None, None, 0
            else:
                print("❌ Unexpected response format")
                return None, None, 0
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return None, None, 0
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return None, None, 0

def test_list_emails_with_folder_id(folder_id, folder_name):
    """Test list_emails with folder ID."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "list-emails-debug",
        "method": "list_emails",
        "params": {
            "folder": folder_id,  # Using folder ID
            "limit": 3
        }
    }
    
    try:
        print(f"\n📧 Testing list_emails with folder ID...")
        print(f"   Folder: {folder_name}")
        print(f"   ID: {folder_id[:20]}...")
        
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"📋 Response received:")
            print(json.dumps(result, indent=2))
            
            if "error" in result:
                print(f"❌ Error: {result['error']['message']}")
                return False
            elif "result" in result:
                emails = result["result"].get("emails", [])
                print(f"✅ Success! Found {len(emails)} emails")
                return True
        else:
            print(f"❌ Request failed with status {response.status_code}")
            print(response.text)
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def test_list_emails_without_folder():
    """Test list_emails without folder parameter."""
    
    request = {
        "jsonrpc": "2.0",
        "id": "list-emails-no-folder",
        "method": "list_emails",
        "params": {
            "limit": 3
        }
    }
    
    try:
        print(f"\n📧 Testing list_emails WITHOUT folder...")
        
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            print(f"📋 Response received:")
            print(json.dumps(result, indent=2))
            
            if "error" in result:
                print(f"❌ Error: {result['error']['message']}")
                return False
            elif "result" in result:
                emails = result["result"].get("emails", [])
                print(f"✅ Success! Found {len(emails)} emails")
                return True
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    """Main debug function."""
    print("🚀 Debugging list_emails functionality...")
    
    # Step 1: Get inbox folder ID
    folder_id, folder_name, item_count = get_inbox_folder_id()
    
    if not folder_id:
        print("❌ Cannot proceed without folder ID")
        return
    
    if item_count == 0:
        print("⚠️ Inbox appears to be empty, but let's test anyway")
    
    # Step 2: Test list_emails without folder (baseline)
    no_folder_success = test_list_emails_without_folder()
    
    # Step 3: Test list_emails with folder ID
    folder_id_success = test_list_emails_with_folder_id(folder_id, folder_name)
    
    print(f"\n📊 Results:")
    print(f"   list_emails (no folder): {'✅' if no_folder_success else '❌'}")
    print(f"   list_emails (folder ID): {'✅' if folder_id_success else '❌'}")

if __name__ == "__main__":
    main()