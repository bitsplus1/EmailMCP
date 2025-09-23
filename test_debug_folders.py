#!/usr/bin/env python3
"""
Test script for debugging folder names and testing search functionality.
"""

import json
import requests
import sys

def test_debug_folders():
    """Test the debug_folder_names endpoint."""
    
    # Test the debug endpoint
    debug_request = {
        "jsonrpc": "2.0",
        "id": "debug-folders-test",
        "method": "debug_folder_names",
        "params": {}
    }
    
    try:
        print("🔍 Testing debug_folder_names endpoint...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=debug_request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            print("✅ Debug request successful!")
            print(json.dumps(result, indent=2, ensure_ascii=False))
            
            # Extract folder names for testing
            if "result" in result and "default_folders" in result["result"]:
                default_folders = result["result"]["default_folders"]
                
                # Find the inbox folder name
                inbox_name = None
                for folder_id, folder_info in default_folders.items():
                    if folder_id == "6" and folder_info.get("accessible"):  # Inbox is folder ID 6
                        inbox_name = folder_info.get("actual_name")
                        break
                
                if inbox_name:
                    print(f"\n📧 Found inbox folder name: '{inbox_name}'")
                    return inbox_name
                else:
                    print("\n❌ Could not find accessible inbox folder")
                    return None
            else:
                print("\n❌ Unexpected response format")
                return None
        else:
            print(f"❌ Debug request failed with status {response.status_code}")
            print(response.text)
            return None
            
    except Exception as e:
        print(f"❌ Error testing debug endpoint: {e}")
        return None

def test_search_with_folder_name(folder_name):
    """Test search_emails with the correct folder name."""
    
    search_request = {
        "jsonrpc": "2.0",
        "id": "search-test",
        "method": "search_emails",
        "params": {
            "query": "subject:Summary AND from:JackieCF_Lin@compal.com",
            "folder": folder_name,
            "limit": 5
        }
    }
    
    try:
        print(f"\n🔍 Testing search_emails with folder '{folder_name}'...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=search_request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "error" in result:
                print(f"❌ Search failed: {result['error']['message']}")
                return False
            else:
                print("✅ Search successful!")
                emails = result.get("result", {}).get("emails", [])
                print(f"📧 Found {len(emails)} emails")
                return True
        else:
            print(f"❌ Search request failed with status {response.status_code}")
            print(response.text)
            return False
            
    except Exception as e:
        print(f"❌ Error testing search: {e}")
        return False

def main():
    """Main test function."""
    print("🚀 Starting folder debug and search test...")
    
    # First, test the debug endpoint to get correct folder names
    inbox_name = test_debug_folders()
    
    if inbox_name:
        # Test search with the correct folder name
        success = test_search_with_folder_name(inbox_name)
        
        if success:
            print("\n🎉 All tests passed! Use the folder name from debug_folder_names in your n8n requests.")
        else:
            print("\n⚠️ Search test failed, but you now have the correct folder name to use.")
    else:
        print("\n❌ Could not determine correct folder name. Check server logs for more details.")

if __name__ == "__main__":
    main()