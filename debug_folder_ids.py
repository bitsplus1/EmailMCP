#!/usr/bin/env python3
"""
Debug script to compare folder IDs from get_folders vs actual Outlook folder IDs.
"""

import json
import requests

def debug_folder_ids():
    """Debug folder ID comparison."""
    
    # Get folders from get_folders
    request = {
        "jsonrpc": "2.0",
        "id": "debug-folder-ids",
        "method": "debug_folder_names",
        "params": {}
    }
    
    try:
        print("🔍 Getting folder debug info...")
        response = requests.post(
            "http://127.0.0.1:8080/mcp",
            json=request,
            headers={"Content-Type": "application/json"},
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            if "result" in result:
                debug_info = result["result"]
                
                print("📁 Default folders from debug:")
                default_folders = debug_info.get("default_folders", {})
                for folder_id, folder_info in default_folders.items():
                    print(f"   Folder {folder_id}: {folder_info}")
                
                print("\n📁 All folders from debug:")
                all_folders = debug_info.get("all_folders", [])
                for folder in all_folders:
                    print(f"   {folder.get('name')}: {folder.get('full_path')}")
                
                return True
            else:
                print("❌ No result in response")
                return False
        else:
            print(f"❌ Request failed with status {response.status_code}")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def main():
    """Main debug function."""
    print("🚀 Debugging folder ID comparison...")
    debug_folder_ids()

if __name__ == "__main__":
    main()