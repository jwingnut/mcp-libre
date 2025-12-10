#!/usr/bin/env python3
"""Test script for LibreOffice MCP Extension on Windows"""
import requests
import time
import sys

def test_mcp_server():
    """Test if MCP server is accessible"""
    server_url = "http://localhost:8765"

    print("🧪 Testing LibreOffice MCP Extension...")
    print(f"📡 Checking server at {server_url}")

    try:
        # Test health endpoint
        response = requests.get(f"{server_url}/health", timeout=5)
        if response.status_code == 200:
            print("✅ MCP server is running and healthy!")

            # Get server info
            try:
                info_response = requests.get(server_url, timeout=5)
                if info_response.status_code == 200:
                    print(f"\n📋 Server Info:")
                    print(info_response.text)
            except Exception as e:
                print(f"⚠️  Could not get server info: {e}")

            # List available tools
            try:
                tools_response = requests.get(f"{server_url}/tools", timeout=5)
                if tools_response.status_code == 200:
                    print(f"\n🔧 Available Tools:")
                    print(tools_response.text)
            except Exception as e:
                print(f"⚠️  Could not list tools: {e}")

            return True
        else:
            print(f"❌ Server responded with status {response.status_code}")
            return False

    except requests.exceptions.ConnectionError:
        print("❌ Cannot connect to MCP server")
        print("\n📝 Make sure:")
        print("   1. LibreOffice is running")
        print("   2. The extension is properly installed")
        print("   3. Try opening LibreOffice Writer first")
        print("\n💡 To start LibreOffice:")
        print('   "D:\\Program Files\\LibreOffice\\program\\swriter.exe"')
        return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    success = test_mcp_server()
    sys.exit(0 if success else 1)
