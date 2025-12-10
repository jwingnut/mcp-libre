#!/usr/bin/env python3
"""
Tests for LibreOffice MCP Server

These tests verify the HTTP API functionality. Run with:
    python3 test_mcp_server.py

Prerequisites:
    - LibreOffice Writer must be running
    - MCP Server must be started (Tools → MCP Server → Start MCP Server)
"""

import json
import urllib.request
import urllib.error
import sys

BASE_URL = "http://localhost:8765"


def make_request(path, method="GET", data=None):
    """Make HTTP request to MCP server"""
    url = f"{BASE_URL}{path}"
    headers = {"Content-Type": "application/json"}

    if data:
        data = json.dumps(data).encode("utf-8")

    req = urllib.request.Request(url, data=data, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.URLError as e:
        return {"error": str(e)}
    except json.JSONDecodeError:
        return {"error": "Invalid JSON response"}


def test_health():
    """Test health endpoint"""
    print("Testing /health endpoint...")
    result = make_request("/health")

    assert "status" in result, f"Expected 'status' in response, got: {result}"
    assert result["status"] == "healthy", f"Expected healthy status, got: {result}"
    print("  ✓ Health check passed")
    return True


def test_server_info():
    """Test root endpoint returns server info"""
    print("Testing / endpoint (server info)...")
    result = make_request("/")

    assert "name" in result, f"Expected 'name' in response, got: {result}"
    assert "version" in result, f"Expected 'version' in response, got: {result}"
    assert "endpoints" in result, f"Expected 'endpoints' in response, got: {result}"
    print(f"  ✓ Server: {result['name']} v{result['version']}")
    return True


def test_list_tools():
    """Test tools endpoint returns list of available tools"""
    print("Testing /tools endpoint...")
    result = make_request("/tools")

    assert "tools" in result, f"Expected 'tools' in response, got: {result}"
    assert isinstance(result["tools"], list), f"Expected tools to be a list, got: {type(result['tools'])}"
    assert len(result["tools"]) > 0, "Expected at least one tool"

    tool_names = [t["name"] for t in result["tools"]]
    print(f"  ✓ Found {len(tool_names)} tools: {', '.join(tool_names)}")

    # Check for expected tools
    expected_tools = ["create_document_live", "insert_text_live", "get_document_info_live"]
    for tool in expected_tools:
        assert tool in tool_names, f"Expected tool '{tool}' not found"
    print(f"  ✓ All expected tools present")
    return True


def test_get_document_info():
    """Test getting document info"""
    print("Testing get_document_info_live tool...")
    result = make_request("/tools/get_document_info_live", method="POST", data={})

    # Should return info or error if no document open
    if "error" in result:
        print(f"  ⚠ No document open or error: {result['error']}")
    else:
        print(f"  ✓ Document info: {result}")
    return True


def test_list_open_documents():
    """Test listing open documents"""
    print("Testing list_open_documents tool...")
    result = make_request("/tools/list_open_documents", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    else:
        docs = result.get("documents", [])
        print(f"  ✓ Found {len(docs)} open document(s)")
    return True


def test_insert_text():
    """Test inserting text into document"""
    print("Testing insert_text_live tool...")
    test_text = "Hello from MCP test! "
    result = make_request("/tools/insert_text_live", method="POST", data={"text": test_text})

    if "error" in result:
        print(f"  ⚠ Error (may need Writer document open): {result['error']}")
    elif result.get("success"):
        print(f"  ✓ Text inserted successfully")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_text_content():
    """Test getting text content from document"""
    print("Testing get_text_content_live tool...")
    result = make_request("/tools/get_text_content_live", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif "content" in result:
        content_preview = result["content"][:100] + "..." if len(result.get("content", "")) > 100 else result.get("content", "")
        print(f"  ✓ Got content: {content_preview}")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def run_all_tests():
    """Run all tests"""
    print("=" * 60)
    print("LibreOffice MCP Server Tests")
    print("=" * 60)
    print()

    # Check if server is running
    print("Checking server connection...")
    try:
        result = make_request("/health")
        if "error" in result:
            print(f"\n❌ Cannot connect to MCP server at {BASE_URL}")
            print("   Make sure LibreOffice is running and MCP Server is started.")
            print("   (Tools → MCP Server → Start MCP Server)")
            return False
    except Exception as e:
        print(f"\n❌ Connection error: {e}")
        return False

    print("✓ Server is running\n")

    tests = [
        test_health,
        test_server_info,
        test_list_tools,
        test_list_open_documents,
        test_get_document_info,
        test_insert_text,
        test_get_text_content,
    ]

    passed = 0
    failed = 0

    for test in tests:
        try:
            test()
            passed += 1
        except AssertionError as e:
            print(f"  ❌ FAILED: {e}")
            failed += 1
        except Exception as e:
            print(f"  ❌ ERROR: {e}")
            failed += 1
        print()

    print("=" * 60)
    print(f"Results: {passed} passed, {failed} failed")
    print("=" * 60)

    return failed == 0


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
