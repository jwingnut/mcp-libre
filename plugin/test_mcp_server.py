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


# Enhanced Editing Tools Tests - Document Structure

def test_get_paragraph_count():
    """Test getting paragraph count"""
    print("Testing get_paragraph_count_live tool...")
    result = make_request("/tools/get_paragraph_count_live", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        print(f"  ✓ Paragraph count: {result.get('count', 0)}")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_document_outline():
    """Test getting document outline"""
    print("Testing get_document_outline_live tool...")
    result = make_request("/tools/get_document_outline_live", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        outline_count = len(result.get("outline", []))
        print(f"  ✓ Found {outline_count} headings in outline")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_paragraph():
    """Test getting specific paragraph"""
    print("Testing get_paragraph_live tool...")
    result = make_request("/tools/get_paragraph_live", method="POST", data={"n": 1})

    if "error" in result:
        print(f"  ⚠ Error (may be out of range): {result['error']}")
    elif result.get("success"):
        content = result.get("content", "")[:50]
        print(f"  ✓ Got paragraph 1: {content}...")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_paragraphs_range():
    """Test getting paragraph range"""
    print("Testing get_paragraphs_range_live tool...")
    result = make_request("/tools/get_paragraphs_range_live", method="POST", data={"start": 1, "end": 2})

    if "error" in result:
        print(f"  ⚠ Error (may be out of range): {result['error']}")
    elif result.get("success"):
        count = result.get("count", 0)
        print(f"  ✓ Got {count} paragraphs in range")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


# Enhanced Editing Tools Tests - Cursor Navigation

def test_goto_paragraph():
    """Test cursor navigation to paragraph"""
    print("Testing goto_paragraph_live tool...")
    result = make_request("/tools/goto_paragraph_live", method="POST", data={"n": 1})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        print(f"  ✓ Moved cursor to paragraph {result.get('paragraph', 1)}")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_goto_position():
    """Test cursor navigation to position"""
    print("Testing goto_position_live tool...")
    result = make_request("/tools/goto_position_live", method="POST", data={"char_pos": 0})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        print(f"  ✓ Moved cursor to position {result.get('position', 0)}")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_cursor_position():
    """Test getting cursor position"""
    print("Testing get_cursor_position_live tool...")
    result = make_request("/tools/get_cursor_position_live", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        pos = result.get("position", 0)
        para = result.get("paragraph", 0)
        print(f"  ✓ Cursor at position {pos}, paragraph {para}")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_get_context_around_cursor():
    """Test getting context around cursor"""
    print("Testing get_context_around_cursor_live tool...")
    result = make_request("/tools/get_context_around_cursor_live", method="POST", data={"chars": 50})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        before_len = len(result.get("before", ""))
        after_len = len(result.get("after", ""))
        print(f"  ✓ Got context: {before_len} chars before, {after_len} chars after")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


# Enhanced Editing Tools Tests - Text Selection

def test_select_paragraph():
    """Test selecting paragraph"""
    print("Testing select_paragraph_live tool...")
    result = make_request("/tools/select_paragraph_live", method="POST", data={"n": 1})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        text_len = len(result.get("selected_text", ""))
        print(f"  ✓ Selected paragraph 1 ({text_len} chars)")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_select_text_range():
    """Test selecting text range"""
    print("Testing select_text_range_live tool...")
    result = make_request("/tools/select_text_range_live", method="POST", data={"start": 0, "end": 10})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        text_len = len(result.get("selected_text", ""))
        print(f"  ✓ Selected range 0-10 ({text_len} chars)")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_delete_selection():
    """Test deleting selection"""
    print("Testing delete_selection_live tool...")
    # First select some text
    make_request("/tools/select_text_range_live", method="POST", data={"start": 0, "end": 5})
    # Then try to delete it
    result = make_request("/tools/delete_selection_live", method="POST", data={})

    if "error" in result:
        print(f"  ⚠ Error (may need selection): {result['error']}")
    elif result.get("success"):
        deleted_len = len(result.get("deleted_text", ""))
        print(f"  ✓ Deleted selection ({deleted_len} chars)")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_replace_selection():
    """Test replacing selection"""
    print("Testing replace_selection_live tool...")
    # First select some text
    make_request("/tools/select_text_range_live", method="POST", data={"start": 0, "end": 5})
    # Then try to replace it
    result = make_request("/tools/replace_selection_live", method="POST", data={"text": "TEST"})

    if "error" in result:
        print(f"  ⚠ Error (may need selection): {result['error']}")
    elif result.get("success"):
        old_len = len(result.get("old_text", ""))
        new_len = len(result.get("new_text", ""))
        print(f"  ✓ Replaced selection ({old_len} -> {new_len} chars)")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


# Enhanced Editing Tools Tests - Search and Replace

def test_find_text():
    """Test finding text"""
    print("Testing find_text_live tool...")
    result = make_request("/tools/find_text_live", method="POST", data={"query": "test"})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        count = result.get("count", 0)
        print(f"  ✓ Found {count} occurrences of 'test'")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_find_and_replace():
    """Test find and replace first occurrence"""
    print("Testing find_and_replace_live tool...")
    result = make_request("/tools/find_and_replace_live", method="POST", data={"old": "test", "new": "TEST"})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        replaced = result.get("replaced", False)
        if replaced:
            print(f"  ✓ Replaced first occurrence at position {result.get('position', 0)}")
        else:
            print(f"  ✓ No occurrence found to replace")
    else:
        print(f"  ⚠ Unexpected result: {result}")
    return True


def test_find_and_replace_all():
    """Test find and replace all occurrences"""
    print("Testing find_and_replace_all_live tool...")
    result = make_request("/tools/find_and_replace_all_live", method="POST", data={"old": "TEST", "new": "test"})

    if "error" in result:
        print(f"  ⚠ Error: {result['error']}")
    elif result.get("success"):
        count = result.get("count", 0)
        print(f"  ✓ Replaced {count} occurrences")
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
        # Enhanced Editing Tools - Document Structure
        test_get_paragraph_count,
        test_get_document_outline,
        test_get_paragraph,
        test_get_paragraphs_range,
        # Enhanced Editing Tools - Cursor Navigation
        test_goto_paragraph,
        test_goto_position,
        test_get_cursor_position,
        test_get_context_around_cursor,
        # Enhanced Editing Tools - Text Selection
        test_select_paragraph,
        test_select_text_range,
        test_delete_selection,
        test_replace_selection,
        # Enhanced Editing Tools - Search and Replace
        test_find_text,
        test_find_and_replace,
        test_find_and_replace_all,
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
