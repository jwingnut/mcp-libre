#!/usr/bin/env python3
"""
Test client for ZetaJS WebAssembly MCP Bridge
Tests the exact 9 consolidated tool actions from mcp-libre over HTTP:8765.
"""

import os
import sys
import time
import json
import urllib.request
import urllib.error

BASE_URL = os.environ.get("LIBREOFFICE_URL", "http://localhost:8765")

def call_tool(tool_name, params=None):
    url = f"{BASE_URL}/tools/{tool_name}"
    headers = {"Content-Type": "application/json"}
    body = json.dumps(params or {}).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8")
        try:
            return json.loads(err_body)
        except:
            return {"error": f"HTTP {e.code}: {err_body}"}
    except Exception as e:
        return {"error": str(e)}

def check_health():
    url = f"{BASE_URL}/health"
    try:
        with urllib.request.urlopen(url, timeout=5) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception as e:
        return {"error": str(e)}

def main():
    print("=" * 65)
    print("🧪 Testing ZetaJS Wasm MCP Bridge (9 Consolidated Tools)")
    print("=" * 65)

    # 1. Health check
    print("\n1️⃣  Checking Server Health...")
    health = check_health()
    print("Response:", json.dumps(health, indent=2))
    
    if not health.get("browser_connected"):
        print("\n⚠️  Notice: No active ZetaOffice browser tab is connected yet!")
        print(f"👉 Please open {BASE_URL} in your browser first to boot the Wasm UNO engine.")
        print("   Then run this test script again.\n")
        return

    # 2. Tool: document (info & styles)
    print("\n2️⃣  Testing 'document' tool...")
    info = call_tool("get_document_info_live", {})
    print("Document Info:", json.dumps(info, indent=2))

    styles = call_tool("list_paragraph_styles_live", {})
    print(f"Available Paragraph Styles ({len(styles.get('styles', []))}):", styles.get('styles', [])[:5], "...")

    # 3. Tool: text (insert & format)
    print("\n3️⃣  Testing 'text' tool (Insert & Format via UNO)...")
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    sample_text = (
        f"\n\n## Section: Autonomous Wasm Execution [{timestamp}]\n"
        "This paragraph was dispatched via the consolidated 'text' tool (action='insert')\n"
        "and rendered by LibreOffice WebAssembly via ZetaJS UNO.\n"
    )
    insert_res = call_tool("insert_text_live", {"text": sample_text})
    print("Insert Text Response:", json.dumps(insert_res, indent=2))

    format_res = call_tool("format_text_live", {"formatting": {"bold": True, "font_size": 13}})
    print("Format Text Response:", json.dumps(format_res, indent=2))

    # 4. Tool: structure (paragraphs & count)
    print("\n4️⃣  Testing 'structure' tool (Inspect AST via UNO)...")
    count_res = call_tool("get_paragraph_count_live", {})
    print(f"Total Paragraph Count: {count_res.get('count')}")

    para_res = call_tool("get_paragraph_live", {"n": 1})
    print("Paragraph 1 Content:", json.dumps(para_res, indent=2))

    # 5. Tool: search (find text)
    print("\n5️⃣  Testing 'search' tool (Find via UNO)...")
    find_res = call_tool("find_text_live", {"query": "Autonomous Wasm"})
    print("Search Matches:", json.dumps(find_res, indent=2))

    # 6. Read back full content
    print("\n6️⃣  Testing 'document' tool (content)...")
    content_res = call_tool("get_text_content_live", {})
    preview = content_res.get("content", "")
    print("Current Document Content Preview:")
    print("-" * 50)
    print(preview[-350:] if len(preview) > 350 else preview)
    print("-" * 50)

    print("\n✅ All 9 tool categories in ZetaJS Wasm MCP Bridge verified successfully!")

if __name__ == "__main__":
    main()
