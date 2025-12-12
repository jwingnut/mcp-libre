#!/usr/bin/env python3
"""
LibreOffice MCP Server - FastMCP Bridge (Consolidated Tools)

This server bridges between the MCP protocol and the LibreOffice HTTP API,
allowing Claude Code and other MCP clients to control LibreOffice.

Tools have been consolidated from 32 individual tools to 9 logical groups
for better UX and fewer permission prompts.

Prerequisites:
    1. LibreOffice Writer must be running
    2. MCP Server must be started (Tools → MCP Server → Start MCP Server)

Usage:
    fastmcp run libreoffice_mcp_server.py

Or configure in Claude Code:
    claude mcp add libreoffice -- fastmcp run /path/to/libreoffice_mcp_server.py
"""

import httpx
from fastmcp import FastMCP

# LibreOffice HTTP API endpoint
LIBREOFFICE_URL = "http://localhost:8765"

# Create the MCP server
mcp = FastMCP("LibreOffice")


def call_libreoffice(path: str, method: str = "GET", data: dict = None) -> dict:
    """Make a request to the LibreOffice HTTP API"""
    url = f"{LIBREOFFICE_URL}{path}"
    try:
        if method == "GET":
            response = httpx.get(url, timeout=30)
        else:
            response = httpx.post(url, json=data or {}, timeout=30)
        return response.json()
    except httpx.ConnectError:
        return {"error": "Cannot connect to LibreOffice. Make sure LibreOffice is running and MCP Server is started (Tools → MCP Server → Start MCP Server)"}
    except Exception as e:
        return {"error": str(e)}


# =============================================================================
# CONSOLIDATED TOOL 1: document
# Actions: create, info, list, content, status
# =============================================================================

@mcp.tool
def document(action: str, doc_type: str = "writer") -> dict:
    """
    Manage LibreOffice documents - create, get info, list, get content, or check status.

    Args:
        action: The operation to perform. Options:
            - "create": Create a new document (use doc_type param)
            - "info": Get active document information including track_changes status
            - "list": List all open documents
            - "content": Get full text content of active document
            - "status": Check LibreOffice MCP server health
        doc_type: Type of document for "create" action. Options: "writer", "calc", "impress", "draw"

    Returns:
        Result based on action performed
    """
    if action == "create":
        return call_libreoffice("/tools/create_document_live", "POST", {"doc_type": doc_type})
    elif action == "info":
        return call_libreoffice("/tools/get_document_info_live", "POST", {})
    elif action == "list":
        return call_libreoffice("/tools/list_open_documents", "POST", {})
    elif action == "content":
        return call_libreoffice("/tools/get_text_content_live", "POST", {})
    elif action == "status":
        return call_libreoffice("/health")
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["create", "info", "list", "content", "status"]}


# =============================================================================
# CONSOLIDATED TOOL 2: structure
# Actions: outline, paragraph, range, count
# =============================================================================

@mcp.tool
def structure(action: str, n: int = None, start: int = None, end: int = None) -> dict:
    """
    Navigate and inspect document structure - outline, paragraphs, ranges, count.

    Args:
        action: The operation to perform. Options:
            - "outline": Get document headings with paragraph numbers and levels
            - "paragraph": Get specific paragraph content (requires n)
            - "range": Get paragraphs in a range (requires start, end)
            - "count": Get total paragraph count
        n: Paragraph number (1-indexed) for "paragraph" action
        start: Starting paragraph number for "range" action
        end: Ending paragraph number for "range" action

    Returns:
        Result based on action performed
    """
    if action == "outline":
        return call_libreoffice("/tools/get_document_outline_live", "POST", {})
    elif action == "paragraph":
        if n is None:
            return {"error": "Action 'paragraph' requires parameter 'n' (paragraph number)"}
        return call_libreoffice("/tools/get_paragraph_live", "POST", {"n": n})
    elif action == "range":
        if start is None or end is None:
            return {"error": "Action 'range' requires parameters 'start' and 'end'"}
        return call_libreoffice("/tools/get_paragraphs_range_live", "POST", {"start": start, "end": end})
    elif action == "count":
        return call_libreoffice("/tools/get_paragraph_count_live", "POST", {})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["outline", "paragraph", "range", "count"]}


# =============================================================================
# CONSOLIDATED TOOL 3: cursor
# Actions: goto_paragraph, goto_position, position, context
# =============================================================================

@mcp.tool
def cursor(action: str, n: int = None, char_pos: int = None, chars: int = 100) -> dict:
    """
    Navigate cursor position in the document.

    Args:
        action: The operation to perform. Options:
            - "goto_paragraph": Move cursor to paragraph n (requires n)
            - "goto_position": Move cursor to character position (requires char_pos)
            - "position": Get current cursor position and paragraph number
            - "context": Get text around cursor (optional chars param, default 100)
        n: Paragraph number (1-indexed) for "goto_paragraph" action
        char_pos: Character position (0-indexed) for "goto_position" action
        chars: Number of characters before/after cursor for "context" action (default: 100)

    Returns:
        Result based on action performed
    """
    if action == "goto_paragraph":
        if n is None:
            return {"error": "Action 'goto_paragraph' requires parameter 'n' (paragraph number)"}
        return call_libreoffice("/tools/goto_paragraph_live", "POST", {"n": n})
    elif action == "goto_position":
        if char_pos is None:
            return {"error": "Action 'goto_position' requires parameter 'char_pos' (character position)"}
        return call_libreoffice("/tools/goto_position_live", "POST", {"char_pos": char_pos})
    elif action == "position":
        return call_libreoffice("/tools/get_cursor_position_live", "POST", {})
    elif action == "context":
        return call_libreoffice("/tools/get_context_around_cursor_live", "POST", {"chars": chars})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["goto_paragraph", "goto_position", "position", "context"]}


# =============================================================================
# CONSOLIDATED TOOL 4: selection
# Actions: paragraph, range, delete, replace
# =============================================================================

@mcp.tool
def selection(action: str, n: int = None, start: int = None, end: int = None, text: str = None) -> dict:
    """
    Select and manipulate text ranges in the document.

    Args:
        action: The operation to perform. Options:
            - "paragraph": Select entire paragraph (requires n)
            - "range": Select character range (requires start, end)
            - "delete": Delete current selection
            - "replace": Replace selection with new text (requires text)
        n: Paragraph number (1-indexed) for "paragraph" action
        start: Starting character position (0-indexed) for "range" action
        end: Ending character position (exclusive) for "range" action
        text: Replacement text for "replace" action

    Returns:
        Result based on action performed
    """
    if action == "paragraph":
        if n is None:
            return {"error": "Action 'paragraph' requires parameter 'n' (paragraph number)"}
        return call_libreoffice("/tools/select_paragraph_live", "POST", {"n": n})
    elif action == "range":
        if start is None or end is None:
            return {"error": "Action 'range' requires parameters 'start' and 'end'"}
        return call_libreoffice("/tools/select_text_range_live", "POST", {"start": start, "end": end})
    elif action == "delete":
        return call_libreoffice("/tools/delete_selection_live", "POST", {})
    elif action == "replace":
        if text is None:
            return {"error": "Action 'replace' requires parameter 'text'"}
        return call_libreoffice("/tools/replace_selection_live", "POST", {"text": text})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["paragraph", "range", "delete", "replace"]}


# =============================================================================
# CONSOLIDATED TOOL 5: search
# Actions: find, replace, replace_all
# =============================================================================

@mcp.tool
def search(action: str, query: str = None, old: str = None, new: str = None) -> dict:
    """
    Find and replace text in the document. Track Changes aware - skips tracked deletions.

    Args:
        action: The operation to perform. Options:
            - "find": Find all occurrences of text (requires query)
            - "replace": Replace first occurrence (requires old, new)
            - "replace_all": Replace all occurrences (requires old, new)
        query: Text to search for ("find" action)
        old: Text to find ("replace" and "replace_all" actions)
        new: Replacement text ("replace" and "replace_all" actions)

    Returns:
        Result with matches or replacement count. Includes track_changes_active field.
    """
    if action == "find":
        if query is None:
            return {"error": "Action 'find' requires parameter 'query'"}
        return call_libreoffice("/tools/find_text_live", "POST", {"query": query})
    elif action == "replace":
        if old is None or new is None:
            return {"error": "Action 'replace' requires parameters 'old' and 'new'"}
        return call_libreoffice("/tools/find_and_replace_live", "POST", {"old": old, "new": new})
    elif action == "replace_all":
        if old is None or new is None:
            return {"error": "Action 'replace_all' requires parameters 'old' and 'new'"}
        return call_libreoffice("/tools/find_and_replace_all_live", "POST", {"old": old, "new": new})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["find", "replace", "replace_all"]}


# =============================================================================
# CONSOLIDATED TOOL 6: track_changes
# Actions: status, enable, disable, list, accept, reject, accept_all, reject_all
# =============================================================================

@mcp.tool
def track_changes(action: str, index: int = None, show: bool = True) -> dict:
    """
    Manage Track Changes / revision tracking in the document.

    Args:
        action: The operation to perform. Options:
            - "status": Get recording/showing status and pending count
            - "enable": Enable Track Changes recording (optional show param)
            - "disable": Disable Track Changes recording
            - "list": List all tracked changes with details
            - "accept": Accept specific change by index (requires index)
            - "reject": Reject specific change by index (requires index)
            - "accept_all": Accept all tracked changes
            - "reject_all": Reject all tracked changes
        index: Change index (0-based) for "accept" and "reject" actions
        show: Whether to show tracked changes when enabling (default: True)

    Returns:
        Result based on action performed
    """
    if action == "status":
        return call_libreoffice("/tools/get_track_changes_status_live", "POST", {})
    elif action == "enable":
        return call_libreoffice("/tools/set_track_changes_live", "POST", {"enabled": True, "show": show})
    elif action == "disable":
        return call_libreoffice("/tools/set_track_changes_live", "POST", {"enabled": False, "show": show})
    elif action == "list":
        return call_libreoffice("/tools/get_tracked_changes_live", "POST", {})
    elif action == "accept":
        if index is None:
            return {"error": "Action 'accept' requires parameter 'index' (0-based change index)"}
        return call_libreoffice("/tools/accept_tracked_change_live", "POST", {"index": index})
    elif action == "reject":
        if index is None:
            return {"error": "Action 'reject' requires parameter 'index' (0-based change index)"}
        return call_libreoffice("/tools/reject_tracked_change_live", "POST", {"index": index})
    elif action == "accept_all":
        return call_libreoffice("/tools/accept_all_changes_live", "POST", {})
    elif action == "reject_all":
        return call_libreoffice("/tools/reject_all_changes_live", "POST", {})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["status", "enable", "disable", "list", "accept", "reject", "accept_all", "reject_all"]}


# =============================================================================
# CONSOLIDATED TOOL 7: comments
# Actions: list, add
# =============================================================================

@mcp.tool
def comments(action: str, text: str = None, author: str = "Claude") -> dict:
    """
    Manage document comments/annotations.

    Args:
        action: The operation to perform. Options:
            - "list": Get all comments with author, content, date, anchor text
            - "add": Add comment at cursor position (requires text)
        text: Comment text for "add" action
        author: Author name for "add" action (default: "Claude")

    Returns:
        Result based on action performed
    """
    if action == "list":
        return call_libreoffice("/tools/get_comments_live", "POST", {})
    elif action == "add":
        if text is None:
            return {"error": "Action 'add' requires parameter 'text'"}
        return call_libreoffice("/tools/add_comment_live", "POST", {"text": text, "author": author})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["list", "add"]}


# =============================================================================
# CONSOLIDATED TOOL 8: save
# Actions: save, export
# =============================================================================

@mcp.tool
def save(action: str, file_path: str = None, export_format: str = "pdf") -> dict:
    """
    Save and export documents.

    Args:
        action: The operation to perform. Options:
            - "save": Save document (optional file_path for Save As)
            - "export": Export to format (requires file_path and export_format)
        file_path: Path to save/export to
        export_format: Format for export. Options: "pdf", "docx", "odt", "html", "txt"

    Returns:
        Result with success status
    """
    if action == "save":
        data = {}
        if file_path:
            data["file_path"] = file_path
        return call_libreoffice("/tools/save_document_live", "POST", data)
    elif action == "export":
        if file_path is None:
            return {"error": "Action 'export' requires parameter 'file_path'"}
        return call_libreoffice("/tools/export_document_live", "POST", {
            "file_path": file_path,
            "format": export_format
        })
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["save", "export"]}


# =============================================================================
# CONSOLIDATED TOOL 9: text
# Actions: insert, format
# =============================================================================

@mcp.tool
def text(action: str, content: str = None, bold: bool = None, italic: bool = None,
         underline: bool = None, font_size: int = None, font_name: str = None) -> dict:
    """
    Insert and format text in the document.

    Args:
        action: The operation to perform. Options:
            - "insert": Insert text at cursor position (requires content)
            - "format": Apply formatting to selected text (use formatting params)
        content: Text to insert for "insert" action
        bold: Set bold formatting (True/False) for "format" action
        italic: Set italic formatting (True/False) for "format" action
        underline: Set underline formatting (True/False) for "format" action
        font_size: Font size in points for "format" action
        font_name: Font family name for "format" action

    Returns:
        Result with success status
    """
    if action == "insert":
        if content is None:
            return {"error": "Action 'insert' requires parameter 'content' (text to insert)"}
        return call_libreoffice("/tools/insert_text_live", "POST", {"text": content})
    elif action == "format":
        formatting = {}
        if bold is not None:
            formatting["bold"] = bold
        if italic is not None:
            formatting["italic"] = italic
        if underline is not None:
            formatting["underline"] = underline
        if font_size is not None:
            formatting["font_size"] = font_size
        if font_name is not None:
            formatting["font_name"] = font_name
        return call_libreoffice("/tools/format_text_live", "POST", {"formatting": formatting})
    else:
        return {"error": f"Invalid action '{action}'", "valid_actions": ["insert", "format"]}


if __name__ == "__main__":
    mcp.run()
