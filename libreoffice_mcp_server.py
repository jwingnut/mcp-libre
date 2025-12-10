#!/usr/bin/env python3
"""
LibreOffice MCP Server - FastMCP Bridge

This server bridges between the MCP protocol and the LibreOffice HTTP API,
allowing Claude Code and other MCP clients to control LibreOffice.

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


@mcp.tool
def insert_text(text: str) -> dict:
    """
    Insert text into the active LibreOffice Writer document at the current cursor position.

    Args:
        text: The text to insert into the document

    Returns:
        Result with success status and message
    """
    return call_libreoffice("/tools/insert_text_live", "POST", {"text": text})


@mcp.tool
def get_document_content() -> dict:
    """
    Get the full text content of the active LibreOffice Writer document.

    Returns:
        The document content as text, or an error message
    """
    return call_libreoffice("/tools/get_text_content_live", "POST", {})


@mcp.tool
def get_document_info() -> dict:
    """
    Get information about the active LibreOffice document including title, type, and modification status.

    Returns:
        Document metadata including title, type, word count, etc.
    """
    return call_libreoffice("/tools/get_document_info_live", "POST", {})


@mcp.tool
def list_open_documents() -> dict:
    """
    List all currently open documents in LibreOffice.

    Returns:
        List of open documents with their titles and types
    """
    return call_libreoffice("/tools/list_open_documents", "POST", {})


@mcp.tool
def create_document(doc_type: str = "writer") -> dict:
    """
    Create a new LibreOffice document.

    Args:
        doc_type: Type of document to create. Options: "writer", "calc", "impress", "draw"

    Returns:
        Result with success status
    """
    return call_libreoffice("/tools/create_document_live", "POST", {"doc_type": doc_type})


@mcp.tool
def save_document(file_path: str = None) -> dict:
    """
    Save the active LibreOffice document.

    Args:
        file_path: Optional path to save to. If not provided, saves to current location.

    Returns:
        Result with success status
    """
    data = {}
    if file_path:
        data["file_path"] = file_path
    return call_libreoffice("/tools/save_document_live", "POST", data)


@mcp.tool
def export_document(file_path: str, export_format: str = "pdf") -> dict:
    """
    Export the active LibreOffice document to a different format.

    Args:
        file_path: Path to export to (e.g., "/home/user/document.pdf")
        export_format: Format to export as. Options: "pdf", "docx", "odt", "html", "txt"

    Returns:
        Result with success status
    """
    return call_libreoffice("/tools/export_document_live", "POST", {
        "file_path": file_path,
        "format": export_format
    })


@mcp.tool
def format_text(bold: bool = None, italic: bool = None, underline: bool = None,
                font_size: int = None, font_name: str = None) -> dict:
    """
    Apply formatting to the currently selected text in LibreOffice Writer.

    Args:
        bold: Set text to bold (True/False)
        italic: Set text to italic (True/False)
        underline: Set text to underline (True/False)
        font_size: Font size in points (e.g., 12, 14, 16)
        font_name: Font family name (e.g., "Arial", "Times New Roman")

    Returns:
        Result with success status
    """
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


@mcp.tool
def check_libreoffice_status() -> dict:
    """
    Check if LibreOffice MCP server is running and accessible.

    Returns:
        Health status of the LibreOffice MCP server
    """
    return call_libreoffice("/health")


@mcp.tool
def get_comments() -> dict:
    """
    Get all comments/annotations from the active LibreOffice document.

    Returns:
        List of comments with author, content, date, and anchor text
    """
    return call_libreoffice("/tools/get_comments_live", "POST", {})


@mcp.tool
def add_comment(text: str, author: str = "Claude") -> dict:
    """
    Add a comment at the current cursor position in LibreOffice Writer.

    Args:
        text: The comment text
        author: Author name for the comment (default: "Claude")

    Returns:
        Result with success status
    """
    return call_libreoffice("/tools/add_comment_live", "POST", {"text": text, "author": author})


if __name__ == "__main__":
    mcp.run()
