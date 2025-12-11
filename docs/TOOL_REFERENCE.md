# MCP-Libre Tool Reference

## Overview

MCP-Libre provides **25 tools** for interacting with LibreOffice Writer documents through the Model Context Protocol (MCP). These tools enable AI assistants to read, edit, navigate, and manipulate documents in real-time.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Claude Code / AI Assistant                    │
└──────────────────────────────┬──────────────────────────────────┘
                               │ MCP Protocol
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│           libreoffice_mcp_server.py (FastMCP Bridge)            │
│                    25 @mcp.tool functions                        │
└──────────────────────────────┬──────────────────────────────────┘
                               │ HTTP POST to localhost:8765
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                mcp_server.py (HTTP API Server)                   │
│          LibreOfficeMCPServer - Tool handlers                    │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                  uno_bridge.py (UNO Bridge)                      │
│         UNOBridge class - Direct LibreOffice UNO API             │
└──────────────────────────────┬──────────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                     LibreOffice Writer                           │
│                   (Running with document open)                    │
└─────────────────────────────────────────────────────────────────┘
```

## Tool Categories

### 1. Document Management (4 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `create_document` | Create new LibreOffice document | `doc_type`: "writer", "calc", "impress", "draw" |
| `get_document_info` | Get document metadata | None |
| `list_open_documents` | List all open documents | None |
| `check_libreoffice_status` | Check MCP server status | None |

### 2. Document Content (3 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `get_document_content` | Get full document text | None |
| `insert_text` | Insert text at cursor | `text`: string |
| `format_text` | Format selected text | `bold`, `italic`, `underline`, `font_size`, `font_name` |

### 3. Save & Export (2 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `save_document` | Save document | `file_path` (optional) |
| `export_document` | Export to different format | `file_path`, `export_format` |

### 4. Document Structure (4 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `get_paragraph_count` | Count paragraphs | None |
| `get_document_outline` | Get headings/outline | None |
| `get_paragraph` | Get specific paragraph | `n`: paragraph number (1-indexed) |
| `get_paragraphs_range` | Get range of paragraphs | `start`, `end` (1-indexed, inclusive) |

### 5. Cursor Navigation (4 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `goto_paragraph` | Move cursor to paragraph | `n`: paragraph number |
| `goto_position` | Move cursor to character | `char_pos`: position (0-indexed) |
| `get_cursor_position` | Get current cursor position | None |
| `get_context_around_cursor` | Get text around cursor | `chars`: characters before/after (default: 100) |

### 6. Text Selection (4 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `select_paragraph` | Select entire paragraph | `n`: paragraph number |
| `select_text_range` | Select character range | `start`, `end` (0-indexed, exclusive end) |
| `delete_selection` | Delete selected text | None |
| `replace_selection` | Replace selected text | `text`: replacement text |

### 7. Search & Replace (3 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `find_text` | Find all occurrences | `query`: search string |
| `find_and_replace` | Replace first occurrence | `old`, `new` |
| `find_and_replace_all` | Replace all occurrences | `old`, `new` |

### 8. Comments (2 tools)

| Tool | Description | Parameters |
|------|-------------|------------|
| `get_comments` | Get all document comments | None |
| `add_comment` | Add comment at cursor | `text`, `author` (default: "Claude") |

## Usage Examples

### Basic Document Editing
```
# Get document info
get_document_info()

# Insert text at cursor
insert_text("Hello, World!")

# Save document
save_document()
```

### Navigation and Selection
```
# Go to paragraph 5
goto_paragraph(5)

# Get surrounding context
get_context_around_cursor(200)

# Select a range and replace
select_text_range(100, 150)
replace_selection("New text")
```

### Search and Replace
```
# Find all occurrences
find_text("old phrase")

# Replace all instances
find_and_replace_all("old phrase", "new phrase")
```

### Working with Comments
```
# Get all comments
get_comments()

# Add a comment at current position
add_comment("This needs review", "Claude")
```

## Response Format

All tools return a dictionary with:

```python
{
    "success": True/False,
    "error": "Error message if success is False",
    # Additional tool-specific fields
}
```

### Common Response Fields by Tool Type

**Document Info:**
```python
{
    "success": True,
    "document_info": {
        "title": "document.odt",
        "url": "file:///path/to/document.odt",
        "modified": False,
        "type": "writer",
        "has_selection": True,
        "word_count": 1500,
        "character_count": 8500
    }
}
```

**Search Results:**
```python
{
    "success": True,
    "matches": [
        {"position": 100, "text": "matched text"},
        {"position": 500, "text": "matched text"}
    ],
    "count": 2,
    "query": "search term"
}
```

**Selection:**
```python
{
    "success": True,
    "selected_text": "The selected content",
    "start": 100,
    "end": 125,
    "length": 25
}
```

## Error Handling

Common errors and their meanings:

| Error | Meaning |
|-------|---------|
| "No document available" | No document is open in LibreOffice |
| "Cannot connect to LibreOffice" | MCP Server not started in LibreOffice |
| "Paragraph X out of range" | Requested paragraph doesn't exist |
| "No text selected" | Operation requires selection but none exists |
| "Text search not supported for X documents" | Tool only works with Writer documents |

## Best Practices

1. **Always check document state** - Use `get_document_info()` before making edits
2. **Use paragraphs for structure** - Navigate by paragraph when possible
3. **Verify selections** - Check `has_selection` before using selection-based tools
4. **Handle errors gracefully** - Always check `success` field in responses
5. **Save frequently** - Use `save_document()` after significant changes
