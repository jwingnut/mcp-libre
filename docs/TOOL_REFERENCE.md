# MCP-Libre Tool Reference

## Overview

MCP-Libre provides **9 consolidated tools** for interacting with LibreOffice Writer documents through the Model Context Protocol (MCP). These tools enable AI assistants to read, edit, navigate, and manipulate documents in real-time.

The tools have been consolidated from 32 individual tools into 9 logical groups, reducing permission prompts while maintaining full functionality.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Claude Code / AI Assistant                    │
└──────────────────────────────┬──────────────────────────────────┘
                               │ MCP Protocol
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│           libreoffice_mcp_server.py (FastMCP Bridge)            │
│                    9 @mcp.tool functions                         │
└──────────────────────────────┬──────────────────────────────────┘
                               │ HTTP POST to localhost:8765
                               ▼
┌─────────────────────────────────────────────────────────────────┐
│                mcp_server.py (HTTP API Server)                   │
│          LibreOfficeMCPServer - 32 Tool handlers                 │
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

## Consolidated Tools

### 1. document

Manage LibreOffice documents - create, get info, list, get content, or check status.

| Action | Description | Parameters |
|--------|-------------|------------|
| `create` | Create new document | `doc_type`: "writer", "calc", "impress", "draw" |
| `info` | Get document metadata (includes track_changes status) | None |
| `list` | List all open documents | None |
| `content` | Get full document text | None |
| `status` | Check MCP server health | None |

**Example:**
```python
document(action="info")
document(action="create", doc_type="writer")
document(action="list")
```

### 2. structure

Navigate and inspect document structure.

| Action | Description | Parameters |
|--------|-------------|------------|
| `outline` | Get headings with paragraph numbers and levels | None |
| `paragraph` | Get specific paragraph content | `n`: paragraph number (1-indexed) |
| `range` | Get range of paragraphs | `start`, `end`: paragraph numbers |
| `count` | Get total paragraph count | None |

**Example:**
```python
structure(action="count")
structure(action="paragraph", n=5)
structure(action="range", start=1, end=10)
```

### 3. cursor

Navigate cursor position in the document.

| Action | Description | Parameters |
|--------|-------------|------------|
| `goto_paragraph` | Move cursor to paragraph n | `n`: paragraph number (1-indexed) |
| `goto_position` | Move cursor to character position | `char_pos`: position (0-indexed) |
| `position` | Get current cursor position and paragraph | None |
| `context` | Get text around cursor | `chars`: characters before/after (default: 100) |

**Example:**
```python
cursor(action="position")
cursor(action="goto_paragraph", n=5)
cursor(action="context", chars=50)
```

### 4. selection

Select and manipulate text ranges.

| Action | Description | Parameters |
|--------|-------------|------------|
| `paragraph` | Select entire paragraph | `n`: paragraph number (1-indexed) |
| `range` | Select character range | `start`, `end`: positions (0-indexed, exclusive end) |
| `delete` | Delete current selection | None |
| `replace` | Replace selection with new text | `text`: replacement text |

**Example:**
```python
selection(action="paragraph", n=3)
selection(action="range", start=100, end=150)
selection(action="replace", text="New text")
selection(action="delete")
```

### 5. search

Find and replace text (Track Changes aware - skips tracked deletions).

| Action | Description | Parameters |
|--------|-------------|------------|
| `find` | Find all occurrences | `query`: search string |
| `replace` | Replace first occurrence | `old`, `new`: strings |
| `replace_all` | Replace all occurrences | `old`, `new`: strings |

**Example:**
```python
search(action="find", query="hello")
search(action="replace", old="old text", new="new text")
search(action="replace_all", old="foo", new="bar")
```

**Note:** Returns `track_changes_active: true` when Track Changes is enabled.

### 6. track_changes

Manage Track Changes / revision tracking.

| Action | Description | Parameters |
|--------|-------------|------------|
| `status` | Get recording/showing status and pending count | None |
| `enable` | Enable Track Changes recording | `show`: bool (default: true) |
| `disable` | Disable Track Changes recording | None |
| `list` | List all tracked changes with details | None |
| `accept` | Accept specific change by index | `index`: 0-based change index |
| `reject` | Reject specific change by index | `index`: 0-based change index |
| `accept_all` | Accept all tracked changes | None |
| `reject_all` | Reject all tracked changes | None |

**Example:**
```python
track_changes(action="status")
track_changes(action="enable")
track_changes(action="list")
track_changes(action="accept", index=0)
track_changes(action="accept_all")
```

### 7. comments

Manage document comments/annotations.

| Action | Description | Parameters |
|--------|-------------|------------|
| `list` | Get all comments | None |
| `add` | Add comment at cursor position | `text`, `author` (default: "Claude") |

**Example:**
```python
comments(action="list")
comments(action="add", text="This needs review", author="Claude")
```

### 8. save

Save and export documents.

| Action | Description | Parameters |
|--------|-------------|------------|
| `save` | Save document | `file_path` (optional, for Save As) |
| `export` | Export to different format | `file_path`, `export_format` |

**Example:**
```python
save(action="save")
save(action="save", file_path="/home/user/doc.odt")
save(action="export", file_path="/home/user/doc.pdf", export_format="pdf")
```

**Export formats:** pdf, docx, odt, html, txt

### 9. text

Insert and format text.

| Action | Description | Parameters |
|--------|-------------|------------|
| `insert` | Insert text at cursor position | `content`: text to insert |
| `format` | Apply formatting to selection | `bold`, `italic`, `underline`, `font_size`, `font_name` |

**Example:**
```python
text(action="insert", content="Hello, World!")
text(action="format", bold=True, font_size=16)
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

### Error Response for Invalid Actions

```python
{
    "error": "Invalid action 'xyz'",
    "valid_actions": ["action1", "action2", ...]
}
```

## Error Handling

Common errors and their meanings:

| Error | Meaning |
|-------|---------|
| "Invalid action 'xyz'" | Action parameter not recognized |
| "Action 'X' requires parameter 'Y'" | Missing required parameter |
| "No document available" | No document is open in LibreOffice |
| "Cannot connect to LibreOffice" | MCP Server not started in LibreOffice |
| "Paragraph X out of range" | Requested paragraph doesn't exist |
| "No text selected" | Operation requires selection but none exists |

## Best Practices

1. **Check document state first** - Use `document(action="info")` before making edits
2. **Use structure for navigation** - Navigate by paragraph when possible
3. **Verify selections** - Check `has_selection` before using selection-based actions
4. **Handle errors gracefully** - Always check `success` field in responses
5. **Save frequently** - Use `save(action="save")` after significant changes
6. **Use Track Changes** - Enable `track_changes(action="enable")` for collaborative editing
