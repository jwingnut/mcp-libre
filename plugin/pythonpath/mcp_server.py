"""
LibreOffice MCP Extension - MCP Server Module

This module implements an embedded MCP server that integrates with LibreOffice
via the UNO API, providing real-time document manipulation capabilities.
"""

import asyncio
import json
import logging
from typing import Dict, Any, Optional, List
import uno_bridge

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class LibreOfficeMCPServer:
    """Embedded MCP server for LibreOffice plugin"""
    
    def __init__(self):
        """Initialize the MCP server"""
        self.uno_bridge = uno_bridge.UNOBridge()
        self.tools = {}
        self._register_tools()
        logger.info("LibreOffice MCP Server initialized")
    
    def _register_tools(self):
        """Register all available MCP tools"""
        
        # Document creation tools
        self.tools["create_document_live"] = {
            "description": "Create a new document in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["writer", "calc", "impress", "draw"],
                        "description": "Type of document to create",
                        "default": "writer"
                    }
                }
            },
            "handler": self.create_document_live
        }
        
        # Text manipulation tools
        self.tools["insert_text_live"] = {
            "description": "Insert text into the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "Text to insert"
                    },
                    "position": {
                        "type": "integer",
                        "description": "Position to insert at (optional, defaults to cursor position)"
                    }
                },
                "required": ["text"]
            },
            "handler": self.insert_text_live
        }
        
        # Document info tools
        self.tools["get_document_info_live"] = {
            "description": "Get information about the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_document_info_live
        }
        
        # Text formatting tools
        self.tools["format_text_live"] = {
            "description": "Apply formatting to selected text in active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "bold": {
                        "type": "boolean",
                        "description": "Apply bold formatting"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Apply italic formatting"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Apply underline formatting"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name"
                    }
                }
            },
            "handler": self.format_text_live
        }
        
        # Paragraph style tools
        self.tools["format_paragraph_live"] = {
            "description": "Apply a paragraph style (e.g., Heading 1) to selected text or a specific paragraph",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {
                        "type": "string",
                        "description": "Name of the paragraph style to apply (e.g. 'Heading 1', 'Heading 2', 'Title', 'Text body')"
                    },
                    "paragraph_n": {
                        "type": "integer",
                        "description": "Optional paragraph number (1-indexed) to target directly instead of current selection"
                    }
                },
                "required": ["style_name"]
            },
            "handler": self.format_paragraph_live
        }
        
        self.tools["list_paragraph_styles_live"] = {
            "description": "List all available paragraph styles in the active document (built-in and user-defined)",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_paragraph_styles_live
        }
        
        self.tools["get_paragraph_style_live"] = {
            "description": "Get the paragraph style name of a specific paragraph or the current selection",
            "parameters": {
                "type": "object",
                "properties": {
                    "paragraph_n": {
                        "type": "integer",
                        "description": "Optional paragraph number (1-indexed) to query. If not provided, queries the current selection."
                    }
                }
            },
            "handler": self.get_paragraph_style_live
        }
        
        self.tools["create_paragraph_style_live"] = {
            "description": "Create a new paragraph style in the active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {
                        "type": "string",
                        "description": "Name for the new style (must not already exist)"
                    },
                    "parent_style": {
                        "type": "string",
                        "description": "Parent style to inherit from (default: 'Standard')",
                        "default": "Standard"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name (optional)"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points (optional)"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Bold weight (optional)"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Italic posture (optional)"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Single underline (optional)"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "right", "center", "justify"],
                        "description": "Paragraph alignment (optional)"
                    }
                },
                "required": ["style_name"]
            },
            "handler": self.create_paragraph_style_live
        }
        
        self.tools["edit_paragraph_style_live"] = {
            "description": "Modify an existing paragraph style's properties",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {
                        "type": "string",
                        "description": "Name of the existing style to modify"
                    },
                    "parent_style": {
                        "type": "string",
                        "description": "New parent style (optional)"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name (optional)"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points (optional)"
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Bold weight (optional)"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Italic posture (optional)"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Single underline (optional)"
                    },
                    "alignment": {
                        "type": "string",
                        "enum": ["left", "right", "center", "justify"],
                        "description": "Paragraph alignment (optional)"
                    }
                },
                "required": ["style_name"]
            },
            "handler": self.edit_paragraph_style_live
        }
        
        self.tools["delete_paragraph_style_live"] = {
            "description": "Delete a paragraph style from the active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {
                        "type": "string",
                        "description": "Name of the style to delete"
                    }
                },
                "required": ["style_name"]
            },
            "handler": self.delete_paragraph_style_live
        }
        
        self.tools["get_style_properties_live"] = {
            "description": "Get all properties (font, size, bold, italic, underline, alignment, parent) of a named paragraph style",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {
                        "type": "string",
                        "description": "Name of the paragraph style to inspect"
                    }
                },
                "required": ["style_name"]
            },
            "handler": self.get_style_properties_live
        }
        
        # Document saving tools
        self.tools["save_document_live"] = {
            "description": "Save the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to save document to (optional, saves to current location if not specified)"
                    }
                }
            },
            "handler": self.save_document_live
        }
        
        # Document export tools
        self.tools["export_document_live"] = {
            "description": "Export the currently active document to a different format",
            "parameters": {
                "type": "object",
                "properties": {
                    "export_format": {
                        "type": "string",
                        "enum": ["pdf", "docx", "doc", "odt", "txt", "rtf", "html"],
                        "description": "Format to export to"
                    },
                    "file_path": {
                        "type": "string",
                        "description": "Path to export document to"
                    }
                },
                "required": ["export_format", "file_path"]
            },
            "handler": self.export_document_live
        }
        
        # Content reading tools
        self.tools["get_text_content_live"] = {
            "description": "Get the text content of the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_text_content_live
        }
        
        # Document list tools
        self.tools["list_open_documents"] = {
            "description": "List all currently open documents in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_open_documents
        }

        # Comment tools
        self.tools["get_comments_live"] = {
            "description": "Get all comments/annotations from the document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_comments_live
        }

        self.tools["add_comment_live"] = {
            "description": "Add a comment at the current cursor position",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "Comment text"
                    },
                    "author": {
                        "type": "string",
                        "description": "Comment author name",
                        "default": "Claude"
                    }
                },
                "required": ["text"]
            },
            "handler": self.add_comment_live
        }

        # Enhanced Editing Tools - Document Structure
        self.tools["get_paragraph_count_live"] = {
            "description": "Get the total number of paragraphs in the document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_paragraph_count_live
        }

        self.tools["get_document_outline_live"] = {
            "description": "Get document outline with headings, paragraph numbers, and levels",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_document_outline_live
        }

        self.tools["get_paragraph_live"] = {
            "description": "Get content of a specific paragraph by number (1-indexed)",
            "parameters": {
                "type": "object",
                "properties": {
                    "n": {
                        "type": "integer",
                        "description": "Paragraph number (1-indexed)"
                    }
                },
                "required": ["n"]
            },
            "handler": self.get_paragraph_live
        }

        self.tools["get_paragraphs_range_live"] = {
            "description": "Get content of paragraphs in a range (inclusive, 1-indexed)",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {
                        "type": "integer",
                        "description": "Starting paragraph number (1-indexed)"
                    },
                    "end": {
                        "type": "integer",
                        "description": "Ending paragraph number (inclusive)"
                    }
                },
                "required": ["start", "end"]
            },
            "handler": self.get_paragraphs_range_live
        }

        # Enhanced Editing Tools - Cursor Navigation
        self.tools["goto_paragraph_live"] = {
            "description": "Move view cursor to the beginning of paragraph n",
            "parameters": {
                "type": "object",
                "properties": {
                    "n": {
                        "type": "integer",
                        "description": "Paragraph number (1-indexed)"
                    }
                },
                "required": ["n"]
            },
            "handler": self.goto_paragraph_live
        }

        self.tools["goto_position_live"] = {
            "description": "Move view cursor to a specific character position",
            "parameters": {
                "type": "object",
                "properties": {
                    "char_pos": {
                        "type": "integer",
                        "description": "Character position (0-indexed)"
                    }
                },
                "required": ["char_pos"]
            },
            "handler": self.goto_position_live
        }

        self.tools["get_cursor_position_live"] = {
            "description": "Get current cursor character position and paragraph number",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_cursor_position_live
        }

        self.tools["get_context_around_cursor_live"] = {
            "description": "Get text context around the current cursor position",
            "parameters": {
                "type": "object",
                "properties": {
                    "chars": {
                        "type": "integer",
                        "description": "Number of characters to get before and after cursor (default: 100)",
                        "default": 100
                    }
                }
            },
            "handler": self.get_context_around_cursor_live
        }

        # Enhanced Editing Tools - Text Selection
        self.tools["select_paragraph_live"] = {
            "description": "Select entire paragraph n (1-indexed)",
            "parameters": {
                "type": "object",
                "properties": {
                    "n": {
                        "type": "integer",
                        "description": "Paragraph number (1-indexed)"
                    }
                },
                "required": ["n"]
            },
            "handler": self.select_paragraph_live
        }

        self.tools["select_text_range_live"] = {
            "description": "Select text from start to end character positions (0-indexed)",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {
                        "type": "integer",
                        "description": "Starting character position (0-indexed)"
                    },
                    "end": {
                        "type": "integer",
                        "description": "Ending character position (exclusive)"
                    }
                },
                "required": ["start", "end"]
            },
            "handler": self.select_text_range_live
        }

        self.tools["delete_selection_live"] = {
            "description": "Delete currently selected text",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.delete_selection_live
        }

        self.tools["replace_selection_live"] = {
            "description": "Replace currently selected text with new text",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "New text to replace selection with"
                    }
                },
                "required": ["text"]
            },
            "handler": self.replace_selection_live
        }

        # Enhanced Editing Tools - Search and Replace
        self.tools["find_text_live"] = {
            "description": "Find all occurrences of query string in the document",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "String to search for"
                    }
                },
                "required": ["query"]
            },
            "handler": self.find_text_live
        }

        self.tools["find_and_replace_live"] = {
            "description": "Find and replace the first occurrence",
            "parameters": {
                "type": "object",
                "properties": {
                    "old": {
                        "type": "string",
                        "description": "String to find"
                    },
                    "new": {
                        "type": "string",
                        "description": "String to replace with"
                    }
                },
                "required": ["old", "new"]
            },
            "handler": self.find_and_replace_live
        }

        self.tools["find_and_replace_all_live"] = {
            "description": "Find and replace all occurrences",
            "parameters": {
                "type": "object",
                "properties": {
                    "old": {
                        "type": "string",
                        "description": "String to find"
                    },
                    "new": {
                        "type": "string",
                        "description": "String to replace with"
                    }
                },
                "required": ["old", "new"]
            },
            "handler": self.find_and_replace_all_live
        }

        # Track Changes tools
        self.tools["get_track_changes_status_live"] = {
            "description": "Get Track Changes recording and display status",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_track_changes_status_live
        }

        self.tools["set_track_changes_live"] = {
            "description": "Enable or disable Track Changes recording and display",
            "parameters": {
                "type": "object",
                "properties": {
                    "enabled": {
                        "type": "boolean",
                        "description": "Enable or disable Track Changes recording"
                    },
                    "show": {
                        "type": "boolean",
                        "description": "Show or hide tracked changes",
                        "default": True
                    }
                },
                "required": ["enabled"]
            },
            "handler": self.set_track_changes_live
        }

        self.tools["get_tracked_changes_live"] = {
            "description": "Get list of all tracked changes in the document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_tracked_changes_live
        }

        self.tools["accept_tracked_change_live"] = {
            "description": "Accept a specific tracked change by index",
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {
                        "type": "integer",
                        "description": "Index of the tracked change to accept"
                    }
                },
                "required": ["index"]
            },
            "handler": self.accept_tracked_change_live
        }

        self.tools["reject_tracked_change_live"] = {
            "description": "Reject a specific tracked change by index",
            "parameters": {
                "type": "object",
                "properties": {
                    "index": {
                        "type": "integer",
                        "description": "Index of the tracked change to reject"
                    }
                },
                "required": ["index"]
            },
            "handler": self.reject_tracked_change_live
        }

        self.tools["accept_all_changes_live"] = {
            "description": "Accept all tracked changes in the document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.accept_all_changes_live
        }

        self.tools["reject_all_changes_live"] = {
            "description": "Reject all tracked changes in the document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.reject_all_changes_live
        }

        # ── Calc / Spreadsheet Tools ──

        self.tools["get_cell_value_live"] = {
            "description": "Get the value of a cell in a Calc spreadsheet (e.g., 'A1' or 'Sheet1.B5')",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell_address": {
                        "type": "string",
                        "description": "Cell address like 'A1', 'B5', or 'Sheet1.A1'"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name (optional; uses active sheet or address prefix if omitted)"
                    }
                },
                "required": ["cell_address"]
            },
            "handler": self.get_cell_value_live
        }

        self.tools["set_cell_value_live"] = {
            "description": "Set the value of a cell in a Calc spreadsheet",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell_address": {
                        "type": "string",
                        "description": "Cell address like 'A1' or 'Sheet1.B5'"
                    },
                    "value": {
                        "type": "string",
                        "description": "Value to set (number or text)"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name (optional)"
                    }
                },
                "required": ["cell_address", "value"]
            },
            "handler": self.set_cell_value_live
        }

        self.tools["get_cell_range_live"] = {
            "description": "Get a rectangular range of cells as a 2D array (e.g., 'A1:C10' or 'Sheet1.A1:C10')",
            "parameters": {
                "type": "object",
                "properties": {
                    "range_address": {
                        "type": "string",
                        "description": "Range like 'A1:B10' or 'Sheet1.A1:B10'"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Sheet name (optional)"
                    }
                },
                "required": ["range_address"]
            },
            "handler": self.get_cell_range_live
        }

        self.tools["list_sheets_live"] = {
            "description": "List all sheet names in the current Calc document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_sheets_live
        }

        logger.info(f"Registered {len(self.tools)} MCP tools")
    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute an MCP tool
        
        Args:
            tool_name: Name of the tool to execute
            parameters: Parameters for the tool
            
        Returns:
            Result dictionary
        """
        try:
            if tool_name not in self.tools:
                return {
                    "success": False,
                    "error": f"Unknown tool: {tool_name}",
                    "available_tools": list(self.tools.keys())
                }
            
            tool = self.tools[tool_name]
            handler = tool["handler"]
            
            # Execute the tool handler
            result = handler(**parameters)
            
            logger.info(f"Executed tool '{tool_name}' successfully")
            return result
            
        except Exception as e:
            logger.error(f"Error executing tool '{tool_name}': {e}")
            return {
                "success": False,
                "error": str(e),
                "tool": tool_name,
                "parameters": parameters
            }
    
    def get_tool_list(self) -> List[Dict[str, Any]]:
        """Get list of available tools with their descriptions"""
        return [
            {
                "name": name,
                "description": tool["description"],
                "parameters": tool["parameters"]
            }
            for name, tool in self.tools.items()
        ]
    
    # Tool handler methods
    
    def create_document_live(self, doc_type: str = "writer") -> Dict[str, Any]:
        """Create a new document in LibreOffice"""
        try:
            doc = self.uno_bridge.create_document(doc_type)
            doc_info = self.uno_bridge.get_document_info(doc)
            
            return {
                "success": True,
                "message": f"Created new {doc_type} document",
                "document_info": doc_info
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def insert_text_live(self, text: str, position: Optional[int] = None) -> Dict[str, Any]:
        """Insert text into the currently active document"""
        return self.uno_bridge.insert_text(text, position)
    
    def get_document_info_live(self) -> Dict[str, Any]:
        """Get information about the currently active document"""
        doc_info = self.uno_bridge.get_document_info()
        if "error" in doc_info:
            return {"success": False, **doc_info}
        else:
            return {"success": True, "document_info": doc_info}
    
    def format_text_live(self, **formatting) -> Dict[str, Any]:
        """Apply formatting to selected text"""
        return self.uno_bridge.format_text(formatting)
    
    def format_paragraph_live(self, style_name: str, paragraph_n: Optional[int] = None) -> Dict[str, Any]:
        """Apply a paragraph style to selected text or a specific paragraph"""
        return self.uno_bridge.set_paragraph_style(style_name, paragraph_n)
    
    def list_paragraph_styles_live(self) -> Dict[str, Any]:
        """List all available paragraph styles"""
        return self.uno_bridge.list_paragraph_styles()
    
    def get_paragraph_style_live(self, paragraph_n: Optional[int] = None) -> Dict[str, Any]:
        """Get the paragraph style of a specific paragraph or current selection"""
        return self.uno_bridge.get_paragraph_style(paragraph_n)
    
    def create_paragraph_style_live(self, style_name: str, parent_style: str = "Standard",
                                    font_name: Optional[str] = None,
                                    font_size: Optional[float] = None,
                                    bold: Optional[bool] = None,
                                    italic: Optional[bool] = None,
                                    underline: Optional[bool] = None,
                                    alignment: Optional[str] = None) -> Dict[str, Any]:
        """Create a new paragraph style"""
        return self.uno_bridge.create_paragraph_style(
            style_name=style_name, parent_style=parent_style,
            font_name=font_name, font_size=font_size,
            bold=bold, italic=italic, underline=underline,
            alignment=alignment)
    
    def edit_paragraph_style_live(self, style_name: str, parent_style: Optional[str] = None,
                                  font_name: Optional[str] = None,
                                  font_size: Optional[float] = None,
                                  bold: Optional[bool] = None,
                                  italic: Optional[bool] = None,
                                  underline: Optional[bool] = None,
                                  alignment: Optional[str] = None) -> Dict[str, Any]:
        """Modify an existing paragraph style"""
        return self.uno_bridge.edit_paragraph_style(
            style_name=style_name, parent_style=parent_style,
            font_name=font_name, font_size=font_size,
            bold=bold, italic=italic, underline=underline,
            alignment=alignment)
    
    def delete_paragraph_style_live(self, style_name: str) -> Dict[str, Any]:
        """Delete a user-defined paragraph style"""
        return self.uno_bridge.delete_paragraph_style(style_name)
    
    def get_style_properties_live(self, style_name: str) -> Dict[str, Any]:
        """Get all properties of a named paragraph style"""
        return self.uno_bridge.get_style_properties(style_name)
    
    def save_document_live(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the currently active document"""
        return self.uno_bridge.save_document(file_path=file_path)
    
    def export_document_live(self, export_format: str, file_path: str) -> Dict[str, Any]:
        """Export the currently active document"""
        return self.uno_bridge.export_document(export_format, file_path)
    
    def get_text_content_live(self) -> Dict[str, Any]:
        """Get text content of the currently active document"""
        return self.uno_bridge.get_text_content()
    
    def list_open_documents(self) -> Dict[str, Any]:
        """List all open documents in LibreOffice"""
        try:
            desktop = self.uno_bridge.desktop
            documents = []
            
            # Get all open documents
            frames = desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController()
                if controller:
                    doc = controller.getModel()
                    if doc:
                        doc_info = self.uno_bridge.get_document_info(doc)
                        documents.append(doc_info)
            
            return {
                "success": True,
                "documents": documents,
                "count": len(documents)
            }

        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_comments_live(self) -> Dict[str, Any]:
        """Get all comments/annotations from the document"""
        return self.uno_bridge.get_comments()

    def add_comment_live(self, text: str, author: str = "Claude") -> Dict[str, Any]:
        """Add a comment at the current cursor position"""
        return self.uno_bridge.add_comment(text, author)

    # Enhanced Editing Tools - Document Structure Handlers

    def get_paragraph_count_live(self) -> Dict[str, Any]:
        """Get the total number of paragraphs in the document"""
        return self.uno_bridge.get_paragraph_count()

    def get_document_outline_live(self) -> Dict[str, Any]:
        """Get document outline with headings, paragraph numbers, and levels"""
        return self.uno_bridge.get_document_outline()

    def get_paragraph_live(self, n: int) -> Dict[str, Any]:
        """Get content of a specific paragraph by number (1-indexed)"""
        return self.uno_bridge.get_paragraph(n)

    def get_paragraphs_range_live(self, start: int, end: int) -> Dict[str, Any]:
        """Get content of paragraphs in a range (inclusive, 1-indexed)"""
        return self.uno_bridge.get_paragraphs_range(start, end)

    # Enhanced Editing Tools - Cursor Navigation Handlers

    def goto_paragraph_live(self, n: int) -> Dict[str, Any]:
        """Move view cursor to the beginning of paragraph n"""
        return self.uno_bridge.goto_paragraph(n)

    def goto_position_live(self, char_pos: int) -> Dict[str, Any]:
        """Move view cursor to a specific character position"""
        return self.uno_bridge.goto_position(char_pos)

    def get_cursor_position_live(self) -> Dict[str, Any]:
        """Get current cursor character position and paragraph number"""
        return self.uno_bridge.get_cursor_position()

    def get_context_around_cursor_live(self, chars: int = 100) -> Dict[str, Any]:
        """Get text context around the current cursor position"""
        return self.uno_bridge.get_context_around_cursor(chars)

    # Enhanced Editing Tools - Text Selection Handlers

    def select_paragraph_live(self, n: int) -> Dict[str, Any]:
        """Select entire paragraph n (1-indexed)"""
        return self.uno_bridge.select_paragraph(n)

    def select_text_range_live(self, start: int, end: int) -> Dict[str, Any]:
        """Select text from start to end character positions (0-indexed)"""
        return self.uno_bridge.select_text_range(start, end)

    def delete_selection_live(self) -> Dict[str, Any]:
        """Delete currently selected text"""
        return self.uno_bridge.delete_selection()

    def replace_selection_live(self, text: str) -> Dict[str, Any]:
        """Replace currently selected text with new text"""
        return self.uno_bridge.replace_selection(text)

    # Enhanced Editing Tools - Search and Replace Handlers

    def find_text_live(self, query: str) -> Dict[str, Any]:
        """Find all occurrences of query string in the document"""
        return self.uno_bridge.find_text(query)

    def find_and_replace_live(self, old: str, new: str) -> Dict[str, Any]:
        """Find and replace the first occurrence"""
        return self.uno_bridge.find_and_replace(old, new)

    def find_and_replace_all_live(self, old: str, new: str) -> Dict[str, Any]:
        """Find and replace all occurrences"""
        return self.uno_bridge.find_and_replace_all(old, new)

    # Track Changes Handlers

    def get_track_changes_status_live(self) -> Dict[str, Any]:
        """Get Track Changes recording and display status"""
        return self.uno_bridge.get_track_changes_status()

    def set_track_changes_live(self, enabled: bool, show: bool = True) -> Dict[str, Any]:
        """Enable or disable Track Changes recording and display"""
        return self.uno_bridge.set_track_changes(enabled, show)

    def get_tracked_changes_live(self) -> Dict[str, Any]:
        """Get list of all tracked changes in the document"""
        return self.uno_bridge.get_tracked_changes()

    def accept_tracked_change_live(self, index: int) -> Dict[str, Any]:
        """Accept a specific tracked change by index"""
        return self.uno_bridge.accept_tracked_change(index)

    def reject_tracked_change_live(self, index: int) -> Dict[str, Any]:
        """Reject a specific tracked change by index"""
        return self.uno_bridge.reject_tracked_change(index)

    def accept_all_changes_live(self) -> Dict[str, Any]:
        """Accept all tracked changes in the document"""
        return self.uno_bridge.accept_all_changes()

    def reject_all_changes_live(self) -> Dict[str, Any]:
        """Reject all tracked changes in the document"""
        return self.uno_bridge.reject_all_changes()

    # ── Calc / Spreadsheet Handlers ──

    def get_cell_value_live(self, cell_address: str, sheet_name: str = None) -> Dict[str, Any]:
        """Get the value of a cell in a Calc spreadsheet"""
        return self.uno_bridge.get_cell_value(sheet_name, cell_address)

    def set_cell_value_live(self, cell_address: str, value: str, sheet_name: str = None) -> Dict[str, Any]:
        """Set the value of a cell in a Calc spreadsheet"""
        # Try numeric conversion
        try:
            numeric = float(value)
            return self.uno_bridge.set_cell_value(sheet_name, cell_address, numeric)
        except ValueError:
            return self.uno_bridge.set_cell_value(sheet_name, cell_address, value)

    def get_cell_range_live(self, range_address: str, sheet_name: str = None) -> Dict[str, Any]:
        """Get a range of cells as a 2D array"""
        return self.uno_bridge.get_cell_range(sheet_name, range_address)

    def list_sheets_live(self) -> Dict[str, Any]:
        """List all sheet names"""
        return self.uno_bridge.list_sheets()


# Global instance
mcp_server = None

def get_mcp_server() -> LibreOfficeMCPServer:
    """Get or create the global MCP server instance"""
    global mcp_server
    if mcp_server is None:
        mcp_server = LibreOfficeMCPServer()
    return mcp_server
