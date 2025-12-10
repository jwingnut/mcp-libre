"""
LibreOffice MCP Extension - UNO Bridge Module

This module provides a bridge between MCP operations and LibreOffice UNO API,
enabling direct manipulation of LibreOffice documents.
"""

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from typing import Any, Optional, Dict, List
import logging
import traceback

# Optional imports - these may not be available in all configurations
try:
    from com.sun.star.text import XTextDocument
except ImportError:
    XTextDocument = None

try:
    from com.sun.star.sheet import XSpreadsheetDocument
except ImportError:
    XSpreadsheetDocument = None

try:
    from com.sun.star.presentation import XPresentationDocument
except ImportError:
    XPresentationDocument = None

try:
    from com.sun.star.document import XDocumentEventListener
except ImportError:
    XDocumentEventListener = None

try:
    from com.sun.star.awt import XActionListener
except ImportError:
    XActionListener = None

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _is_instance(obj, cls):
    """Safe isinstance check that handles None class types"""
    if cls is None:
        return False
    return isinstance(obj, cls)


class UNOBridge:
    """Bridge between MCP operations and LibreOffice UNO API"""
    
    def __init__(self):
        """Initialize the UNO bridge"""
        try:
            self.ctx = uno.getComponentContext()
            self.smgr = self.ctx.ServiceManager
            self.desktop = self.smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            logger.info("UNO Bridge initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize UNO Bridge: {e}")
            raise
    
    def create_document(self, doc_type: str = "writer") -> Any:
        """
        Create new document using UNO API
        
        Args:
            doc_type: Type of document ('writer', 'calc', 'impress', 'draw')
            
        Returns:
            Document object
        """
        try:
            url_map = {
                "writer": "private:factory/swriter",
                "calc": "private:factory/scalc", 
                "impress": "private:factory/simpress",
                "draw": "private:factory/sdraw"
            }
            
            url = url_map.get(doc_type, "private:factory/swriter")
            doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
            logger.info(f"Created new {doc_type} document")
            return doc
            
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            raise
    
    def get_active_document(self) -> Optional[Any]:
        """Get currently active document"""
        try:
            doc = self.desktop.getCurrentComponent()
            if doc:
                logger.info("Retrieved active document")
            return doc
        except Exception as e:
            logger.error(f"Failed to get active document: {e}")
            return None
    
    def get_document_info(self, doc: Any = None) -> Dict[str, Any]:
        """Get information about a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"error": "No document available"}
            
            info = {
                "title": getattr(doc, 'Title', 'Unknown') if hasattr(doc, 'Title') else "Unknown",
                "url": doc.getURL() if hasattr(doc, 'getURL') else "",
                "modified": doc.isModified() if hasattr(doc, 'isModified') else False,
                "type": self._get_document_type(doc),
                "has_selection": self._has_selection(doc)
            }
            
            # Add document-specific information
            if _is_instance(doc, XTextDocument):
                text = doc.getText()
                info["word_count"] = len(text.getString().split())
                info["character_count"] = len(text.getString())
            elif _is_instance(doc, XSpreadsheetDocument):
                sheets = doc.getSheets()
                info["sheet_count"] = sheets.getCount()
                info["sheet_names"] = [sheets.getByIndex(i).getName() 
                                     for i in range(sheets.getCount())]
            
            return info
            
        except Exception as e:
            logger.error(f"Failed to get document info: {e}")
            return {"error": str(e)}
    
    def insert_text(self, text: str, position: Optional[int] = None, doc: Any = None) -> Dict[str, Any]:
        """
        Insert text into a document
        
        Args:
            text: Text to insert
            position: Position to insert at (None for current cursor position)
            doc: Document to insert into (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No active document"}

            # Check if it's a Writer document
            is_writer = _is_instance(doc, XTextDocument) or \
                        (hasattr(doc, 'supportsService') and doc.supportsService("com.sun.star.text.TextDocument")) or \
                        hasattr(doc, 'getText')

            # Handle Writer documents
            if is_writer:
                text_obj = doc.getText()

                if position is None:
                    # Insert at current cursor position
                    cursor = doc.getCurrentController().getViewCursor()
                else:
                    # Insert at specific position
                    cursor = text_obj.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.goRight(position, False)

                text_obj.insertString(cursor, text, False)
                logger.info(f"Inserted {len(text)} characters into Writer document")
                return {"success": True, "message": f"Inserted {len(text)} characters"}

            # Handle other document types
            else:
                return {"success": False, "error": f"Text insertion not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to insert text: {e}")
            return {"success": False, "error": str(e)}
    
    def format_text(self, formatting: Dict[str, Any], doc: Any = None) -> Dict[str, Any]:
        """
        Apply formatting to selected text
        
        Args:
            formatting: Dictionary of formatting options
            doc: Document to format (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc or not _is_instance(doc, XTextDocument):
                return {"success": False, "error": "No Writer document available"}
            
            # Get current selection
            selection = doc.getCurrentController().getSelection()
            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}
            
            # Apply formatting to selection
            text_range = selection.getByIndex(0)
            
            # Apply various formatting options
            if "bold" in formatting:
                text_range.CharWeight = 150.0 if formatting["bold"] else 100.0
            
            if "italic" in formatting:
                text_range.CharPosture = 2 if formatting["italic"] else 0
            
            if "underline" in formatting:
                text_range.CharUnderline = 1 if formatting["underline"] else 0
            
            if "font_size" in formatting:
                text_range.CharHeight = formatting["font_size"]
            
            if "font_name" in formatting:
                text_range.CharFontName = formatting["font_name"]
            
            logger.info("Applied formatting to selected text")
            return {"success": True, "message": "Formatting applied successfully"}
            
        except Exception as e:
            logger.error(f"Failed to format text: {e}")
            return {"success": False, "error": str(e)}
    
    def save_document(self, doc: Any = None, file_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Save a document
        
        Args:
            doc: Document to save (None for active document)
            file_path: Path to save to (None to save to current location)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to save"}
            
            if file_path:
                # Save as new file
                url = uno.systemPathToFileUrl(file_path)
                doc.storeAsURL(url, ())
                logger.info(f"Saved document to {file_path}")
                return {"success": True, "message": f"Document saved to {file_path}"}
            else:
                # Save to current location
                if doc.hasLocation():
                    doc.store()
                    logger.info("Saved document to current location")
                    return {"success": True, "message": "Document saved"}
                else:
                    return {"success": False, "error": "Document has no location, specify file_path"}
                    
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return {"success": False, "error": str(e)}
    
    def export_document(self, export_format: str, file_path: str, doc: Any = None) -> Dict[str, Any]:
        """
        Export document to different format
        
        Args:
            export_format: Target format ('pdf', 'docx', 'odt', 'txt', etc.)
            file_path: Path to export to
            doc: Document to export (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to export"}
            
            # Filter map for different formats
            filter_map = {
                'pdf': 'writer_pdf_Export',
                'docx': 'MS Word 2007 XML',
                'doc': 'MS Word 97',
                'odt': 'writer8',
                'txt': 'Text',
                'rtf': 'Rich Text Format',
                'html': 'HTML (StarWriter)'
            }
            
            filter_name = filter_map.get(export_format.lower())
            if not filter_name:
                return {"success": False, "error": f"Unsupported export format: {export_format}"}
            
            # Prepare export properties
            properties = (
                PropertyValue("FilterName", 0, filter_name, 0),
                PropertyValue("Overwrite", 0, True, 0),
            )
            
            # Export document
            url = uno.systemPathToFileUrl(file_path)
            doc.storeToURL(url, properties)
            
            logger.info(f"Exported document to {file_path} as {export_format}")
            return {"success": True, "message": f"Document exported to {file_path}"}
            
        except Exception as e:
            logger.error(f"Failed to export document: {e}")
            return {"success": False, "error": str(e)}
    
    def get_text_content(self, doc: Any = None) -> Dict[str, Any]:
        """Get text content from a document"""
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            # Check if it's a Writer document
            is_writer = _is_instance(doc, XTextDocument) or \
                        (hasattr(doc, 'supportsService') and doc.supportsService("com.sun.star.text.TextDocument")) or \
                        hasattr(doc, 'getText')

            if is_writer:
                text = doc.getText().getString()
                return {"success": True, "content": text, "length": len(text)}
            else:
                return {"success": False, "error": f"Text extraction not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to get text content: {e}")
            return {"success": False, "error": str(e)}
    
    def get_comments(self, doc: Any = None) -> Dict[str, Any]:
        """Get all comments/annotations from the document"""
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            comments = []

            # Try to get text fields enumeration (comments are stored as text fields)
            if hasattr(doc, 'getTextFields'):
                text_fields = doc.getTextFields()
                enum = text_fields.createEnumeration()

                while enum.hasMoreElements():
                    field = enum.nextElement()
                    # Check if it's an annotation (comment)
                    if hasattr(field, 'supportsService') and field.supportsService("com.sun.star.text.TextField.Annotation"):
                        comment_data = {
                            "author": field.Author if hasattr(field, 'Author') else "",
                            "content": field.Content if hasattr(field, 'Content') else "",
                            "date": str(field.Date) if hasattr(field, 'Date') else "",
                        }
                        # Try to get the anchor text (what the comment is attached to)
                        if hasattr(field, 'getAnchor'):
                            anchor = field.getAnchor()
                            if hasattr(anchor, 'getString'):
                                comment_data["anchor_text"] = anchor.getString()[:100]  # First 100 chars
                        comments.append(comment_data)

            return {"success": True, "comments": comments, "count": len(comments)}

        except Exception as e:
            logger.error(f"Failed to get comments: {e}")
            return {"success": False, "error": str(e)}

    def add_comment(self, text: str, author: str = "Claude", doc: Any = None) -> Dict[str, Any]:
        """Add a comment at the current cursor position"""
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            # Get the current cursor position
            controller = doc.getCurrentController()
            cursor = controller.getViewCursor()

            # Create annotation field
            annotation = doc.createInstance("com.sun.star.text.TextField.Annotation")
            annotation.Content = text
            annotation.Author = author

            # Insert at cursor position
            text_obj = doc.getText()
            text_obj.insertTextContent(cursor, annotation, False)

            logger.info(f"Added comment by {author}: {text[:50]}...")
            return {"success": True, "message": f"Comment added by {author}"}

        except Exception as e:
            logger.error(f"Failed to add comment: {e}")
            return {"success": False, "error": str(e)}

    def _get_document_type(self, doc: Any) -> str:
        """Determine document type"""
        # Try isinstance first if types are available
        if _is_instance(doc, XTextDocument):
            return "writer"
        elif _is_instance(doc, XSpreadsheetDocument):
            return "calc"
        elif _is_instance(doc, XPresentationDocument):
            return "impress"

        # Fallback: check supportsService (works even if types not imported)
        if hasattr(doc, 'supportsService'):
            if doc.supportsService("com.sun.star.text.TextDocument"):
                return "writer"
            elif doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return "calc"
            elif doc.supportsService("com.sun.star.presentation.PresentationDocument"):
                return "impress"
            elif doc.supportsService("com.sun.star.drawing.DrawingDocument"):
                return "draw"

        # Fallback: check for getText method (Writer documents)
        if hasattr(doc, 'getText'):
            return "writer"

        return "unknown"
    
    def _has_selection(self, doc: Any) -> bool:
        """Check if document has selected content"""
        try:
            if hasattr(doc, 'getCurrentController'):
                controller = doc.getCurrentController()
                if hasattr(controller, 'getSelection'):
                    selection = controller.getSelection()
                    return selection.getCount() > 0
        except:
            pass
        return False
