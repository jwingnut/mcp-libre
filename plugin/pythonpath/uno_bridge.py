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

            doc_type = self._get_document_type(doc)

            info = {
                "title": getattr(doc, 'Title', 'Unknown') if hasattr(doc, 'Title') else "Unknown",
                "url": doc.getURL() if hasattr(doc, 'getURL') else "",
                "modified": doc.isModified() if hasattr(doc, 'isModified') else False,
                "type": doc_type,
                "has_selection": self._has_selection(doc)
            }

            # Add document-specific information
            if _is_instance(doc, XTextDocument):
                text = doc.getText()
                info["word_count"] = len(text.getString().split())
                info["character_count"] = len(text.getString())

                # Add track_changes status for Writer documents
                tc_status = self.get_track_changes_status(doc)
                if tc_status.get("success"):
                    info["track_changes"] = {
                        "recording": tc_status.get("recording", False),
                        "showing": tc_status.get("showing", False),
                        "pending_count": tc_status.get("pending_count", 0)
                    }
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

    # ============== Track Changes Tools ==============

    def get_track_changes_status(self, doc: Any = None) -> Dict[str, Any]:
        """
        Get Track Changes status for the document.

        Args:
            doc: Document to check (None for active document)

        Returns:
            Result dictionary with recording, showing, and pending_count
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            # Get RecordChanges and ShowChanges properties
            recording = False
            showing = False
            pending_count = 0

            # Access document properties via XPropertySet
            if hasattr(doc, 'getPropertyValue'):
                try:
                    recording = doc.getPropertyValue("RecordChanges")
                except:
                    pass
                try:
                    showing = doc.getPropertyValue("ShowChanges")
                except:
                    pass

            # Count pending redlines using XRedlinesSupplier
            if hasattr(doc, 'getRedlines'):
                try:
                    redlines = doc.getRedlines()
                    if redlines:
                        pending_count = redlines.getCount()
                except:
                    pass

            logger.info(f"Track Changes status: recording={recording}, showing={showing}, pending={pending_count}")
            return {
                "success": True,
                "recording": recording,
                "showing": showing,
                "pending_count": pending_count
            }

        except Exception as e:
            logger.error(f"Failed to get track changes status: {e}")
            return {"success": False, "error": str(e)}

    def set_track_changes(self, enabled: bool, show: bool = True, doc: Any = None) -> Dict[str, Any]:
        """
        Enable or disable Track Changes recording.

        Args:
            enabled: Whether to enable Track Changes recording
            show: Whether to show tracked changes (default: True)
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with new state
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            # Set properties via XPropertySet
            if hasattr(doc, 'setPropertyValue'):
                try:
                    doc.setPropertyValue("RecordChanges", enabled)
                except Exception as e:
                    return {"success": False, "error": f"Cannot set RecordChanges: {e}"}
                try:
                    doc.setPropertyValue("ShowChanges", show)
                except Exception as e:
                    return {"success": False, "error": f"Cannot set ShowChanges: {e}"}
            else:
                return {"success": False, "error": "Document does not support property modification"}

            logger.info(f"Set Track Changes: recording={enabled}, showing={show}")
            return {
                "success": True,
                "recording": enabled,
                "showing": show
            }

        except Exception as e:
            logger.error(f"Failed to set track changes: {e}")
            return {"success": False, "error": str(e)}

    def get_tracked_changes(self, doc: Any = None) -> Dict[str, Any]:
        """
        Get list of all tracked changes in the document.

        Args:
            doc: Document to check (None for active document)

        Returns:
            Result dictionary with list of changes
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            changes = []

            # Get redlines using XRedlinesSupplier
            if hasattr(doc, 'getRedlines'):
                redlines = doc.getRedlines()
                if redlines:
                    for i in range(redlines.getCount()):
                        try:
                            redline = redlines.getByIndex(i)

                            # Get redline properties
                            redline_type = ""
                            if hasattr(redline, 'RedlineType'):
                                redline_type = redline.RedlineType

                            text = ""
                            if hasattr(redline, 'getText'):
                                text_obj = redline.getText()
                                if text_obj and hasattr(text_obj, 'getString'):
                                    text = text_obj.getString()

                            author = ""
                            if hasattr(redline, 'RedlineAuthor'):
                                author = redline.RedlineAuthor

                            date_str = ""
                            if hasattr(redline, 'RedlineDateTime'):
                                dt = redline.RedlineDateTime
                                # Format as ISO string
                                date_str = f"{dt.Year:04d}-{dt.Month:02d}-{dt.Day:02d}T{dt.Hours:02d}:{dt.Minutes:02d}:{dt.Seconds:02d}"

                            description = ""
                            if hasattr(redline, 'RedlineComment'):
                                description = redline.RedlineComment

                            changes.append({
                                "index": i,
                                "type": redline_type.lower() if redline_type else "unknown",
                                "text": text[:500] if text else "",  # Limit text length
                                "author": author,
                                "date": date_str,
                                "description": description
                            })
                        except Exception as e:
                            logger.warning(f"Failed to read redline {i}: {e}")
                            continue

            logger.info(f"Found {len(changes)} tracked changes")
            return {
                "success": True,
                "changes": changes,
                "count": len(changes)
            }

        except Exception as e:
            logger.error(f"Failed to get tracked changes: {e}")
            return {"success": False, "error": str(e)}

    def accept_tracked_change(self, index: int, doc: Any = None) -> Dict[str, Any]:
        """
        Accept a specific tracked change by index.

        Args:
            index: Index of the change to accept (0-based)
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with accepted index
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            if not hasattr(doc, 'getRedlines'):
                return {"success": False, "error": "Document does not support redlines"}

            redlines = doc.getRedlines()
            if not redlines:
                return {"success": False, "error": "No tracked changes in document"}

            count = redlines.getCount()
            if index < 0 or index >= count:
                return {"success": False, "error": f"Index {index} out of range. Valid range: 0-{count-1}"}

            # Get the redline and accept it
            redline = redlines.getByIndex(index)

            # Accept by getting the text range and accepting via the document
            if hasattr(redline, 'getAnchor'):
                anchor = redline.getAnchor()
                if hasattr(anchor, 'getString'):
                    # Use the document's text to accept the redline
                    text = doc.getText()
                    if hasattr(text, 'createTextCursor'):
                        cursor = text.createTextCursorByRange(anchor)
                        # Accept redline - in UNO API, accepting means the change becomes permanent
                        if hasattr(doc, 'acceptRedline'):
                            doc.acceptRedline(index)
                        else:
                            # Alternative: use dispatcher
                            return {"success": False, "error": "Document does not support acceptRedline method"}

            logger.info(f"Accepted tracked change at index {index}")
            return {
                "success": True,
                "accepted_index": index
            }

        except Exception as e:
            logger.error(f"Failed to accept tracked change: {e}")
            return {"success": False, "error": str(e)}

    def reject_tracked_change(self, index: int, doc: Any = None) -> Dict[str, Any]:
        """
        Reject a specific tracked change by index.

        Args:
            index: Index of the change to reject (0-based)
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with rejected index
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            if not hasattr(doc, 'getRedlines'):
                return {"success": False, "error": "Document does not support redlines"}

            redlines = doc.getRedlines()
            if not redlines:
                return {"success": False, "error": "No tracked changes in document"}

            count = redlines.getCount()
            if index < 0 or index >= count:
                return {"success": False, "error": f"Index {index} out of range. Valid range: 0-{count-1}"}

            # Reject the redline
            if hasattr(doc, 'rejectRedline'):
                doc.rejectRedline(index)
            else:
                return {"success": False, "error": "Document does not support rejectRedline method"}

            logger.info(f"Rejected tracked change at index {index}")
            return {
                "success": True,
                "rejected_index": index
            }

        except Exception as e:
            logger.error(f"Failed to reject tracked change: {e}")
            return {"success": False, "error": str(e)}

    def accept_all_changes(self, doc: Any = None) -> Dict[str, Any]:
        """
        Accept all tracked changes in the document.

        Args:
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with count of accepted changes
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            if not hasattr(doc, 'getRedlines'):
                return {"success": False, "error": "Document does not support redlines"}

            redlines = doc.getRedlines()
            if not redlines:
                return {"success": True, "accepted_count": 0}

            count = redlines.getCount()
            if count == 0:
                return {"success": True, "accepted_count": 0}

            # Accept in reverse order to avoid index shifting
            accepted = 0
            for i in range(count - 1, -1, -1):
                try:
                    if hasattr(doc, 'acceptRedline'):
                        doc.acceptRedline(i)
                        accepted += 1
                except Exception as e:
                    logger.warning(f"Failed to accept redline {i}: {e}")

            logger.info(f"Accepted {accepted} tracked changes")
            return {
                "success": True,
                "accepted_count": accepted
            }

        except Exception as e:
            logger.error(f"Failed to accept all changes: {e}")
            return {"success": False, "error": str(e)}

    def reject_all_changes(self, doc: Any = None) -> Dict[str, Any]:
        """
        Reject all tracked changes in the document.

        Args:
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with count of rejected changes
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Track Changes not supported for {doc_type} documents"}

            if not hasattr(doc, 'getRedlines'):
                return {"success": False, "error": "Document does not support redlines"}

            redlines = doc.getRedlines()
            if not redlines:
                return {"success": True, "rejected_count": 0}

            count = redlines.getCount()
            if count == 0:
                return {"success": True, "rejected_count": 0}

            # Reject in reverse order to avoid index shifting
            rejected = 0
            for i in range(count - 1, -1, -1):
                try:
                    if hasattr(doc, 'rejectRedline'):
                        doc.rejectRedline(i)
                        rejected += 1
                except Exception as e:
                    logger.warning(f"Failed to reject redline {i}: {e}")

            logger.info(f"Rejected {rejected} tracked changes")
            return {
                "success": True,
                "rejected_count": rejected
            }

        except Exception as e:
            logger.error(f"Failed to reject all changes: {e}")
            return {"success": False, "error": str(e)}

    def _is_in_tracked_deletion(self, text_range: Any, doc: Any = None) -> bool:
        """
        Check if a text range is within a tracked deletion.

        Args:
            text_range: The text range to check
            doc: Document to check (None for active document)

        Returns:
            True if range is in a tracked deletion, False otherwise
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc or not hasattr(doc, 'getRedlines'):
                return False

            redlines = doc.getRedlines()
            if not redlines:
                return False

            text = doc.getText()

            for i in range(redlines.getCount()):
                try:
                    redline = redlines.getByIndex(i)

                    # Only check deletion redlines
                    if hasattr(redline, 'RedlineType'):
                        redline_type = redline.RedlineType
                        if redline_type and redline_type.lower() == "delete":
                            # Get redline anchor/range
                            if hasattr(redline, 'getAnchor'):
                                redline_range = redline.getAnchor()

                                # Compare ranges
                                # Check if text_range start is within redline range
                                try:
                                    start_compare = text.compareRegionStarts(text_range, redline_range)
                                    end_compare = text.compareRegionEnds(text_range, redline_range)

                                    # If text_range is fully contained within redline_range
                                    # start_compare >= 0 means text_range starts at or after redline start
                                    # end_compare <= 0 means text_range ends at or before redline end
                                    if start_compare >= 0 and end_compare <= 0:
                                        return True
                                except:
                                    pass
                except:
                    continue

            return False

        except Exception as e:
            logger.warning(f"Error checking tracked deletion: {e}")
            return False

    # ============== Enhanced Editing Tools ==============

    def get_paragraph_count(self, doc: Any = None) -> Dict[str, Any]:
        """
        Get the total number of paragraphs in the document.

        Args:
            doc: Document to analyze (None for active document)

        Returns:
            Result dictionary with paragraph count
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            # Check if it's a Writer document
            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Paragraph count not supported for {doc_type} documents"}

            # Get text and enumerate paragraphs
            text = doc.getText()
            enum = text.createEnumeration()

            count = 0
            while enum.hasMoreElements():
                para = enum.nextElement()
                # Check if it's a paragraph (not a table or other content)
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    count += 1

            logger.info(f"Document has {count} paragraphs")
            return {"success": True, "count": count}

        except Exception as e:
            logger.error(f"Failed to get paragraph count: {e}")
            return {"success": False, "error": str(e)}

    def get_document_outline(self, doc: Any = None) -> Dict[str, Any]:
        """
        Get document outline (headings) with paragraph numbers and levels.

        Args:
            doc: Document to analyze (None for active document)

        Returns:
            Result dictionary with outline and paragraph count
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Document outline not supported for {doc_type} documents"}

            text = doc.getText()
            enum = text.createEnumeration()

            outline = []
            paragraph_count = 0

            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    paragraph_count += 1

                    # Check if paragraph has a heading style
                    if hasattr(para, 'ParaStyleName'):
                        style_name = para.ParaStyleName
                        # Check for Heading 1-6 styles
                        if style_name and style_name.startswith("Heading"):
                            try:
                                level = int(style_name.replace("Heading ", "").replace("Heading", "1"))
                            except ValueError:
                                level = 1

                            # Get paragraph text
                            para_text = para.getString() if hasattr(para, 'getString') else ""

                            outline.append({
                                "paragraph": paragraph_count,
                                "level": level,
                                "text": para_text[:200]  # Limit text length
                            })

            logger.info(f"Document outline: {len(outline)} headings, {paragraph_count} paragraphs")
            return {
                "success": True,
                "outline": outline,
                "heading_count": len(outline),
                "paragraph_count": paragraph_count
            }

        except Exception as e:
            logger.error(f"Failed to get document outline: {e}")
            return {"success": False, "error": str(e)}

    def get_paragraph(self, n: int, doc: Any = None) -> Dict[str, Any]:
        """
        Get the content of a specific paragraph by number (1-indexed).

        Args:
            n: Paragraph number (1-indexed)
            doc: Document to read from (None for active document)

        Returns:
            Result dictionary with paragraph content
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Paragraph access not supported for {doc_type} documents"}

            if n < 1:
                return {"success": False, "error": "Paragraph number must be >= 1"}

            text = doc.getText()
            enum = text.createEnumeration()

            current = 0
            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    current += 1
                    if current == n:
                        content = para.getString() if hasattr(para, 'getString') else ""

                        # Build result with original content
                        result = {
                            "success": True,
                            "paragraph_number": n,
                            "content": content
                        }

                        # Add visible_content if Track Changes is enabled
                        tc_status = self.get_track_changes_status(doc)
                        if tc_status.get("success") and tc_status.get("recording"):
                            # Filter out tracked deletions
                            visible_content = self._filter_tracked_deletions(para, doc)
                            result["visible_content"] = visible_content

                        logger.info(f"Retrieved paragraph {n}")
                        return result

            # Paragraph not found
            return {
                "success": False,
                "error": f"Paragraph {n} out of range. Valid range: 1-{current}"
            }

        except Exception as e:
            logger.error(f"Failed to get paragraph: {e}")
            return {"success": False, "error": str(e)}

    def _filter_tracked_deletions(self, para: Any, doc: Any) -> str:
        """
        Filter out tracked deletions from paragraph content.

        Args:
            para: Paragraph text element
            doc: Document containing the paragraph

        Returns:
            String with tracked deletions filtered out
        """
        try:
            if not hasattr(doc, 'getRedlines'):
                return para.getString() if hasattr(para, 'getString') else ""

            redlines = doc.getRedlines()
            if not redlines or redlines.getCount() == 0:
                return para.getString() if hasattr(para, 'getString') else ""

            # Get paragraph range
            para_start = para.getStart()
            para_end = para.getEnd()
            text = doc.getText()

            # Collect all deletion ranges within this paragraph
            deletion_ranges = []
            for i in range(redlines.getCount()):
                try:
                    redline = redlines.getByIndex(i)

                    # Only check deletion redlines
                    if hasattr(redline, 'RedlineType'):
                        redline_type = redline.RedlineType
                        if redline_type and redline_type.lower() == "delete":
                            if hasattr(redline, 'getAnchor'):
                                redline_range = redline.getAnchor()

                                # Check if deletion overlaps with this paragraph
                                try:
                                    # Use compareRegionStarts/Ends to check overlap
                                    # If deletion is within paragraph, add to list
                                    deletion_ranges.append(redline_range)
                                except:
                                    pass
                except:
                    continue

            # If no deletions, return original text
            if not deletion_ranges:
                return para.getString() if hasattr(para, 'getString') else ""

            # Build visible content by iterating through paragraph portions
            visible_text = []
            if hasattr(para, 'createEnumeration'):
                portion_enum = para.createEnumeration()
                while portion_enum.hasMoreElements():
                    portion = portion_enum.nextElement()

                    # Check if this portion is in a tracked deletion
                    is_deleted = False
                    for del_range in deletion_ranges:
                        try:
                            # Check if portion overlaps with deletion
                            if self._is_in_tracked_deletion(portion, doc):
                                is_deleted = True
                                break
                        except:
                            pass

                    # Add portion text if not deleted
                    if not is_deleted and hasattr(portion, 'getString'):
                        visible_text.append(portion.getString())
            else:
                # Fallback to full paragraph text if can't enumerate portions
                return para.getString() if hasattr(para, 'getString') else ""

            return ''.join(visible_text)

        except Exception as e:
            logger.warning(f"Failed to filter tracked deletions: {e}")
            # Fallback to original content
            return para.getString() if hasattr(para, 'getString') else ""

    def get_paragraphs_range(self, start: int, end: int, doc: Any = None) -> Dict[str, Any]:
        """
        Get content of paragraphs in a range (inclusive, 1-indexed).

        Args:
            start: Starting paragraph number (1-indexed)
            end: Ending paragraph number (inclusive)
            doc: Document to read from (None for active document)

        Returns:
            Result dictionary with paragraphs content
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Paragraph access not supported for {doc_type} documents"}

            if start < 1:
                return {"success": False, "error": "Start paragraph must be >= 1"}
            if end < start:
                return {"success": False, "error": "End paragraph must be >= start paragraph"}

            text = doc.getText()
            enum = text.createEnumeration()

            paragraphs = []
            current = 0
            total_paragraphs = 0

            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    current += 1
                    total_paragraphs = current

                    if start <= current <= end:
                        content = para.getString() if hasattr(para, 'getString') else ""
                        paragraphs.append({
                            "number": current,
                            "content": content
                        })

                    if current > end:
                        break

            if not paragraphs:
                return {
                    "success": False,
                    "error": f"Range {start}-{end} out of bounds. Document has {total_paragraphs} paragraphs"
                }

            logger.info(f"Retrieved paragraphs {start}-{end}")
            return {
                "success": True,
                "paragraphs": paragraphs,
                "count": len(paragraphs)
            }

        except Exception as e:
            logger.error(f"Failed to get paragraphs range: {e}")
            return {"success": False, "error": str(e)}

    # ============== Cursor Navigation Tools ==============

    def goto_paragraph(self, n: int, doc: Any = None) -> Dict[str, Any]:
        """
        Move the view cursor to the beginning of paragraph n.

        Args:
            n: Paragraph number (1-indexed)
            doc: Document to navigate (None for active document)

        Returns:
            Result dictionary with cursor position
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Cursor navigation not supported for {doc_type} documents"}

            if n < 1:
                return {"success": False, "error": "Paragraph number must be >= 1"}

            text = doc.getText()
            enum = text.createEnumeration()

            current = 0
            target_para = None
            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    current += 1
                    if current == n:
                        target_para = para
                        break

            if target_para is None:
                return {"success": False, "error": f"Paragraph {n} out of range. Valid range: 1-{current}"}

            # Get the view cursor and move it to the paragraph start
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()

            # Get paragraph start position
            para_start = target_para.getStart()
            view_cursor.gotoRange(para_start, False)

            logger.info(f"Moved cursor to paragraph {n}")
            return {
                "success": True,
                "message": f"Cursor moved to paragraph {n}",
                "paragraph": n
            }

        except Exception as e:
            logger.error(f"Failed to goto paragraph: {e}")
            return {"success": False, "error": str(e)}

    def goto_position(self, char_pos: int, doc: Any = None) -> Dict[str, Any]:
        """
        Move the view cursor to a specific character position.

        Args:
            char_pos: Character position (0-indexed)
            doc: Document to navigate (None for active document)

        Returns:
            Result dictionary with actual position reached
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Cursor navigation not supported for {doc_type} documents"}

            if char_pos < 0:
                return {"success": False, "error": "Character position must be >= 0"}

            text = doc.getText()
            text_cursor = text.createTextCursor()
            text_cursor.gotoStart(False)

            # Move to position (goRight returns False if it can't move that far)
            actual_moved = 0
            if char_pos > 0:
                moved = text_cursor.goRight(char_pos, False)
                # Count actual position
                text_cursor_check = text.createTextCursor()
                text_cursor_check.gotoStart(False)
                text_cursor_check.gotoRange(text_cursor, True)
                actual_moved = len(text_cursor_check.getString())

            # Move view cursor to this position
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            view_cursor.gotoRange(text_cursor, False)

            logger.info(f"Moved cursor to position {actual_moved}")
            return {
                "success": True,
                "message": f"Cursor moved to position {actual_moved}",
                "position": actual_moved,
                "requested_position": char_pos
            }

        except Exception as e:
            logger.error(f"Failed to goto position: {e}")
            return {"success": False, "error": str(e)}

    def get_cursor_position(self, doc: Any = None) -> Dict[str, Any]:
        """
        Get the current cursor character position and paragraph number.

        Args:
            doc: Document to check (None for active document)

        Returns:
            Result dictionary with position and paragraph info
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Cursor position not supported for {doc_type} documents"}

            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()

            # Get character position by measuring from start
            text = doc.getText()
            text_cursor = text.createTextCursor()
            text_cursor.gotoStart(False)
            text_cursor.gotoRange(view_cursor, True)
            char_position = len(text_cursor.getString())

            # Find paragraph number
            enum = text.createEnumeration()
            paragraph_num = 0
            char_count = 0

            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    paragraph_num += 1
                    para_text = para.getString() if hasattr(para, 'getString') else ""
                    char_count += len(para_text) + 1  # +1 for paragraph break

                    if char_count >= char_position:
                        break

            logger.info(f"Cursor at position {char_position}, paragraph {paragraph_num}")
            return {
                "success": True,
                "position": char_position,
                "paragraph": paragraph_num
            }

        except Exception as e:
            logger.error(f"Failed to get cursor position: {e}")
            return {"success": False, "error": str(e)}

    def get_context_around_cursor(self, chars: int = 100, doc: Any = None) -> Dict[str, Any]:
        """
        Get text context around the current cursor position.

        Args:
            chars: Number of characters to get before and after cursor
            doc: Document to read from (None for active document)

        Returns:
            Result dictionary with text before and after cursor
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Cursor context not supported for {doc_type} documents"}

            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()
            text = doc.getText()

            # Get text before cursor
            before_cursor = text.createTextCursor()
            before_cursor.gotoStart(False)
            before_cursor.gotoRange(view_cursor, True)
            full_before = before_cursor.getString()
            text_before = full_before[-chars:] if len(full_before) > chars else full_before

            # Get text after cursor
            after_cursor = text.createTextCursor()
            after_cursor.gotoRange(view_cursor, False)
            after_cursor.gotoEnd(True)
            full_after = after_cursor.getString()
            text_after = full_after[:chars] if len(full_after) > chars else full_after

            # Get current position
            char_position = len(full_before)

            logger.info(f"Got context around position {char_position}")
            return {
                "success": True,
                "before": text_before,
                "after": text_after,
                "position": char_position,
                "chars_requested": chars
            }

        except Exception as e:
            logger.error(f"Failed to get context around cursor: {e}")
            return {"success": False, "error": str(e)}

    # ============== Text Selection Tools ==============

    def select_paragraph(self, n: int, doc: Any = None) -> Dict[str, Any]:
        """
        Select entire paragraph n (1-indexed).

        Args:
            n: Paragraph number (1-indexed)
            doc: Document to work with (None for active document)

        Returns:
            Result dictionary with selected text content
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Paragraph selection not supported for {doc_type} documents"}

            if n < 1:
                return {"success": False, "error": "Paragraph number must be >= 1"}

            # Find the paragraph
            text = doc.getText()
            enum = text.createEnumeration()

            current = 0
            target_para = None
            while enum.hasMoreElements():
                para = enum.nextElement()
                if hasattr(para, 'supportsService') and para.supportsService("com.sun.star.text.Paragraph"):
                    current += 1
                    if current == n:
                        target_para = para
                        break

            if target_para is None:
                return {"success": False, "error": f"Paragraph {n} out of range. Valid range: 1-{current}"}

            # Get the view cursor and select the paragraph
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()

            # Move to paragraph start
            para_start = target_para.getStart()
            view_cursor.gotoRange(para_start, False)

            # Extend selection to paragraph end
            para_end = target_para.getEnd()
            view_cursor.gotoRange(para_end, True)

            # Get selected text
            selected_text = target_para.getString() if hasattr(target_para, 'getString') else ""

            logger.info(f"Selected paragraph {n}")
            return {
                "success": True,
                "selected_text": selected_text,
                "paragraph": n
            }

        except Exception as e:
            logger.error(f"Failed to select paragraph: {e}")
            return {"success": False, "error": str(e)}

    def select_text_range(self, start: int, end: int, doc: Any = None) -> Dict[str, Any]:
        """
        Select text from start to end character positions (0-indexed).

        Args:
            start: Starting character position (0-indexed)
            end: Ending character position (exclusive)
            doc: Document to work with (None for active document)

        Returns:
            Result dictionary with selected text
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Text range selection not supported for {doc_type} documents"}

            if start < 0:
                return {"success": False, "error": "Start position must be >= 0"}
            if end < start:
                return {"success": False, "error": "End position must be >= start position"}

            text = doc.getText()
            controller = doc.getCurrentController()
            view_cursor = controller.getViewCursor()

            # Create text cursor for selection
            text_cursor = text.createTextCursor()
            text_cursor.gotoStart(False)

            # Move to start position
            if start > 0:
                text_cursor.goRight(start, False)

            # Store start position
            start_range = text.createTextCursor()
            start_range.gotoRange(text_cursor, False)

            # Move to end position (selecting)
            length = end - start
            if length > 0:
                text_cursor.goRight(length, True)

            # Get selected text
            selected_text = text_cursor.getString()

            # Move view cursor to match selection
            view_cursor.gotoRange(start_range, False)
            view_cursor.gotoRange(text_cursor, True)

            logger.info(f"Selected text range {start}-{end}")
            return {
                "success": True,
                "selected_text": selected_text,
                "start": start,
                "end": end,
                "length": len(selected_text)
            }

        except Exception as e:
            logger.error(f"Failed to select text range: {e}")
            return {"success": False, "error": str(e)}

    def delete_selection(self, doc: Any = None) -> Dict[str, Any]:
        """
        Delete currently selected text.

        Args:
            doc: Document to work with (None for active document)

        Returns:
            Result dictionary with deleted text content
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Delete selection not supported for {doc_type} documents"}

            # Get current selection
            controller = doc.getCurrentController()
            selection = controller.getSelection()

            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}

            # Get the selected text range
            text_range = selection.getByIndex(0)

            # Get the text before deleting
            deleted_text = text_range.getString()

            # Delete by setting empty string
            text_range.setString("")

            logger.info(f"Deleted selection: {len(deleted_text)} characters")
            return {
                "success": True,
                "deleted_text": deleted_text,
                "length": len(deleted_text)
            }

        except Exception as e:
            logger.error(f"Failed to delete selection: {e}")
            return {"success": False, "error": str(e)}

    def replace_selection(self, text: str, doc: Any = None) -> Dict[str, Any]:
        """
        Replace currently selected text with new text.

        Args:
            text: New text to replace selection with
            doc: Document to work with (None for active document)

        Returns:
            Result dictionary with old and new text
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Replace selection not supported for {doc_type} documents"}

            # Get current selection
            controller = doc.getCurrentController()
            selection = controller.getSelection()

            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}

            # Get the selected text range
            text_range = selection.getByIndex(0)

            # Get the old text
            old_text = text_range.getString()

            # Replace with new text
            text_range.setString(text)

            logger.info(f"Replaced selection: {len(old_text)} -> {len(text)} characters")
            return {
                "success": True,
                "old_text": old_text,
                "new_text": text,
                "old_length": len(old_text),
                "new_length": len(text)
            }

        except Exception as e:
            logger.error(f"Failed to replace selection: {e}")
            return {"success": False, "error": str(e)}

    # ============== Search and Replace Tools ==============

    def find_text(self, query: str, doc: Any = None) -> Dict[str, Any]:
        """
        Find all occurrences of query string in the document.

        Args:
            query: String to search for
            doc: Document to search in (None for active document)

        Returns:
            Result dictionary with list of matches and their positions
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Text search not supported for {doc_type} documents"}

            # Check if Track Changes is enabled
            track_changes_active = False
            if hasattr(doc, 'getPropertyValue'):
                try:
                    recording = doc.getPropertyValue("RecordChanges")
                    showing = doc.getPropertyValue("ShowChanges")
                    track_changes_active = recording or showing
                except:
                    pass

            # Create search descriptor
            search = doc.createSearchDescriptor()
            search.SearchString = query

            # Find all occurrences
            found = doc.findAll(search)

            matches = []
            if found and found.getCount() > 0:
                text = doc.getText()

                for i in range(found.getCount()):
                    match_range = found.getByIndex(i)

                    # Filter out matches in tracked deletions when Track Changes is active
                    if track_changes_active and self._is_in_tracked_deletion(match_range, doc):
                        continue

                    # Calculate character position from start
                    text_cursor = text.createTextCursor()
                    text_cursor.gotoStart(False)
                    text_cursor.gotoRange(match_range.getStart(), True)
                    position = len(text_cursor.getString())

                    # Get matched text
                    matched_text = match_range.getString()

                    matches.append({
                        "position": position,
                        "text": matched_text
                    })

            logger.info(f"Found {len(matches)} occurrences of '{query}' (Track Changes: {track_changes_active})")
            return {
                "success": True,
                "matches": matches,
                "count": len(matches),
                "query": query,
                "track_changes_active": track_changes_active
            }

        except Exception as e:
            logger.error(f"Failed to find text: {e}")
            return {"success": False, "error": str(e)}

    def find_and_replace(self, old: str, new: str, doc: Any = None) -> Dict[str, Any]:
        """
        Find and replace the first occurrence of old with new.

        When Track Changes is enabled, only replaces visible text occurrences,
        skipping matches that are within tracked deletions.

        Args:
            old: String to find
            new: String to replace with
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with replacement status and position
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Find and replace not supported for {doc_type} documents"}

            # Check if Track Changes is enabled
            track_changes_status = self.get_track_changes_status(doc)
            track_changes_active = track_changes_status.get("success") and track_changes_status.get("recording", False)

            # Create search descriptor
            search = doc.createSearchDescriptor()
            search.SearchString = old

            # Find first occurrence
            found = doc.findFirst(search)

            # If Track Changes is active, skip matches in tracked deletions
            if track_changes_active and found:
                while found and self._is_in_tracked_deletion(found, doc):
                    # Continue searching for next match
                    found = doc.findNext(found.getEnd(), search)

            if found:
                # Calculate position before replacement
                text = doc.getText()
                text_cursor = text.createTextCursor()
                text_cursor.gotoStart(False)
                text_cursor.gotoRange(found.getStart(), True)
                position = len(text_cursor.getString())

                # Replace the text
                found.setString(new)

                logger.info(f"Replaced first occurrence of '{old}' with '{new}' at position {position}")
                return {
                    "success": True,
                    "replaced": True,
                    "position": position,
                    "old": old,
                    "new": new
                }
            else:
                logger.info(f"No occurrence of '{old}' found")
                return {
                    "success": True,
                    "replaced": False,
                    "old": old,
                    "new": new
                }

        except Exception as e:
            logger.error(f"Failed to find and replace: {e}")
            return {"success": False, "error": str(e)}

    def find_and_replace_all(self, old: str, new: str, doc: Any = None) -> Dict[str, Any]:
        """
        Find and replace all occurrences of old with new.

        Track Changes aware: When Track Changes is enabled, this method iterates
        through matches manually to skip replacements in tracked deletions.
        When Track Changes is disabled, it uses native replaceAll for performance.

        Args:
            old: String to find
            new: String to replace with
            doc: Document to modify (None for active document)

        Returns:
            Result dictionary with count of replacements
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc:
                return {"success": False, "error": "No document available"}

            doc_type = self._get_document_type(doc)
            if doc_type != "writer":
                return {"success": False, "error": f"Find and replace all not supported for {doc_type} documents"}

            # Check if Track Changes is enabled
            track_changes_active = False
            if hasattr(doc, 'getPropertyValue'):
                try:
                    track_changes_active = doc.getPropertyValue("RecordChanges")
                except:
                    pass

            # If Track Changes is disabled, use native replaceAll for performance
            if not track_changes_active:
                replace = doc.createReplaceDescriptor()
                replace.SearchString = old
                replace.ReplaceString = new
                count = doc.replaceAll(replace)

                logger.info(f"Replaced {count} occurrences of '{old}' with '{new}' (Track Changes disabled)")
                return {
                    "success": True,
                    "count": count,
                    "old": old,
                    "new": new,
                    "track_changes_active": False
                }

            # Track Changes is enabled - must iterate manually to skip tracked deletions
            # Native replaceAll ignores Track Changes, so we use findFirst/findNext
            search = doc.createSearchDescriptor()
            search.SearchString = old

            count = 0
            found = doc.findFirst(search)

            while found:
                # Check if this match is in a tracked deletion
                if not self._is_in_tracked_deletion(found, doc):
                    # Replace this visible occurrence
                    found.setString(new)
                    count += 1

                # Find next occurrence
                # Note: We need to recreate the search after replacement
                # to avoid issues with modified text ranges
                search = doc.createSearchDescriptor()
                search.SearchString = old
                found = doc.findNext(found.getEnd(), search)

            logger.info(f"Replaced {count} visible occurrences of '{old}' with '{new}' (Track Changes enabled)")
            return {
                "success": True,
                "count": count,
                "old": old,
                "new": new,
                "track_changes_active": True
            }

        except Exception as e:
            logger.error(f"Failed to find and replace all: {e}")
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
