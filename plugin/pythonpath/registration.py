"""
LibreOffice MCP Extension - Registration Module

This module handles the registration and lifecycle of the LibreOffice MCP extension.
"""

import uno
import unohelper
import logging
import threading
import traceback
import sys
import os
from com.sun.star.lang import XServiceInfo
from com.sun.star.frame import XDispatchProvider, XDispatch

# Add the pythonpath directory to sys.path for imports
_this_dir = os.path.dirname(__file__)
if _this_dir not in sys.path:
    sys.path.insert(0, _this_dir)

# ── File-based logging so we can see what happens inside LO ──
LOG_FILE = os.path.join(os.environ.get("TEMP", "/tmp"), "mcp_extension.log")
_file_handler = None
try:
    _file_handler = logging.FileHandler(LOG_FILE, mode="w")
    _file_handler.setLevel(logging.DEBUG)
    _file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logging.getLogger().addHandler(_file_handler)
    logging.getLogger().setLevel(logging.DEBUG)
except Exception:
    pass
logger = logging.getLogger("MCPExtension")
logger.info("=== Extension module loaded, log at %s ===", LOG_FILE)

# Implementation name and service name for the extension
IMPLEMENTATION_NAME = "org.mcp.libreoffice.MCPExtension"
SERVICE_NAMES = ("com.sun.star.frame.ProtocolHandler",)

# Global server state (shared across all handler instances)
_server_started = False
_ai_interface = None
_lock = threading.Lock()


def _start_server():
    """Start the MCP HTTP server"""
    global _server_started, _ai_interface

    with _lock:
        if _server_started:
            logger.info("Server already started")
            return

        try:
            logger.info("Starting MCP server...")

            # Import using absolute import (no relative imports in LO Python)
            import ai_interface

            _ai_interface = ai_interface.start_ai_interface(port=8765, host="localhost")
            _server_started = True
            logger.info("MCP server started successfully on http://localhost:8765")

        except Exception as e:
            logger.error(f"Failed to start server: {e}")
            logger.error(traceback.format_exc())


def _stop_server():
    """Stop the MCP HTTP server"""
    global _server_started, _ai_interface

    with _lock:
        if not _server_started:
            logger.info("Server not running")
            return

        try:
            logger.info("Stopping MCP server...")
            import ai_interface
            ai_interface.stop_ai_interface()
            _ai_interface = None
            _server_started = False
            logger.info("MCP server stopped")

        except Exception as e:
            logger.error(f"Failed to stop server: {e}")
            logger.error(traceback.format_exc())


class MCPProtocolHandler(unohelper.Base, XServiceInfo, XDispatchProvider, XDispatch):
    """Protocol handler for MCP extension menu commands"""

    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None
        logger.info("MCPProtocolHandler.__init__ called, ctx=%s", ctx)

    # XServiceInfo
    def getImplementationName(self):
        return IMPLEMENTATION_NAME

    def supportsService(self, name):
        return name in SERVICE_NAMES

    def getSupportedServiceNames(self):
        return SERVICE_NAMES

    # XDispatchProvider
    def queryDispatch(self, url, target, flags):
        logger.info("queryDispatch: Complete=%r Protocol=%r Path=%r",
                     url.Complete, url.Protocol, url.Path)
        # Accept ALL service: URLs for our protocol
        if url.Complete and "org.mcp.libreoffice.MCPExtension" in url.Complete:
            logger.info("queryDispatch: ACCEPTING")
            return self
        return None

    def queryDispatches(self, requests):
        return tuple([self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags) for r in requests])

    # XDispatch
    def dispatch(self, url, args):
        logger.info("dispatch called: Complete=%r", url.Complete)
        try:
            if "?" in url.Complete:
                command = url.Complete.split("?")[1]
                logger.info("Executing command: %s", command)

                if command == "start_mcp_server":
                    # Run in thread to not block UI
                    threading.Thread(target=_start_server, daemon=True).start()
                elif command == "stop_mcp_server":
                    threading.Thread(target=_stop_server, daemon=True).start()
                elif command == "restart_mcp_server":
                    def restart():
                        _stop_server()
                        _start_server()
                    threading.Thread(target=restart, daemon=True).start()
                elif command == "get_status":
                    logger.info(f"Server status: started={_server_started}")
                else:
                    logger.warning(f"Unknown command: {command}")

        except Exception as e:
            logger.error(f"Error in dispatch: {e}")
            logger.error(traceback.format_exc())

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass


# Component factory - pass the class directly, not a function
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MCPProtocolHandler,
    IMPLEMENTATION_NAME,
    SERVICE_NAMES
)
logger.info("=== addImplementation completed ===")

# ── Auto-start the MCP server on extension load ──
logger.info("=== Auto-starting MCP server... ===")
try:
    _start_server()
    logger.info("=== Auto-start complete, server_started=%s ===", _server_started)
except Exception as e:
    logger.error("=== Auto-start FAILED: %s ===", e)
    logger.error(traceback.format_exc())
