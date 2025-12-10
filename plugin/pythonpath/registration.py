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
from com.sun.star.lang import XInitialization

# Add the pythonpath directory to sys.path for imports
_this_dir = os.path.dirname(__file__)
if _this_dir not in sys.path:
    sys.path.insert(0, _this_dir)

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger("MCPExtension")

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


class MCPProtocolHandler(unohelper.Base, XServiceInfo, XDispatchProvider, XDispatch, XInitialization):
    """Protocol handler for MCP extension menu commands"""

    def __init__(self, ctx):
        self.ctx = ctx
        self.frame = None
        logger.debug("MCPProtocolHandler created")

    # XInitialization
    def initialize(self, args):
        if args:
            self.frame = args[0]
        logger.debug("MCPProtocolHandler initialized with frame")

    # XServiceInfo
    def getImplementationName(self):
        return IMPLEMENTATION_NAME

    def supportsService(self, name):
        return name in SERVICE_NAMES

    def getSupportedServiceNames(self):
        return SERVICE_NAMES

    # XDispatchProvider
    def queryDispatch(self, url, target, flags):
        logger.debug(f"queryDispatch: {url.Complete}")
        if url.Protocol == "service:":
            return self
        return None

    def queryDispatches(self, requests):
        return tuple([self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags) for r in requests])

    # XDispatch
    def dispatch(self, url, args):
        logger.info(f"dispatch called: {url.Complete}")
        try:
            if "?" in url.Complete:
                command = url.Complete.split("?")[1]
                logger.info(f"Executing command: {command}")

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
