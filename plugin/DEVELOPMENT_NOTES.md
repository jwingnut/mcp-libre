# LibreOffice MCP Extension - Development Notes

## Critical Lessons Learned

### 1. LibreOffice Python Environment Limitations

#### Windows vs Linux
- **Windows**: LibreOffice's embedded Python lacks required modules (`asyncio`, `threading` with HTTP servers, proper socket support). The extension installs but silently fails at runtime.
- **Linux (Ubuntu)**: Full Python module support. The extension works as expected.

**Recommendation**: Target Linux for production use of this extension.

#### Python Imports
- **Relative imports don't work** in LibreOffice's Python environment
- The `pythonpath` directory is NOT automatically in `sys.path`
- **Must add to sys.path first**, then use absolute imports:
  ```python
  # At the top of registration.py:
  import sys
  import os
  _this_dir = os.path.dirname(__file__)
  if _this_dir not in sys.path:
      sys.path.insert(0, _this_dir)

  # Then use absolute imports:
  import ai_interface
  ai_interface.start_ai_interface(port=8765)
  ```

### 2. Extension Installation

#### Snap vs APT LibreOffice
- **Snap installation**: Does NOT expose `unopkg` command properly. Library errors occur.
- **APT installation**: Provides working `unopkg` at `/usr/bin/unopkg`

**Solution**: Use APT-installed LibreOffice:
```bash
sudo snap remove libreoffice
sudo apt install -y libreoffice libreoffice-writer
```

#### Line Endings
Files checked out from Git on Windows have CRLF line endings. Linux requires LF.

**Fix all files before building**:
```bash
find plugin -type f \( -name "*.py" -o -name "*.xml" -o -name "*.xcu" -o -name "*.txt" \) -exec sed -i 's/\r$//' {} \;
```

### 3. Protocol Handler Architecture

#### Required Interfaces
For menu items to be enabled (not greyed out), the extension must implement:
- `XDispatchProvider` - provides dispatch objects for URLs
- `XDispatch` - executes the dispatch
- `XServiceInfo` - service information
- `XInitialization` - receives frame during initialization

#### Service Registration
Register as `com.sun.star.frame.ProtocolHandler`:
```python
SERVICE_NAMES = ("com.sun.star.frame.ProtocolHandler",)
```

#### Factory Pattern
Pass the class directly to `addImplementation`, not a factory function:
```python
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MCPProtocolHandler,  # Class, not function
    IMPLEMENTATION_NAME,
    SERVICE_NAMES
)
```

### 4. Extension Package Structure

The `.oxt` file must include:
```
libreoffice-mcp-extension.oxt
├── META-INF/
│   └── manifest.xml      # Lists all components
├── pythonpath/
│   ├── registration.py   # Protocol handler
│   ├── ai_interface.py   # HTTP server
│   ├── mcp_server.py     # MCP implementation
│   └── uno_bridge.py     # UNO API wrapper
├── Addons.xcu            # Menu definitions (MUST be *.xcu, not *.xml)
├── ProtocolHandler.xcu   # Protocol handler config
├── description.xml       # Extension metadata
├── description-en.txt    # Description text
└── release-notes-en.txt  # Release notes
```

**Critical**: The build command must include `*.xcu` files:
```bash
zip -r extension.oxt META-INF/ pythonpath/ *.xml *.xcu *.txt
```

### 5. UNO API Imports

Some UNO imports may not be available in all configurations. Use try/except:
```python
try:
    from com.sun.star.presentation import XPresentationDocument
except ImportError:
    XPresentationDocument = None
```

### 6. HTTP Server in Extension

#### Threading
The HTTP server must run in a background thread to not block the UI:
```python
threading.Thread(target=_start_server, daemon=True).start()
```

#### Server Lifecycle
Don't use context managers that close the server:
```python
# WRONG - server closes when method returns
with socketserver.TCPServer(("", port), Handler) as server:
    server.serve_forever()

# CORRECT - server stays alive
self.server = socketserver.TCPServer(("", port), Handler)
self.server.serve_forever()
```

### 7. Global State

Since each dispatch may create a new handler instance, use module-level globals for shared state:
```python
_server_started = False
_ai_interface = None
_lock = threading.Lock()

def _start_server():
    global _server_started, _ai_interface
    with _lock:
        if _server_started:
            return
        # ... start server
```

### 8. Debugging

#### View Logs
LibreOffice Python logs go to stderr. Run from terminal to see them:
```bash
libreoffice --writer
```

Or capture output:
```bash
libreoffice --writer 2>&1 | tee /tmp/lo-output.log
```

#### Check Extension Installation
```bash
unopkg list | grep -i mcp
```

#### Check Port
```bash
ss -tulpn | grep 8765
curl http://localhost:8765/health
```

## Build & Install Commands

```bash
# Fix line endings
find plugin -type f \( -name "*.py" -o -name "*.xml" -o -name "*.xcu" \) -exec sed -i 's/\r$//' {} \;

# Build extension
cd plugin
rm -f ../build/libreoffice-mcp-extension.oxt
zip -r ../build/libreoffice-mcp-extension.oxt META-INF/ pythonpath/ *.xml *.xcu *.txt -x "*.pyc" "*/__pycache__/*"

# Remove old extension
unopkg remove org.mcp.libreoffice.extension

# Install new extension
unopkg add ~/libreoffice-mcp-extension.oxt

# Start LibreOffice
libreoffice --writer
```

## Testing

1. Open LibreOffice Writer
2. Go to **Tools → MCP Server → Start MCP Server**
3. Test the API:
   ```bash
   curl http://localhost:8765/health
   curl http://localhost:8765/tools
   ```

## Known Issues

1. Extension does not work on Windows due to Python environment limitations
2. Snap-installed LibreOffice doesn't expose unopkg properly
3. Files from Windows Git checkouts need line ending conversion
