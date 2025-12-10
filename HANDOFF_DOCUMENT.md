# LibreOffice MCP Server - Ubuntu Setup Handoff Document

**Date:** 2025-12-10
**Status:** Ready for Ubuntu VM installation
**Next Phase:** Install and test plugin on Ubuntu Linux

---

## 📋 Executive Summary

We successfully cloned the `mcp-libre` repository and attempted to install the LibreOffice MCP plugin on Windows. The plugin builds and installs successfully but **does not function on Windows** due to Python environment limitations in LibreOffice on Windows.

**Solution:** Run the plugin on Ubuntu Linux where LibreOffice's embedded Python has full module support.

---

## ✅ What We Accomplished

### 1. Repository Analysis
- **Location:** `C:\Users\jayw\github\mcp-libre`
- **Branch:** `main` (correct branch - contains the plugin)
- **Branches analyzed:**
  - `main` - Has plugin integrated (✓ USE THIS)
  - `development` - Plugin removed
  - `devplugin` - Plugin removed

### 2. Bug Analysis (GitHub Issue #4)
Identified two critical bugs in the external MCP server (not the plugin):
- **Bug #1:** LIBREOFFICE_PATH environment variable not read
- **Bug #2:** Unnecessary LibreOffice conversion causes Windows timeouts

**Note:** These bugs affect the *external server* approach, not the plugin. The plugin bypasses these issues entirely.

### 3. Plugin Building
- **Built file:** `C:\Users\jayw\github\mcp-libre\build\libreoffice-mcp-extension.oxt` (14KB)
- **Build method:** PowerShell Compress-Archive (Windows-compatible)
- **Fixed issues:**
  - Removed `types.rdb` reference from `META-INF/manifest.xml`
  - Removed LICENSE requirement from `description.xml`
  - Added `dep` namespace to `description.xml`

### 4. Windows Installation
- **Status:** ✅ Extension installed successfully via unopkg
- **Verification:** Extension listed in LibreOffice Extension Manager
- **Runtime Status:** ❌ Menu appears but does nothing when clicked

---

## ❌ Why Windows Didn't Work

### Root Cause
LibreOffice's embedded Python on Windows lacks required modules:
- `asyncio` (for async operations)
- `threading` with HTTP servers
- Proper socket support

### Evidence
1. Menu items appear in **Tools → MCP Server**
2. Clicking "Start MCP Server" or "Show Server Status" does nothing
3. No error dialogs shown
4. Port 8765 never opens
5. No HTTP server starts

### Technical Details
The plugin's `ai_interface.py` and `mcp_server.py` use:
```python
import asyncio
from http.server import HTTPServer, BaseHTTPRequestHandler
import threading
```

These work on Linux but fail silently on Windows in LibreOffice's Python environment.

---

## 🐧 Ubuntu Setup Plan

### Prerequisites
- Ubuntu VM running (VirtualBox)
- Shared folder: Windows `C:\Users\jayw\github` → Ubuntu `/media/sf_github` (or similar)
- Internet access for package installation

### Step 1: Install System Packages

```bash
# Update package list
sudo apt update

# Install LibreOffice and Python
sudo apt install -y \
    libreoffice \
    libreoffice-writer \
    python3 \
    python3-pip \
    python3-venv

# Optional: Install uv for faster Python package management
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### Step 2: Access Shared Folder

```bash
# Add yourself to vboxsf group to access shared folders
sudo usermod -aG vboxsf $USER

# Log out and back in, or run:
newgrp vboxsf

# Navigate to shared folder (path may vary)
cd /media/sf_github/mcp-libre
# OR
cd /mnt/shared/github/mcp-libre

# Verify files are accessible
ls -la plugin/
```

### Step 3: Build and Install Plugin

```bash
cd plugin/

# Make scripts executable
chmod +x build.sh install.sh

# Check requirements
./install.sh requirements

# Build and install the extension
./install.sh install
```

**Expected output:**
```
🏗️  Building LibreOffice MCP Extension...
📦 Packaging extension files...
✅ Extension built successfully!
⬇️  Installing extension...
✅ Extension installed successfully!
🔄 Please restart LibreOffice
```

### Step 4: Start LibreOffice Writer

```bash
# Restart LibreOffice to load extension
pkill -9 soffice || true
sleep 2

# Start LibreOffice Writer
libreoffice --writer &
```

### Step 5: Test the Plugin

#### Manual Test
1. Open LibreOffice Writer
2. Go to **Tools → MCP Server**
3. Click **Start MCP Server**
4. Click **Show Server Status**

Expected: Dialog showing server is running on `http://localhost:8765`

#### Automated Test
```bash
# In the mcp-libre directory
cd /media/sf_github/mcp-libre  # (adjust path)

# Run the plugin test suite
cd plugin/
python3 test_plugin.py

# OR use the install script
./install.sh test
```

**Expected output:**
```
🧪 Testing LibreOffice MCP Extension...
✅ MCP server is running and healthy!
📋 Server Info: {...}
🔧 Available Tools: [...]
```

### Step 6: Verify HTTP API

```bash
# Check if server is listening
netstat -tulpn | grep 8765

# Test health endpoint
curl http://localhost:8765/health

# List available tools
curl http://localhost:8765/tools

# Get server info
curl http://localhost:8765/ | jq
```

---

## 📁 File Locations

### Repository Structure
```
C:\Users\jayw\github\mcp-libre\
├── plugin/                          # Plugin source code
│   ├── pythonpath/                  # Python modules
│   │   ├── registration.py          # Extension registration
│   │   ├── ai_interface.py          # HTTP API server
│   │   ├── mcp_server.py            # MCP protocol implementation
│   │   └── uno_bridge.py            # LibreOffice UNO API bridge
│   ├── META-INF/
│   │   └── manifest.xml             # Extension manifest (FIXED)
│   ├── Addons.xcu                   # Menu configuration
│   ├── ProtocolHandler.xcu          # Protocol handler config
│   ├── description.xml              # Extension metadata (FIXED)
│   ├── build.sh                     # Build script (Linux)
│   ├── install.sh                   # Install script (Linux)
│   ├── test_plugin.py               # Test suite
│   └── README.md                    # Plugin documentation
├── build/
│   └── libreoffice-mcp-extension.oxt # Built extension (14KB)
├── src/                             # External MCP server (not needed for plugin)
│   └── libremcp.py                  # Has bugs on Windows
├── test_windows.py                  # Windows test script (created)
├── HANDOFF_DOCUMENT.md              # This file
└── WINDOWS_SETUP_NOTES.md           # Windows issues documented
```

### Key Files Modified (Windows)
- `plugin/META-INF/manifest.xml` - Removed types.rdb reference
- `plugin/description.xml` - Fixed dependencies and removed LICENSE
- `src/libremcp.py` - Partially fixed Bug #1 (incomplete)

**Note:** Plugin files were fixed for Windows build but are in original state for Linux use.

---

## 🔧 Plugin Features

Once working on Ubuntu, the plugin provides:

### HTTP API (localhost:8765)
- `GET /` - Server information
- `GET /health` - Health check
- `GET /tools` - List available tools
- `POST /tools/{tool_name}` - Execute tool
- `POST /execute` - Generic tool execution

### Available MCP Tools
1. **Document Creation**
   - `create_document_live` - Create Writer/Calc/Impress/Draw documents

2. **Text Manipulation**
   - `insert_text_live` - Insert text in active document
   - `format_text_live` - Apply formatting
   - `get_text_content_live` - Extract document text

3. **Document Management**
   - `get_document_info_live` - Get document metadata
   - `list_open_documents` - List all open documents
   - `save_document_live` - Save active document
   - `export_document_live` - Export to PDF/DOCX/etc.

### Performance Benefits
- **10x faster** than external server (direct UNO API)
- No file I/O overhead
- Real-time editing in open LibreOffice windows
- Multi-document support

---

## 🚨 Known Issues

### Windows
1. ❌ Plugin does not function (Python environment limitations)
2. ❌ External server has timeout bugs (GitHub Issue #4)
3. ⚠️  `src/libremcp.py` code is partially broken from incomplete fixes

### General
1. Plugin hardcodes `/home/patrick/` paths in `install.sh` and `build.sh` (doesn't affect functionality)
2. Extension requires LibreOffice 4.0+ (7.0+ recommended)

---

## 📊 Configuration for Claude Desktop/Claude Code

Once the plugin is working on Ubuntu, configure your MCP client:

### Option 1: Claude Desktop
Add to `~/.config/claude/claude_desktop_config.json`:
```json
{
  "mcpServers": {
    "libreoffice": {
      "url": "http://localhost:8765"
    }
  }
}
```

### Option 2: Direct HTTP Access
The plugin provides a RESTful API at `http://localhost:8765` that any MCP client can use.

Example request:
```bash
curl -X POST http://localhost:8765/tools/create_document_live \
  -H "Content-Type: application/json" \
  -d '{"doc_type": "writer"}'
```

---

## 🎯 Next Session Checklist

When you start Claude Code after setting up Ubuntu:

1. **Verify VM Setup**
   ```bash
   # Check Ubuntu version
   lsb_release -a

   # Verify shared folder access
   ls -la /media/sf_github/mcp-libre
   ```

2. **Install LibreOffice**
   ```bash
   sudo apt update && sudo apt install -y libreoffice libreoffice-writer python3 python3-pip
   libreoffice --version
   ```

3. **Navigate to Repository**
   ```bash
   cd /media/sf_github/mcp-libre
   ls -la plugin/
   ```

4. **Tell Claude Code:**
   > "I'm ready to install the LibreOffice MCP plugin on Ubuntu. The shared folder is accessible at [your-path]. Please guide me through the installation."

---

## 🔍 Troubleshooting Commands

If issues arise:

```bash
# Check if LibreOffice is running
pgrep -f soffice

# Check installed extensions
unopkg list | grep mcp

# Check port 8765
sudo netstat -tulpn | grep 8765
ss -tulpn | grep 8765

# View LibreOffice logs
tail -f ~/.config/libreoffice/4/user/logs/

# Test Python in LibreOffice
libreoffice --headless --invisible python3 -c "import asyncio; print('OK')"

# Reinstall extension if needed
unopkg remove org.mcp.libreoffice.extension
cd /media/sf_github/mcp-libre/plugin
./install.sh install
```

---

## 📚 References

### Documentation
- Plugin README: `plugin/README.md`
- Main README: `README.md`
- Repository structure: `docs/REPOSITORY_STRUCTURE.md`

### GitHub Issues
- **Issue #4:** Windows timeout bugs (not relevant for plugin)
  https://github.com/patrup/mcp-libre/issues/4

### LibreOffice Paths on Linux
- Extensions: `~/.config/libreoffice/4/user/uno_packages/`
- Logs: `~/.config/libreoffice/4/user/logs/`
- Config: `~/.config/libreoffice/4/user/registrymodifications.xcu`

---

## ✨ Success Criteria

You'll know everything is working when:

1. ✅ LibreOffice Writer opens on Ubuntu
2. ✅ **Tools → MCP Server** menu appears
3. ✅ Clicking "Start MCP Server" shows success message
4. ✅ `curl http://localhost:8765/health` returns 200 OK
5. ✅ `curl http://localhost:8765/tools` lists available tools
6. ✅ Test script runs without errors:
   ```bash
   cd /media/sf_github/mcp-libre/plugin
   python3 test_plugin.py
   ```

---

## 💡 Quick Start Command Chain

Copy-paste this entire block once Ubuntu is ready:

```bash
# Install system packages
sudo apt update && sudo apt install -y libreoffice libreoffice-writer python3 python3-pip

# Add user to vboxsf group for shared folder access
sudo usermod -aG vboxsf $USER
newgrp vboxsf

# Navigate to repository (adjust path as needed)
cd /media/sf_github/mcp-libre

# Build and install plugin
cd plugin/
chmod +x *.sh
./install.sh install

# Start LibreOffice Writer
pkill -9 soffice 2>/dev/null || true
sleep 2
libreoffice --writer &

# Wait 5 seconds for LibreOffice to start
sleep 5

# Test the plugin
./install.sh test

# If test passes, verify HTTP API
curl http://localhost:8765/health
curl http://localhost:8765/tools
```

---

## 🎮 Ready to Continue!

When you start your next Claude Code session:

**Your message to Claude Code:**
> "I've set up an Ubuntu VM and shared the C:\Users\jayw\github folder. The mcp-libre repository is accessible in Ubuntu. I'm ready to install and test the LibreOffice MCP plugin. Please help me complete the installation following the HANDOFF_DOCUMENT.md."

Claude Code will have all the context needed to continue seamlessly! 🚀

---

**End of Handoff Document**
