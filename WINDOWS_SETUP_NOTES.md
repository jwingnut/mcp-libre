# LibreOffice MCP on Windows - Setup Notes

## Current Status

### Plugin Extension (Not Recommended for Windows)
❌ The LibreOffice plugin extension has compatibility issues on Windows:
- LibreOffice's embedded Python on Windows lacks required modules (`asyncio`, threading HTTP servers)
- Menu items appear but do nothing when clicked
- No error messages displayed

**Installed at:** Extension is installed via unopkg but not functional

### External MCP Server (Recommended for Windows)
⚠️  Has two critical bugs from [GitHub Issue #4](https://github.com/patrup/mcp-libre/issues/4):

#### Bug #1: LIBREOFFICE_PATH not respected
- Environment variable documented in README but never checked
- **Status:** Fixed in working copy at `C:\Users\jayw\github\mcp-libre\src\libremcp.py`

#### Bug #2: Unnecessary LibreOffice conversion causes timeouts
- `create_document()` tries to use LibreOffice for conversion
- Causes 30-second timeouts and file locking errors on Windows
- **Status:** Needs to be fixed (partially modified)

## Recommended Next Steps

1. **Fix Bug #2 properly:**
   - Replace LibreOffice conversion with direct `_create_minimal_odt()` calls
   - This function already exists and creates valid ODT files

2. **Test External Server:**
   ```bash
   cd C:\Users\jayw\github\mcp-libre
  uv run python src/main.py
   ```

3. **Set Environment Variable:**
   ```powershell
   $env:LIBREOFFICE_PATH="D:\Program Files\LibreOffice\program\soffice.exe"
   ```

## Files Modified

- `C:\Users\jayw\github\mcp-libre\plugin\META-INF\manifest.xml` - Removed types.rdb reference
- `C:\Users\jayw\github\mcp-libre\plugin\description.xml` - Removed LICENSE requirement, added dep namespace
- `C:\Users\jayw\github\mcp-libre\src\libremcp.py` - Partially fixed (needs completion)

## Extension Build

Built extension located at:
```
C:\Users\jayw\github\mcp-libre\build\libreoffice-mcp-extension.oxt
```

## Recommendation

**Use the external MCP server approach** instead of the plugin for Windows. The external server is more reliable and easier to debug on Windows systems.
