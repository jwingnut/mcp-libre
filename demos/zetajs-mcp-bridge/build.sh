#!/bin/bash
# ZetaJS WebAssembly MCP Bridge Build & Setup Script
# Zero-installation: No LibreOffice desktop installation or .oxt package required!

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "🏗️  Setting up ZetaJS WebAssembly MCP Bridge..."

# 1. Check Node.js
if ! command -v node >/dev/null 2>&1; then
    echo "❌ Node.js is required but not installed. Please install Node.js >= 18."
    exit 1
fi
echo "✅ Node.js $(node --version) detected"

# 2. Install dependencies (prefer pnpm, fallback to npm)
echo "📦 Installing bridge dependencies..."
if command -v pnpm >/dev/null 2>&1; then
    CI=true pnpm install
elif command -v npm >/dev/null 2>&1; then
    npm install
else
    echo "❌ Neither pnpm nor npm found in PATH."
    exit 1
fi

# 3. Ensure scripts are executable
chmod +x server.mjs mcp_stdio_server.mjs test_client.py

# 4. Verify runtime assets
if [ ! -f "public/zetajs/zeta.js" ] || [ ! -f "public/office_thread.js" ]; then
    echo "❌ Missing required WebAssembly / ZetaJS runtime assets in public/"
    exit 1
fi

echo ""
echo "🎉 Build & setup complete! No .oxt package is needed (pure WebAssembly execution)."
echo ""
echo "🚀 To run the bridge:"
echo "   1. Start the bridge daemon:  node server.mjs (or pnpm start)"
echo "   2. Open browser canvas:       http://localhost:8765"
echo "   3. Connect MCP client:       node mcp_stdio_server.mjs"
echo "   4. Or run test client:       python3 test_client.py"
