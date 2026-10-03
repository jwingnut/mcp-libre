import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawn, execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { WebSocketServer } from 'ws';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PUBLIC_DIR = path.join(__dirname, 'public');

const PORT = parseInt(process.env.LIBREOFFICE_PORT || process.env.PORT || '8765', 10);

const IS_HEADLESS = process.argv.includes('--headless') ||
  process.env.HEADLESS === 'true' ||
  process.env.HEADLESS === '1' ||
  process.env.LIBREOFFICE_HEADLESS === 'true' ||
  process.env.LIBREOFFICE_HEADLESS === '1';

const AUTO_OPEN_BROWSER = process.argv.includes('--browser') ||
  process.argv.includes('--open') ||
  process.env.AUTO_OPEN_BROWSER === 'true' ||
  process.env.AUTO_OPEN_BROWSER === '1';

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.mjs': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.wasm': 'application/wasm',
};

// Track active browser Wasm sessions
let headlessProcess = null;
let headlessTmpDir = null;
const pendingRequests = new Map();
const connectedClients = new Set();

function getActiveClient() {
  const arr = Array.from(connectedClients);
  for (let i = arr.length - 1; i >= 0; i--) {
    if (arr[i].readyState === 1) return arr[i];
  }
  return null;
}

// Find Chrome/Chromium executable on the current OS
function findChromeBinary() {
  if (process.env.CHROME_BIN && fs.existsSync(process.env.CHROME_BIN)) {
    return process.env.CHROME_BIN;
  }
  const candidates = [];
  if (process.platform === 'linux') {
    candidates.push(
      '/usr/bin/google-chrome-stable',
      '/usr/bin/google-chrome',
      '/usr/bin/chromium',
      '/usr/bin/chromium-browser',
      '/snap/bin/chromium'
    );
  } else if (process.platform === 'darwin') {
    candidates.push(
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Chromium.app/Contents/MacOS/Chromium'
    );
  } else if (process.platform === 'win32') {
    const prefixes = [process.env.LOCALAPPDATA, process.env.PROGRAMFILES, process.env['PROGRAMFILES(X86)']];
    for (const p of prefixes) {
      if (p) {
        candidates.push(path.join(p, 'Google/Chrome/Application/chrome.exe'));
        candidates.push(path.join(p, 'Chromium/Application/chrome.exe'));
      }
    }
  }

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) return candidate;
  }

  try {
    const cmd = process.platform === 'win32'
      ? 'where chrome'
      : 'which google-chrome-stable || which google-chrome || which chromium';
    const found = execSync(cmd, { stdio: ['ignore', 'pipe', 'ignore'], encoding: 'utf-8' }).trim().split('\n')[0];
    if (found && fs.existsSync(found)) return found;
  } catch {}

  return null;
}

// Spawn headless Chrome runner for automated background Wasm execution
function spawnHeadlessRunner(url) {
  if (headlessProcess) return headlessProcess;

  const chromeBin = findChromeBinary();
  if (!chromeBin) {
    console.warn(`[Headless] Warning: No Chrome/Chromium binary found. Please open ${url} in your browser.`);
    return null;
  }

  try {
    headlessTmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lowa-headless-'));
    const args = [
      '--headless=new',
      '--disable-gpu',
      '--no-sandbox',
      '--disable-dev-shm-usage',
      '--enable-unsafe-swiftshader',
      '--disable-background-networking',
      '--disable-default-apps',
      '--disable-extensions',
      '--disable-sync',
      '--disable-translate',
      '--mute-audio',
      '--no-first-run',
      `--user-data-dir=${headlessTmpDir}`,
      url
    ];

    console.log(`[Headless] Launching background Wasm runner: ${chromeBin}`);
    headlessProcess = spawn(chromeBin, args, {
      stdio: ['ignore', 'ignore', 'pipe']
    });

    headlessProcess.stderr.on('data', (data) => {
      const msg = data.toString();
      if (msg.includes('Fatal') || msg.includes('Crash')) {
        console.error(`[Headless Error] ${msg.trim()}`);
      }
    });

    headlessProcess.on('exit', (code, signal) => {
      console.log(`[Headless] Runner exited (code: ${code}, signal: ${signal})`);
      headlessProcess = null;
      cleanupHeadlessTmpDir();
    });

    return headlessProcess;
  } catch (err) {
    console.error(`[Headless] Failed to spawn runner: ${err.message}`);
    return null;
  }
}

function cleanupHeadlessTmpDir() {
  if (headlessTmpDir && fs.existsSync(headlessTmpDir)) {
    try {
      fs.rmSync(headlessTmpDir, { recursive: true, force: true });
    } catch {}
    headlessTmpDir = null;
  }
}

function stopHeadless() {
  if (headlessProcess) {
    try {
      headlessProcess.kill('SIGTERM');
    } catch {}
    headlessProcess = null;
  }
  cleanupHeadlessTmpDir();
}

// Open URL in interactive desktop browser
function openDesktopBrowser(url) {
  let cmd;
  let args = [url];

  if (process.platform === 'darwin') {
    cmd = 'open';
  } else if (process.platform === 'win32') {
    cmd = 'cmd.exe';
    args = ['/c', 'start', '""', url];
  } else {
    cmd = 'xdg-open';
  }

  try {
    const cp = spawn(cmd, args, { detached: true, stdio: 'ignore' });
    cp.unref();
    cp.on('error', () => {
      const fallback = findChromeBinary() || 'firefox';
      try {
        const fbCp = spawn(fallback, [url], { detached: true, stdio: 'ignore' });
        fbCp.unref();
      } catch {}
    });
    return true;
  } catch (e) {
    console.error(`[Browser] Failed to launch ${cmd}:`, e.message);
    return false;
  }
}

process.on('exit', () => { stopHeadless(); });
process.on('SIGINT', () => { stopHeadless(); process.exit(0); });
process.on('SIGTERM', () => { stopHeadless(); process.exit(0); });

// Create HTTP server
const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host}`);
  const pathname = parsedUrl.pathname;

  // Cross-Origin Isolation headers required for SharedArrayBuffer / Wasm pthreads
  res.setHeader('Cross-Origin-Opener-Policy', 'same-origin');
  res.setHeader('Cross-Origin-Embedder-Policy', 'require-corp');
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // --- MCP API Endpoints ---
  if (pathname === '/health') {
    const client = getActiveClient();
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      status: 'healthy',
      server: 'ZetaJS Wasm MCP Bridge',
      version: '1.0.0',
      browser_connected: !!client,
      connected_clients: connectedClients.size,
      mode: IS_HEADLESS ? 'headless' : 'browser',
      headless_running: !!headlessProcess,
      url: `http://localhost:${PORT}`
    }));
    return;
  }

  // Trigger desktop browser window
  if ((pathname === '/tools/open_browser_live' || pathname === '/open_browser') && (req.method === 'POST' || req.method === 'GET')) {
    const url = `http://localhost:${PORT}`;
    const opened = openDesktopBrowser(url);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      success: opened,
      message: opened ? `Opened ${url} in your desktop browser` : `Failed to launch desktop browser for ${url}`,
      url,
      mode: IS_HEADLESS ? 'headless' : 'browser',
      browser_connected: !!getActiveClient()
    }));
    return;
  }

  if (pathname === '/' && req.headers.accept?.includes('application/json')) {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      server: 'ZetaJS Wasm MCP Bridge',
      version: '1.0.0',
      status: 'running',
      mode: IS_HEADLESS ? 'headless' : 'browser',
      endpoints: ['/health', '/tools', '/tools/:tool_name', '/open_browser']
    }));
    return;
  }

  if (pathname === '/tools' && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({
      tools: [
        { name: 'create_document_live', description: 'Create a new Writer or Impress document via ZetaJS UNO' },
        { name: 'insert_text_live', description: 'Insert text into the active document via ZetaJS UNO' },
        { name: 'format_text_live', description: 'Format text properties via ZetaJS UNO' },
        { name: 'get_text_content_live', description: 'Retrieve document text content via ZetaJS UNO' },
        { name: 'get_document_info_live', description: 'Get document statistics and title via ZetaJS UNO' },
        { name: 'open_browser_live', description: 'Open live document in desktop browser' }
      ]
    }));
    return;
  }

  if (pathname.startsWith('/tools/') && req.method === 'POST') {
    const toolName = pathname.replace('/tools/', '');
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      let params = {};
      try {
        if (body.trim()) params = JSON.parse(body);
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Invalid JSON body' }));
        return;
      }

      const client = getActiveClient();
      if (!client) {
        res.writeHead(503, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          error: 'No active ZetaOffice Wasm session connected. Open http://localhost:' + PORT + ' in your browser.',
          status: 'no_browser_session'
        }));
        return;
      }

      const reqId = 'req_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
      const promise = new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
          pendingRequests.delete(reqId);
          reject(new Error('Timed out waiting for ZetaJS UNO response (15s)'));
        }, 15000);

        pendingRequests.set(reqId, { resolve, reject, timeout, responded: false });
      });

      // Dispatch to connected Wasm sessions
      for (const c of connectedClients) {
        if (c.readyState === 1) {
          c.send(JSON.stringify({
            type: 'mcp_request',
            id: reqId,
            tool: toolName,
            params
          }));
        }
      }

      try {
        const result = await promise;
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify(result));
      } catch (err) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  // --- Static Files Delivery ---
  let filePath = path.join(PUBLIC_DIR, pathname === '/' ? 'index.html' : pathname);
  
  // Normalize and prevent directory traversal
  filePath = path.normalize(filePath);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('File not found: ' + pathname);
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  });
});

// Setup WebSocket Server
const wss = new WebSocketServer({ server, path: '/ws' });

wss.on('connection', (ws, req) => {
  console.log(`[WebSocket] New browser tab connected from ${req.socket.remoteAddress}`);
  connectedClients.add(ws);

  ws.on('message', (message) => {
    try {
      const data = JSON.parse(message.toString());
      if (data.type === 'ready') {
        console.log(`[ZetaJS] ZetaOffice UNO Engine ready in browser: ${data.info || 'OK'}`);
      } else if (data.type === 'mcp_response') {
        const pending = pendingRequests.get(data.id);
        if (pending && !pending.responded) {
          pending.responded = true;
          clearTimeout(pending.timeout);
          pendingRequests.delete(data.id);
          pending.resolve(data.result);
        }
      } else if (data.type === 'log') {
        console.log(`[Browser Console] ${data.message}`);
      }
    } catch (e) {
      console.error('[WebSocket] Error parsing message:', e);
    }
  });

  ws.on('close', () => {
    console.log('[WebSocket] Browser tab disconnected');
    connectedClients.delete(ws);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  const url = `http://localhost:${PORT}`;
  console.log(`
============================================================
🚀 ZetaJS Wasm MCP Bridge running!
------------------------------------------------------------
📍 Mode:                    ${IS_HEADLESS ? 'Headless (Background Node.js Runner)' : 'Interactive Browser'}
📍 Browser UI & Wasm Host:  ${url}
📍 MCP API Endpoint:        ${url}/tools
📍 Health Check:            ${url}/health

${IS_HEADLESS ? '🤖 Headless mode active: starting silent background Wasm runner...' : '👉 Open ' + url + ' in your browser to view the live canvas.'}
============================================================
`);

  if (IS_HEADLESS) {
    spawnHeadlessRunner(url);
  } else if (AUTO_OPEN_BROWSER) {
    console.log(`[Browser] Automatically launching desktop browser: ${url}`);
    openDesktopBrowser(url);
  }
});
