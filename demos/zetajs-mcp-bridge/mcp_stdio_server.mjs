#!/usr/bin/env node
/**
 * MCP Stdio Server for ZetaJS WebAssembly LibreOffice Bridge
 * 
 * Provides the EXACT SAME 9 consolidated tool definitions and action signatures
 * as native mcp-libre (libreoffice_mcp_server.py).
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const LIBREOFFICE_URL = process.env.LIBREOFFICE_URL || "http://localhost:8765";

let bridgeProcess = null;

async function checkHealth(url) {
  try {
    const res = await fetch(`${url}/health`, { signal: AbortSignal.timeout(1000) });
    if (!res.ok) return false;
    const data = await res.json();
    return data.status === 'healthy';
  } catch {
    return false;
  }
}

async function isWasmReady(url) {
  try {
    const res = await fetch(`${url}/health`, { signal: AbortSignal.timeout(1000) });
    if (!res.ok) return false;
    const data = await res.json();
    return data.browser_connected === true;
  } catch {
    return false;
  }
}

async function ensureBridge() {
  const isHealthy = await checkHealth(LIBREOFFICE_URL);
  if (isHealthy) {
    return;
  }

  // Only auto-spawn if running against localhost
  if (!LIBREOFFICE_URL.includes("localhost") && !LIBREOFFICE_URL.includes("127.0.0.1")) {
    return;
  }

  console.error(`[MCP] LibreOffice bridge not running at ${LIBREOFFICE_URL}. Auto-spawning Wasm bridge...`);

  const serverScript = path.join(__dirname, 'server.mjs');
  const isHeadless = process.env.HEADLESS !== 'false' && process.env.HEADLESS !== '0';
  const args = [serverScript];
  if (isHeadless) {
    args.push('--headless');
  } else {
    args.push('--browser');
  }

  bridgeProcess = spawn(process.execPath, args, {
    stdio: ['ignore', 'pipe', 'inherit'],
    env: { ...process.env }
  });

  if (bridgeProcess.stdout) {
    bridgeProcess.stdout.on('data', (d) => {
      // Forward bridge logs to stderr so stdout remains pure JSON-RPC
      process.stderr.write(d);
    });
  }

  // Wait for server to become healthy and Wasm runner to connect (up to 15s)
  const startTime = Date.now();
  while (Date.now() - startTime < 15000) {
    await new Promise(r => setTimeout(r, 500));
    if (await isWasmReady(LIBREOFFICE_URL)) {
      console.error(`[MCP] LibreOffice Wasm bridge is ready and connected!`);
      return;
    }
  }
  console.error(`[MCP] Warning: Bridge started, but Wasm runner not yet confirmed connected.`);
}

function cleanupBridge() {
  if (bridgeProcess) {
    try {
      bridgeProcess.kill('SIGTERM');
    } catch {}
    bridgeProcess = null;
  }
}

process.on('exit', cleanupBridge);
process.on('SIGINT', () => { cleanupBridge(); process.exit(0); });
process.on('SIGTERM', () => { cleanupBridge(); process.exit(0); });

const server = new McpServer({
  name: "libreoffice-wasm",
  version: "1.0.0"
});

// Helper function to call the ZetaJS Wasm HTTP bridge
async function callLibreoffice(endpoint, method = "POST", data = {}) {
  await ensureBridge();
  try {
    const options = {
      method,
      headers: { "Content-Type": "application/json" }
    };
    if (method === "POST") {
      options.body = JSON.stringify(data);
    }
    const res = await fetch(`${LIBREOFFICE_URL}${endpoint}`, options);
    return await res.json();
  } catch (err) {
    return {
      error: `Cannot connect to LibreOffice Wasm bridge at ${LIBREOFFICE_URL}. Error: ${err.message}`
    };
  }
}

function jsonResponse(data) {
  return {
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }]
  };
}

// =============================================================================
// CONSOLIDATED TOOL 1: document
// Actions: create, info, list, content, status, styles,
//          create_style, edit_style, delete_style, style_properties
// =============================================================================
server.tool(
  "document",
  "Manage LibreOffice documents - create, get info, list, content, status, open in browser, paragraph styles (list, create, edit, delete).",
  {
    action: z.enum([
      "create", "info", "list", "content", "status",
      "styles", "create_style", "edit_style", "delete_style", "style_properties",
      "open_browser"
    ]).describe("The operation to perform (use 'open_browser' to view the live document in your desktop browser)"),
    doc_type: z.enum(["writer", "calc", "impress", "draw"]).default("writer").describe("Document type for create action"),
    style_name: z.string().optional().describe("Style name for create_style / edit_style / delete_style actions"),
    parent_style: z.string().optional().describe("Parent style for create_style / edit_style (default: Standard)"),
    font_name: z.string().optional().describe("Font family for create_style / edit_style"),
    font_size: z.number().optional().describe("Font size in points for create_style / edit_style"),
    bold: z.boolean().optional().describe("Bold weight for create_style / edit_style"),
    italic: z.boolean().optional().describe("Italic posture for create_style / edit_style"),
    underline: z.boolean().optional().describe("Underline for create_style / edit_style"),
    alignment: z.enum(["left", "right", "center", "justify"]).optional().describe("Alignment for create_style / edit_style")
  },
  async ({ action, doc_type, style_name, parent_style, font_name, font_size, bold, italic, underline, alignment }) => {
    if (action === "create") {
      return jsonResponse(await callLibreoffice("/tools/create_document_live", "POST", { doc_type }));
    } else if (action === "info") {
      return jsonResponse(await callLibreoffice("/tools/get_document_info_live", "POST", {}));
    } else if (action === "list") {
      return jsonResponse(await callLibreoffice("/tools/list_open_documents", "POST", {}));
    } else if (action === "content") {
      return jsonResponse(await callLibreoffice("/tools/get_text_content_live", "POST", {}));
    } else if (action === "status") {
      return jsonResponse(await callLibreoffice("/health", "GET"));
    } else if (action === "open_browser") {
      return jsonResponse(await callLibreoffice("/tools/open_browser_live", "POST", {}));
    } else if (action === "styles") {
      return jsonResponse(await callLibreoffice("/tools/list_paragraph_styles_live", "POST", {}));
    } else if (action === "create_style") {
      if (!style_name) return jsonResponse({ error: "Action 'create_style' requires parameter 'style_name'" });
      const data = { style_name };
      if (parent_style) data.parent_style = parent_style;
      if (font_name) data.font_name = font_name;
      if (font_size !== undefined) data.font_size = font_size;
      if (bold !== undefined) data.bold = bold;
      if (italic !== undefined) data.italic = italic;
      if (underline !== undefined) data.underline = underline;
      if (alignment) data.alignment = alignment;
      return jsonResponse(await callLibreoffice("/tools/create_paragraph_style_live", "POST", data));
    } else if (action === "edit_style") {
      if (!style_name) return jsonResponse({ error: "Action 'edit_style' requires parameter 'style_name'" });
      const data = { style_name };
      if (parent_style) data.parent_style = parent_style;
      if (font_name) data.font_name = font_name;
      if (font_size !== undefined) data.font_size = font_size;
      if (bold !== undefined) data.bold = bold;
      if (italic !== undefined) data.italic = italic;
      if (underline !== undefined) data.underline = underline;
      if (alignment) data.alignment = alignment;
      return jsonResponse(await callLibreoffice("/tools/edit_paragraph_style_live", "POST", data));
    } else if (action === "delete_style") {
      if (!style_name) return jsonResponse({ error: "Action 'delete_style' requires parameter 'style_name'" });
      return jsonResponse(await callLibreoffice("/tools/delete_paragraph_style_live", "POST", { style_name }));
    } else if (action === "style_properties") {
      if (!style_name) return jsonResponse({ error: "Action 'style_properties' requires parameter 'style_name'" });
      return jsonResponse(await callLibreoffice("/tools/get_style_properties_live", "POST", { style_name }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 2: structure
// Actions: outline, paragraph, range, count
// =============================================================================
server.tool(
  "structure",
  "Navigate and inspect document structure - outline, paragraphs, ranges, count.",
  {
    action: z.enum(["outline", "paragraph", "range", "count"]).describe("The operation to perform"),
    n: z.number().int().optional().describe("Paragraph number (1-indexed) for 'paragraph' action"),
    start: z.number().int().optional().describe("Starting paragraph number for 'range' action"),
    end: z.number().int().optional().describe("Ending paragraph number for 'range' action")
  },
  async ({ action, n, start, end }) => {
    if (action === "outline") {
      return jsonResponse(await callLibreoffice("/tools/get_document_outline_live", "POST", {}));
    } else if (action === "paragraph") {
      if (n === undefined) return jsonResponse({ error: "Action 'paragraph' requires parameter 'n'" });
      return jsonResponse(await callLibreoffice("/tools/get_paragraph_live", "POST", { n }));
    } else if (action === "range") {
      if (start === undefined || end === undefined) return jsonResponse({ error: "Action 'range' requires parameters 'start' and 'end'" });
      return jsonResponse(await callLibreoffice("/tools/get_paragraphs_range_live", "POST", { start, end }));
    } else if (action === "count") {
      return jsonResponse(await callLibreoffice("/tools/get_paragraph_count_live", "POST", {}));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 3: cursor
// Actions: goto_paragraph, goto_position, position, context
// =============================================================================
server.tool(
  "cursor",
  "Navigate cursor position in the document.",
  {
    action: z.enum(["goto_paragraph", "goto_position", "position", "context"]).describe("The operation to perform"),
    n: z.number().int().optional().describe("Paragraph number (1-indexed) for 'goto_paragraph' action"),
    char_pos: z.number().int().optional().describe("Character position (0-indexed) for 'goto_position' action"),
    chars: z.number().int().default(100).describe("Number of characters before/after cursor for 'context' action (default: 100)")
  },
  async ({ action, n, char_pos, chars }) => {
    if (action === "goto_paragraph") {
      if (n === undefined) return jsonResponse({ error: "Action 'goto_paragraph' requires parameter 'n'" });
      return jsonResponse(await callLibreoffice("/tools/goto_paragraph_live", "POST", { n }));
    } else if (action === "goto_position") {
      if (char_pos === undefined) return jsonResponse({ error: "Action 'goto_position' requires parameter 'char_pos'" });
      return jsonResponse(await callLibreoffice("/tools/goto_position_live", "POST", { char_pos }));
    } else if (action === "position") {
      return jsonResponse(await callLibreoffice("/tools/get_cursor_position_live", "POST", {}));
    } else if (action === "context") {
      return jsonResponse(await callLibreoffice("/tools/get_context_around_cursor_live", "POST", { chars }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 4: selection
// Actions: paragraph, range, delete, replace
// =============================================================================
server.tool(
  "selection",
  "Select and manipulate text ranges in the document.",
  {
    action: z.enum(["paragraph", "range", "delete", "replace"]).describe("The operation to perform"),
    n: z.number().int().optional().describe("Paragraph number (1-indexed) for 'paragraph' action"),
    start: z.number().int().optional().describe("Starting character position (0-indexed) for 'range' action"),
    end: z.number().int().optional().describe("Ending character position (exclusive) for 'range' action"),
    text: z.string().optional().describe("Replacement text for 'replace' action")
  },
  async ({ action, n, start, end, text }) => {
    if (action === "paragraph") {
      if (n === undefined) return jsonResponse({ error: "Action 'paragraph' requires parameter 'n'" });
      return jsonResponse(await callLibreoffice("/tools/select_paragraph_live", "POST", { n }));
    } else if (action === "range") {
      if (start === undefined || end === undefined) return jsonResponse({ error: "Action 'range' requires parameters 'start' and 'end'" });
      return jsonResponse(await callLibreoffice("/tools/select_text_range_live", "POST", { start, end }));
    } else if (action === "delete") {
      return jsonResponse(await callLibreoffice("/tools/delete_selection_live", "POST", {}));
    } else if (action === "replace") {
      if (text === undefined) return jsonResponse({ error: "Action 'replace' requires parameter 'text'" });
      return jsonResponse(await callLibreoffice("/tools/replace_selection_live", "POST", { text }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 5: search
// Actions: find, replace, replace_all
// =============================================================================
server.tool(
  "search",
  "Find and replace text in the document. Track Changes aware - skips tracked deletions.",
  {
    action: z.enum(["find", "replace", "replace_all"]).describe("The operation to perform"),
    query: z.string().optional().describe("Text to search for ('find' action)"),
    old: z.string().optional().describe("Text to find ('replace' and 'replace_all' actions)"),
    new: z.string().optional().describe("Replacement text ('replace' and 'replace_all' actions)")
  },
  async ({ action, query, old: oldText, new: newText }) => {
    if (action === "find") {
      if (!query) return jsonResponse({ error: "Action 'find' requires parameter 'query'" });
      return jsonResponse(await callLibreoffice("/tools/find_text_live", "POST", { query }));
    } else if (action === "replace") {
      if (!oldText || !newText) return jsonResponse({ error: "Action 'replace' requires parameters 'old' and 'new'" });
      return jsonResponse(await callLibreoffice("/tools/find_and_replace_live", "POST", { old: oldText, new: newText }));
    } else if (action === "replace_all") {
      if (!oldText || !newText) return jsonResponse({ error: "Action 'replace_all' requires parameters 'old' and 'new'" });
      return jsonResponse(await callLibreoffice("/tools/find_and_replace_all_live", "POST", { old: oldText, new: newText }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 6: track_changes
// Actions: status, enable, disable, list, accept, reject, accept_all, reject_all
// =============================================================================
server.tool(
  "track_changes",
  "Manage Track Changes / revision tracking in the document.",
  {
    action: z.enum(["status", "enable", "disable", "list", "accept", "reject", "accept_all", "reject_all"]).describe("The operation to perform"),
    index: z.number().int().optional().describe("Change index (0-based) for 'accept' and 'reject' actions"),
    show: z.boolean().default(true).describe("Whether to show tracked changes when enabling (default: true)")
  },
  async ({ action, index, show }) => {
    if (action === "status") {
      return jsonResponse(await callLibreoffice("/tools/get_track_changes_status_live", "POST", {}));
    } else if (action === "enable") {
      return jsonResponse(await callLibreoffice("/tools/set_track_changes_live", "POST", { enabled: true, show }));
    } else if (action === "disable") {
      return jsonResponse(await callLibreoffice("/tools/set_track_changes_live", "POST", { enabled: false, show }));
    } else if (action === "list") {
      return jsonResponse(await callLibreoffice("/tools/get_tracked_changes_live", "POST", {}));
    } else if (action === "accept") {
      if (index === undefined) return jsonResponse({ error: "Action 'accept' requires parameter 'index'" });
      return jsonResponse(await callLibreoffice("/tools/accept_tracked_change_live", "POST", { index }));
    } else if (action === "reject") {
      if (index === undefined) return jsonResponse({ error: "Action 'reject' requires parameter 'index'" });
      return jsonResponse(await callLibreoffice("/tools/reject_tracked_change_live", "POST", { index }));
    } else if (action === "accept_all") {
      return jsonResponse(await callLibreoffice("/tools/accept_all_changes_live", "POST", {}));
    } else if (action === "reject_all") {
      return jsonResponse(await callLibreoffice("/tools/reject_all_changes_live", "POST", {}));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 7: comments
// Actions: list, add
// =============================================================================
server.tool(
  "comments",
  "Manage document comments/annotations.",
  {
    action: z.enum(["list", "add"]).describe("The operation to perform"),
    text: z.string().optional().describe("Comment text for 'add' action"),
    author: z.string().default("Claude").describe("Author name for 'add' action (default: 'Claude')")
  },
  async ({ action, text, author }) => {
    if (action === "list") {
      return jsonResponse(await callLibreoffice("/tools/get_comments_live", "POST", {}));
    } else if (action === "add") {
      if (!text) return jsonResponse({ error: "Action 'add' requires parameter 'text'" });
      return jsonResponse(await callLibreoffice("/tools/add_comment_live", "POST", { text, author }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 8: save
// Actions: save, export
// =============================================================================
server.tool(
  "save",
  "Save and export documents.",
  {
    action: z.enum(["save", "export"]).describe("The operation to perform"),
    file_path: z.string().optional().describe("Path to save/export to"),
    export_format: z.enum(["pdf", "docx", "odt", "html", "txt"]).default("pdf").describe("Format for export. Options: pdf, docx, odt, html, txt")
  },
  async ({ action, file_path, export_format }) => {
    if (action === "save") {
      const data = {};
      if (file_path) data.file_path = file_path;
      return jsonResponse(await callLibreoffice("/tools/save_document_live", "POST", data));
    } else if (action === "export") {
      if (!file_path) return jsonResponse({ error: "Action 'export' requires parameter 'file_path'" });
      return jsonResponse(await callLibreoffice("/tools/export_document_live", "POST", {
        file_path,
        format: export_format
      }));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// =============================================================================
// CONSOLIDATED TOOL 9: text
// Actions: insert, format, style
// =============================================================================
server.tool(
  "text",
  "Insert, format text, or apply paragraph styles in the document.",
  {
    action: z.enum(["insert", "format", "style", "page_break"]).describe("The operation to perform"),
    content: z.string().optional().describe("Text to insert for 'insert' action"),
    bold: z.boolean().optional().describe("Set bold formatting (true/false) for 'format' action"),
    italic: z.boolean().optional().describe("Set italic formatting (true/false) for 'format' action"),
    underline: z.boolean().optional().describe("Set underline formatting (true/false) for 'format' action"),
    font_size: z.number().optional().describe("Font size in points for 'format' action"),
    font_name: z.string().optional().describe("Font family name for 'format' action"),
    style_name: z.string().optional().describe("Name of paragraph style for 'style' action (e.g. 'Heading 1', 'Heading 2', 'Title')"),
    paragraph_n: z.number().int().optional().describe("Optional paragraph number (1-indexed) to target with 'style' action directly"),
    page_break: z.boolean().optional().describe("Insert a page break before text (default: false)")
  },
  async ({ action, content, bold, italic, underline, font_size, font_name, style_name, paragraph_n, page_break }) => {
    if (action === "insert") {
      if (content === undefined) return jsonResponse({ error: "Action 'insert' requires parameter 'content' (text to insert)" });
      return jsonResponse(await callLibreoffice("/tools/insert_text_live", "POST", { text: content, page_break }));
    } else if (action === "page_break") {
      return jsonResponse(await callLibreoffice("/tools/insert_page_break_live", "POST", {}));
    } else if (action === "format") {
      const formatting = {};
      if (bold !== undefined) formatting.bold = bold;
      if (italic !== undefined) formatting.italic = italic;
      if (underline !== undefined) formatting.underline = underline;
      if (font_size !== undefined) formatting.font_size = font_size;
      if (font_name !== undefined) formatting.font_name = font_name;
      return jsonResponse(await callLibreoffice("/tools/format_text_live", "POST", { formatting }));
    } else if (action === "style") {
      if (!style_name) return jsonResponse({ error: "Action 'style' requires parameter 'style_name' (e.g. 'Heading 1', 'Heading 2')" });
      const data = { style_name };
      if (paragraph_n !== undefined) data.paragraph_n = paragraph_n;
      return jsonResponse(await callLibreoffice("/tools/format_paragraph_live", "POST", data));
    }
    return jsonResponse({ error: `Invalid action '${action}'` });
  }
);

// Connect via Stdio transport
const transport = new StdioServerTransport();
await server.connect(transport);
