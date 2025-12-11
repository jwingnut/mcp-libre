# Known Issues and Roadmap

## Known Issues

### Critical Issues

#### 1. Track Changes Awareness (SPEC CREATED)
**Status:** Specification complete, ready for implementation
**Location:** `.spec-workflow/specs/track-changes-awareness/`

**Problem:** When Track Changes is enabled in LibreOffice Writer, the search/replace tools operate on all text including tracked deletions (text marked for deletion but not yet accepted/rejected). This causes:

- `find_text` returns matches in deleted text that appears crossed out
- `find_and_replace` may find the same text repeatedly because it's matching deleted text that won't actually be replaced
- `find_and_replace_all` reports incorrect replacement counts
- Tools cannot distinguish between "visible" document content and tracked deletions

**Impact:** High - Makes automated editing unreliable when Track Changes is enabled

**Solution:** 17 tasks defined in specification:
- Add 7 new Track Changes management tools
- Modify existing search/replace tools to skip tracked deletions
- Add Track Changes status to document info

#### 2. Comment Reply Feature Missing
**Status:** Not yet implemented
**Problem:** While `add_comment` and `get_comments` tools exist, there's no way to reply to existing comments.

**Workaround:** None currently available

### Minor Issues

#### 3. Calc/Impress/Draw Limited Support
**Status:** By design
**Problem:** Most tools only work with Writer documents. Other document types return "not supported" errors.

**Affected Tools:** All 25 tools except `create_document`, `get_document_info`, `list_open_documents`, `check_libreoffice_status`

#### 4. No Undo/Redo Support
**Status:** Not implemented
**Problem:** No tools to programmatically undo or redo changes.

**Workaround:** User can use Ctrl+Z in LibreOffice manually

#### 5. Table Support Limited
**Status:** Not implemented
**Problem:** No dedicated tools for creating, editing, or navigating tables.

**Workaround:** Insert text containing table formatting manually

---

## Roadmap

### Phase 1: Track Changes Awareness (Priority: High)
**Spec:** `track-changes-awareness`
**Tasks:** 17
**Estimated effort:** Medium

Implements full Track Changes awareness including:
- Status detection (recording, showing, pending count)
- Management tools (accept/reject individual or all changes)
- Search/replace modifications to skip tracked deletions
- Document info integration

### Phase 2: Enhanced Comment Features (Priority: Medium)
**Spec:** Not yet created

Potential features:
- Reply to existing comments
- Delete comments
- Resolve/unresolve comments
- Get comments for specific text range
- Navigate between comments

### Phase 3: Advanced Navigation (Priority: Medium)
**Spec:** Not yet created

Potential features:
- Find/goto next heading
- Navigate by page number
- Navigate by section
- Bookmark support
- Cross-reference support

### Phase 4: Table Support (Priority: Low)
**Spec:** Not yet created

Potential features:
- Create table
- Get table content
- Edit table cells
- Add/remove rows and columns
- Table formatting

### Phase 5: Calc & Spreadsheet Support (Priority: Low)
**Spec:** Not yet created

Potential features:
- Read cell values
- Write cell values
- Get/set formulas
- Navigate sheets
- Range selection

---

## Contributing

To contribute to any of these features:

1. Check if a spec exists in `.spec-workflow/specs/`
2. If no spec exists, create one using the spec-workflow MCP tools
3. Follow the requirements → design → tasks → implementation workflow
4. Use `log-implementation` to document changes
5. Submit PR with all tests passing

---

## Version History

### v0.2.0 (Current)
- Added 15 enhanced editing tools (paragraphs, navigation, selection, search/replace)
- Fixed build script to include .xcu files
- Total tools: 25

### v0.1.0
- Initial release with 10 basic tools
- Document management, text insertion, formatting, comments
