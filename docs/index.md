# Google Docs MCP Server - Documentation Index

**Version**: 1.0.0
**Generated**: 2026-01-22
**Project Type**: MCP Server / Library

---

## Quick Links

| Document | Description |
|----------|-------------|
| [Architecture Overview](architecture.md) | System design, components, tool catalog |
| [Source Tree Analysis](source-tree.md) | File structure, code metrics, dependencies |
| [Developer Guide](developer-guide.md) | Contributing, adding tools, testing |
| [README](../README.md) | Setup instructions, usage examples |
| [Sample Tasks](../SAMPLE_TASKS.md) | 15 example workflows |
| [VS Code Guide](../vscode.md) | VS Code MCP extension integration |
| [CLAUDE.md](../CLAUDE.md) | AI assistant quick reference |

---

## Project Summary

The **Google Docs MCP Server** is a Model Context Protocol server that provides **92 tools** for interacting with Google Workspace APIs. It enables AI assistants to programmatically manage Google Docs, Sheets, Slides, Drive, and Gmail.

### Key Statistics

| Metric | Value |
|--------|-------|
| Total Tools | 92 |
| Source Files | 9 TypeScript files |
| Test Files | 2 test files |
| Total LOC | 9,180 |
| Dependencies | 4 runtime |
| API Integrations | 5 (Docs v1, Sheets v4, Slides v1, Drive v3, Gmail v1) |

---

## Tool Categories

### Document Operations (24 tools)
| Category | Count | Examples |
|----------|-------|----------|
| Access & Editing | 5 | `readGoogleDoc`, `appendToGoogleDoc`, `insertText` |
| Formatting | 3 | `applyTextStyle`, `applyParagraphStyle`, `formatMatchingText` |
| Structure | 7 | `insertTable`, `insertPageBreak`, `insertImageFromUrl` |
| Comments | 6 | `listComments`, `addComment`, `replyToComment` |
| Enhanced | 4 | `createFormattedDocument`, `replaceDocumentContent` |

### Google Sheets (8 tools)
- `readSpreadsheet` - Read data from range
- `writeSpreadsheet` - Write data to range
- `appendSpreadsheetRows` - Append rows
- `clearSpreadsheetRange` - Clear range values
- `getSpreadsheetInfo` - Get spreadsheet metadata
- `addSpreadsheetSheet` - Add new sheet/tab
- `createSpreadsheet` - Create new spreadsheet
- `listGoogleSheets` - List spreadsheets

### Google Drive (13 tools)
- Document discovery: `listGoogleDocs`, `searchGoogleDocs`, `getRecentGoogleDocs`
- Document info: `getDocumentInfo`
- Folder management: `createFolder`, `listFolderContents`, `listAllFolders`, `getFolderInfo`
- File operations: `moveFile`, `copyFile`, `renameFile`, `deleteFile`
- Document creation: `createDocument`

### Google Slides (16 tools)
- Presentation: `getPresentation`, `listSlides`, `getSlide`, `mapSlide`, `createPresentation`
- Slides: `addSlide`, `duplicateSlide`, `deleteSlide`, `moveSlide`
- Elements: `addTextBox`, `addShape`, `addImage`, `addTable`, `deleteElement`
- Content: `insertTextInElement`, `updateSpeakerNotes`

### Gmail (34 tools)
| Subcategory | Count | Examples |
|-------------|-------|----------|
| Core Email | 7 | `send_email`, `read_email`, `search_emails`, `draft_email` |
| Labels | 5 | `list_email_labels`, `create_label`, `get_or_create_label` |
| Batch | 2 | `batch_modify_emails`, `batch_delete_emails` |
| Filters | 5 | `create_filter`, `list_filters`, `create_filter_from_template` |
| Threads | 4 | `get_thread`, `list_threads`, `reply_to_email`, `forward_email` |
| Actions | 5 | `trash_email`, `archive_email`, `mark_as_read` |
| Drafts | 5 | `list_drafts`, `get_draft`, `update_draft`, `send_draft` |
| Profile | 1 | `get_user_profile` |

---

## Architecture Highlights

```
MCP Client (Claude/VS Code)
         │
    MCP Protocol (stdio)
         │
    FastMCP Server (92 tools)
         │
    ┌────┴────┬────────┬────────┬────────┐
    │         │        │        │        │
  Docs    Sheets   Slides   Drive    Gmail
  API v1   API v4   API v1   API v3   API v1
```

### Authentication Methods
1. **OAuth2 (User Consent)** - Interactive flow, stores token locally
2. **Service Account** - Domain-wide delegation for enterprise use

---

## Technology Stack

| Layer | Technology |
|-------|------------|
| Runtime | Node.js 18+ |
| Language | TypeScript 5.x (ES2022) |
| MCP Framework | FastMCP 3.24.0 |
| API Client | googleapis 148.x |
| Validation | Zod 3.x |
| Auth | google-auth-library 9.x |

---

## Source Files

| File | Lines | Purpose |
|------|-------|---------|
| `server.ts` | 5,301 | 92 tool definitions |
| `googleDocsApiHelpers.ts` | 827 | Docs API helpers |
| `googleGmailApiHelpers.ts` | 802 | Gmail API helpers |
| `googleSlidesApiHelpers.ts` | 595 | Slides API helpers |
| `googleSheetsApiHelpers.ts` | 427 | Sheets API helpers |
| `types.ts` | 386 | Zod schemas |
| `gmailFilterManager.ts` | 321 | Filter management |
| `gmailLabelManager.ts` | 299 | Label management |
| `auth.ts` | 226 | Authentication |
| **Total** | **9,180** | |

---

## Known Limitations

1. **Comment anchoring** - Programmatic comments appear in list but not anchored to text
2. **Resolved status** - May not persist in Google Docs UI
3. **editTableCell** - Not implemented (complex cell indexing)
4. **findElement** - Not implemented
5. **fixListFormatting** - Experimental, unreliable
6. **Gmail first use** - Delete existing tokens after adding Gmail scope

---

## Getting Started

### For Users
1. Follow [README.md](../README.md) setup instructions
2. Review [SAMPLE_TASKS.md](../SAMPLE_TASKS.md) for usage examples
3. See [vscode.md](../vscode.md) for VS Code integration
4. Reference [CLAUDE.md](../CLAUDE.md) for tool quick reference

### For Developers
1. Read [Architecture Overview](architecture.md) for system design
2. Review [Source Tree Analysis](source-tree.md) for code structure
3. Follow [Developer Guide](developer-guide.md) for contributing

---

## File Map

```
mcp-googledocs-server/
├── docs/                    # ← YOU ARE HERE
│   ├── index.md            # This file
│   ├── architecture.md     # System design (92 tools)
│   ├── source-tree.md      # Code structure (9,180 LOC)
│   └── developer-guide.md  # Contributing guide
├── src/                     # TypeScript source (9 files)
├── tests/                   # Test files
├── README.md               # Main documentation
├── SAMPLE_TASKS.md         # Example workflows
├── CLAUDE.md               # AI quick reference
└── vscode.md               # VS Code guide
```

---

## Document Generation Info

| Field | Value |
|-------|-------|
| Generated by | BMAD Document Project Workflow v1.2.0 |
| Scan Level | Exhaustive |
| Files Analyzed | All 9 source files |
| Tools Documented | 92 |
| Generation Date | 2026-01-22 |
