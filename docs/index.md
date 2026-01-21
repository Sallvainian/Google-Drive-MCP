# Google Docs MCP Server - Documentation Index

**Version**: 1.0.0
**Generated**: 2026-01-20
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

---

## Project Summary

The **Google Docs MCP Server** is a Model Context Protocol server that provides **42 tools** for interacting with Google Workspace APIs. It enables AI assistants to programmatically manage Google Docs, Sheets, and Drive.

### Key Statistics

| Metric | Value |
|--------|-------|
| Total Tools | 42 |
| Source Files | 5 TypeScript files |
| Test Files | 2 test files |
| Total LOC | ~4,400 |
| Dependencies | 4 runtime, 2 dev |
| API Integrations | Google Docs v1, Sheets v4, Drive v3 |

---

## Tool Categories

### Document Operations (5 tools)
- `readGoogleDoc` - Read document content in text, JSON, or markdown format
- `listDocumentTabs` - List all tabs in a document
- `appendToGoogleDoc` - Append text to document end
- `insertText` - Insert text at specific position
- `deleteRange` - Delete content in a range

### Formatting (3 tools)
- `applyTextStyle` - Apply character formatting (bold, italic, colors)
- `applyParagraphStyle` - Apply paragraph formatting (alignment, spacing)
- `formatMatchingText` - Find and format specific text

### Document Structure (7 tools)
- `insertTable` - Create tables
- `insertPageBreak` - Insert page breaks
- `insertImageFromUrl` - Insert image from URL
- `insertLocalImage` - Upload and insert local image
- `editTableCell` - Edit table cells *(not implemented)*
- `fixListFormatting` - Auto-format lists *(experimental)*
- `findElement` - Find elements *(not implemented)*

### Comments (6 tools)
- `listComments` - List all comments
- `getComment` - Get comment with replies
- `addComment` - Add comment to text range
- `replyToComment` - Reply to comment
- `resolveComment` - Mark as resolved
- `deleteComment` - Delete comment

### Google Sheets (8 tools)
- `readSpreadsheet` - Read data from range
- `writeSpreadsheet` - Write data to range
- `appendSpreadsheetRows` - Append rows
- `clearSpreadsheetRange` - Clear range values
- `getSpreadsheetInfo` - Get spreadsheet metadata
- `addSpreadsheetSheet` - Add new sheet/tab
- `createSpreadsheet` - Create new spreadsheet
- `listGoogleSheets` - List spreadsheets

### Google Drive (14 tools)
- `listGoogleDocs` - List documents
- `searchGoogleDocs` - Search documents
- `getRecentGoogleDocs` - Get recent documents
- `getDocumentInfo` - Get document metadata
- `createFolder` - Create folder
- `listFolderContents` - List folder contents
- `listAllFolders` - Recursive folder listing
- `getFolderInfo` - Get folder metadata
- `moveFile` - Move file/folder
- `copyFile` - Copy file
- `renameFile` - Rename file/folder
- `deleteFile` - Delete file/folder
- `createDocument` - Create new document
- `createFromTemplate` - Create from template

### Enhanced Formatting (3 tools)
- `createFormattedDocument` - Create document with structured formatting
- `insertFormattedContent` - Insert formatted content at index
- `replaceDocumentContent` - Replace entire document content

---

## Architecture Highlights

```
MCP Client (Claude/VS Code)
         │
    MCP Protocol (stdio)
         │
    FastMCP Server
         │
    ┌────┴────┐
    │         │
Google APIs   Auth Module
    │         │
    └────┬────┘
         │
   OAuth2 / Service Account
```

### Authentication Methods
1. **OAuth2 (User Consent)** - Interactive flow, stores token locally
2. **Service Account** - Domain-wide delegation for enterprise use

---

## Technology Stack

| Layer | Technology |
|-------|------------|
| Runtime | Node.js 18+ |
| Language | TypeScript 5.x |
| MCP Framework | FastMCP 0.3.x |
| API Client | googleapis 148.x |
| Validation | Zod 3.x |
| Auth | google-auth-library 9.x |

---

## Known Limitations

1. **Comment anchoring** - Programmatic comments appear in list but not anchored to text
2. **Resolved status** - May not persist in Google Docs UI
3. **editTableCell** - Not implemented (complex cell indexing)
4. **findElement** - Not implemented
5. **fixListFormatting** - Experimental, unreliable

---

## Getting Started

### For Users
1. Follow [README.md](../README.md) setup instructions
2. Review [SAMPLE_TASKS.md](../SAMPLE_TASKS.md) for usage examples
3. See [vscode.md](../vscode.md) for VS Code integration

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
│   ├── architecture.md     # System design
│   ├── source-tree.md      # Code structure
│   └── developer-guide.md  # Contributing guide
├── src/                     # TypeScript source
├── tests/                   # Test files
├── README.md               # Main documentation
├── SAMPLE_TASKS.md         # Example workflows
├── CLAUDE.md               # AI instructions
└── vscode.md               # VS Code guide
```

---

## Document Generation Info

| Field | Value |
|-------|-------|
| Generated by | BMAD Document Project Workflow v1.2.0 |
| Scan Level | Exhaustive |
| Files Analyzed | All source files |
| Generation Date | 2026-01-20 |
