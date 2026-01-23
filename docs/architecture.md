# Architecture Overview

## Project: Google Docs MCP Server

**Version**: 1.0.0
**Type**: Library (MCP Server)
**Framework**: FastMCP 3.24.0
**Language**: TypeScript
**Generated**: 2026-01-22

---

## Executive Summary

The Google Docs MCP Server is a Model Context Protocol (MCP) server that provides **92 tools** for interacting with Google Workspace APIs (Docs, Sheets, Slides, Drive, and Gmail). It enables AI assistants like Claude to programmatically read, write, format, and manage Google documents, spreadsheets, presentations, and email.

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        MCP Client                               │
│              (Claude Desktop, VS Code, etc.)                    │
└────────────────────────────┬────────────────────────────────────┘
                             │ MCP Protocol (stdio)
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                      FastMCP Server                             │
│                     (src/server.ts)                             │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │                    92 Tool Handlers                      │   │
│  │  ┌──────────┬──────────┬──────────┬──────────┬────────┐ │   │
│  │  │  Docs    │  Sheets  │  Slides  │  Drive   │ Gmail  │ │   │
│  │  │  (24)    │   (8)    │  (16)    │  (13)    │  (34)  │ │   │
│  │  └──────────┴──────────┴──────────┴──────────┴────────┘ │   │
│  └─────────────────────────────────────────────────────────┘   │
└────────────────────────────┬────────────────────────────────────┘
                             │
         ┌───────────────────┼───────────────────┐
         ▼                   ▼                   ▼
┌─────────────────┐ ┌─────────────────┐ ┌─────────────────┐
│    auth.ts      │ │  API Helpers    │ │    types.ts     │
│  OAuth2/Service │ │                 │ │  Zod Schemas    │
│  Account Auth   │ │ ┌─────────────┐ │ │  & Type Defs    │
└─────────────────┘ │ │ Docs Helper │ │ └─────────────────┘
                    │ │ Sheets Help │ │
                    │ │ Slides Help │ │
                    │ │ Gmail Help  │ │
                    │ │ Label Mgr   │ │
                    │ │ Filter Mgr  │ │
                    │ └─────────────┘ │
                    └────────┬────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Google APIs                                │
│  ┌──────────┬──────────┬──────────┬──────────┬────────────┐    │
│  │ Docs v1  │Sheets v4 │Slides v1 │ Drive v3 │  Gmail v1  │    │
│  └──────────┴──────────┴──────────┴──────────┴────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

---

## Component Breakdown

### 1. Entry Points

| File | Purpose |
|------|---------|
| `index.js` | Node.js entry point, imports compiled `dist/server.js` |
| `src/server.ts` | Main server implementation with all 92 tool definitions |

### 2. Core Modules

| Module | Lines | Responsibility |
|--------|-------|----------------|
| `server.ts` | 5,301 | Main server, 92 tool definitions, MCP protocol handling |
| `auth.ts` | 226 | OAuth2 flow, service account JWT, token management |
| `types.ts` | 386 | Zod schemas, TypeScript types, parameter validation |

### 3. API Helper Modules

| Module | Lines | Responsibility |
|--------|-------|----------------|
| `googleDocsApiHelpers.ts` | 827 | Document operations, text ranges, batch updates, tabs |
| `googleSheetsApiHelpers.ts` | 427 | A1 notation, range operations, cell formatting |
| `googleSlidesApiHelpers.ts` | 595 | EMU conversion, element positioning, batch updates |
| `googleGmailApiHelpers.ts` | 802 | Email MIME handling, attachments, threads, drafts |
| `gmailLabelManager.ts` | 299 | Label CRUD, system labels, label resolution |
| `gmailFilterManager.ts` | 321 | Filter CRUD, template-based filters |

**Total:** 9,180 lines of TypeScript

---

## Tool Catalog (92 Tools)

### Document Access & Editing (5 tools)
| Tool | Description |
|------|-------------|
| `readGoogleDoc` | Read document content (text/json/markdown) |
| `listDocumentTabs` | List all tabs in a document |
| `appendToGoogleDoc` | Append text to document end |
| `insertText` | Insert text at specific position |
| `deleteRange` | Delete content in range |

### Formatting & Styling (3 tools)
| Tool | Description |
|------|-------------|
| `applyTextStyle` | Apply character formatting (bold, italic, colors) |
| `applyParagraphStyle` | Apply paragraph formatting (alignment, spacing) |
| `formatMatchingText` | Find and format specific text |

### Document Structure (7 tools)
| Tool | Description |
|------|-------------|
| `insertTable` | Create a new table |
| `editTableCell` | Edit table cell |
| `insertPageBreak` | Insert page break |
| `insertImageFromUrl` | Insert image from URL |
| `insertLocalImage` | Upload and insert local image |
| `fixListFormatting` | Auto-format lists *(EXPERIMENTAL)* |
| `findElement` | Find elements *(NOT IMPLEMENTED)* |

### Comment Management (6 tools)
| Tool | Description |
|------|-------------|
| `listComments` | List all comments |
| `getComment` | Get comment with replies |
| `addComment` | Add comment to text range |
| `replyToComment` | Reply to existing comment |
| `resolveComment` | Mark comment as resolved |
| `deleteComment` | Delete a comment |

### Google Sheets (8 tools)
| Tool | Description |
|------|-------------|
| `readSpreadsheet` | Read data from range |
| `writeSpreadsheet` | Write data to range |
| `appendSpreadsheetRows` | Append rows to sheet |
| `clearSpreadsheetRange` | Clear values from range |
| `getSpreadsheetInfo` | Get spreadsheet metadata |
| `addSpreadsheetSheet` | Add new sheet/tab |
| `createSpreadsheet` | Create new spreadsheet |
| `listGoogleSheets` | List spreadsheets |

### Google Drive (13 tools)
| Tool | Description |
|------|-------------|
| `listGoogleDocs` | List documents |
| `searchGoogleDocs` | Search documents |
| `getRecentGoogleDocs` | Get recently modified docs |
| `getDocumentInfo` | Get document metadata |
| `createFolder` | Create new folder |
| `listFolderContents` | List folder contents |
| `listAllFolders` | Recursive folder listing |
| `getFolderInfo` | Get folder metadata |
| `moveFile` | Move file/folder |
| `copyFile` | Copy file |
| `renameFile` | Rename file/folder |
| `deleteFile` | Delete file/folder |
| `createDocument` | Create new document |

### Enhanced Formatting (4 tools)
| Tool | Description |
|------|-------------|
| `createFormattedDocument` | Create doc with structured formatting |
| `insertFormattedContent` | Insert formatted content at index |
| `replaceDocumentContent` | Replace entire document content |
| `createFromTemplate` | Create document from template |

### Google Slides (16 tools)
| Tool | Description |
|------|-------------|
| `getPresentation` | Get presentation metadata and content |
| `listSlides` | List all slides in presentation |
| `getSlide` | Get detailed slide content |
| `mapSlide` | Analyze slide dimensions and layout |
| `createPresentation` | Create new presentation |
| `addSlide` | Add new slide |
| `duplicateSlide` | Duplicate existing slide |
| `addTextBox` | Add text box with styling |
| `addShape` | Add shape (rectangle, ellipse, etc.) |
| `addImage` | Add image from URL |
| `addTable` | Add table to slide |
| `deleteSlide` | Delete a slide |
| `deleteElement` | Delete element from slide |
| `updateSpeakerNotes` | Update speaker notes |
| `moveSlide` | Move slide to new position |
| `insertTextInElement` | Insert/replace text in element |

### Gmail (34 tools)

#### Core Email Operations (7)
| Tool | Description |
|------|-------------|
| `send_email` | Send email with attachments |
| `draft_email` | Create draft with attachments |
| `read_email` | Get message content, headers, attachments |
| `search_emails` | Search with Gmail syntax |
| `modify_email` | Add/remove labels |
| `delete_email` | Permanently delete (irreversible) |
| `download_attachment` | Save attachment to filesystem |

#### Label Management (5)
| Tool | Description |
|------|-------------|
| `list_email_labels` | Get all labels |
| `create_label` | Create new label |
| `update_label` | Rename or change visibility |
| `delete_label` | Remove user label |
| `get_or_create_label` | Idempotent label creation |

#### Batch Operations (2)
| Tool | Description |
|------|-------------|
| `batch_modify_emails` | Bulk label changes |
| `batch_delete_emails` | Bulk permanent delete |

#### Filter Management (5)
| Tool | Description |
|------|-------------|
| `create_filter` | Create custom filter |
| `list_filters` | Get all filters |
| `get_filter` | Get filter details |
| `delete_filter` | Remove filter |
| `create_filter_from_template` | Template-based filter creation |

#### Thread & Conversation (4)
| Tool | Description |
|------|-------------|
| `get_thread` | Full conversation with all messages |
| `list_threads` | Search/list threads |
| `reply_to_email` | Reply maintaining thread |
| `forward_email` | Forward to new recipients |

#### Message Actions (5)
| Tool | Description |
|------|-------------|
| `trash_email` | Move to trash (recoverable) |
| `untrash_email` | Restore from trash |
| `archive_email` | Remove from inbox only |
| `mark_as_read` | Mark message read |
| `mark_as_unread` | Mark message unread |

#### Draft Management (5)
| Tool | Description |
|------|-------------|
| `list_drafts` | Get all drafts |
| `get_draft` | Get draft content |
| `update_draft` | Modify draft |
| `delete_draft` | Delete draft |
| `send_draft` | Send existing draft |

#### Profile (1)
| Tool | Description |
|------|-------------|
| `get_user_profile` | Get authenticated email address |

---

## Authentication Architecture

### Method 1: OAuth2 (User Consent)

```
1. Server starts
2. Checks for existing token.json
3. If no token:
   a. Load credentials.json
   b. Start local HTTP server on port 3000
   c. Generate OAuth URL
   d. User visits URL, grants access
   e. Callback receives authorization code
   f. Token saved to token.json
4. Initialize Google API clients with token
```

### Method 2: Service Account (Domain-Wide Delegation)

```
Environment Variables:
- SERVICE_ACCOUNT_PATH: Path to service account key file
- GOOGLE_IMPERSONATE_USER: Email of user to impersonate

1. Server starts
2. Detects SERVICE_ACCOUNT_PATH
3. Loads service account credentials
4. Creates JWT client with impersonation
5. Initialize Google API clients
```

### Required OAuth Scopes

```typescript
const SCOPES = [
  'https://www.googleapis.com/auth/documents',      // Docs read/write
  'https://www.googleapis.com/auth/drive',          // Drive full access
  'https://www.googleapis.com/auth/spreadsheets',   // Sheets read/write
  'https://www.googleapis.com/auth/presentations',  // Slides read/write
  'https://mail.google.com/',                       // Gmail full access
];
```

---

## Data Flow

### Read Operation
```
Client Request → FastMCP → getDocsClient() → Docs API → Parse Response → Format Output → Client
```

### Write Operation
```
Client Request → Validate Params (Zod) → Build Request → executeBatchUpdate → Docs API → Success Response
```

### Error Handling
```
API Error → Check Error Code → Map to UserError → Return Friendly Message
- 400: Invalid request details
- 404: "Document/Resource not found"
- 403: "Permission denied"
- Other: Original error message
```

---

## Key Design Decisions

### 1. Monolithic Server Architecture
**Decision:** All 92 tools in a single server process.
**Rationale:** Simplified deployment, shared authentication state, lower memory footprint.

### 2. Zod Schema Validation
**Decision:** Use Zod for all tool parameter validation.
**Rationale:** Runtime type safety, automatic TypeScript inference, clear error messages.

### 3. Helper Module Separation
**Decision:** Separate helper modules per Google API.
**Rationale:** Separation of concerns, testable units, reusable across tools.

### 4. Dual Authentication Support
**Decision:** Support both OAuth2 and Service Account.
**Rationale:** OAuth2 for individual users, Service Account for enterprise deployment.

---

## Known Limitations

1. **Comment Anchoring**: Programmatically created comments appear in "All Comments" but aren't visibly anchored to text
2. **Resolved Status**: Comment resolution may not persist in Google Docs UI
3. **editTableCell**: May have edge cases with merged cells
4. **findElement**: Not implemented
5. **fixListFormatting**: Experimental, may not work reliably
6. **First Gmail use**: After adding Gmail scope, delete existing tokens and re-authenticate

---

## Dependencies

### Runtime Dependencies
| Package | Version | Purpose |
|---------|---------|---------|
| fastmcp | ^3.24.0 | MCP server framework |
| googleapis | ^148.0.0 | Google API client |
| google-auth-library | ^9.15.1 | OAuth2 and service account auth |
| zod | ^3.24.2 | Schema validation |

---

## Security Considerations

1. **Credential Protection**: `credentials.json` and `token.json` must never be committed
2. **Token Storage**: Tokens stored in plain text (consider keychain for production)
3. **Scope Minimization**: Server requests only necessary OAuth scopes
4. **Error Sanitization**: API errors sanitized before returning to client
5. **Gmail Access**: Uses full Gmail scope for complete functionality
