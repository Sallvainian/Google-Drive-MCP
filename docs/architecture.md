# Architecture Overview

## Project: Google Docs MCP Server

**Version**: 1.0.0
**Type**: Library (MCP Server)
**Framework**: FastMCP
**Language**: TypeScript
**Generated**: 2026-01-20

---

## Executive Summary

The Google Docs MCP Server is a Model Context Protocol (MCP) server that provides 42 tools for interacting with Google Workspace APIs (Docs, Sheets, Drive). It enables AI assistants like Claude to programmatically read, write, format, and manage Google documents and spreadsheets.

---

## System Architecture

```
                                    +-------------------+
                                    |   MCP Client      |
                                    | (Claude Desktop,  |
                                    |  VS Code, etc.)   |
                                    +--------+----------+
                                             |
                                    MCP Protocol (stdio)
                                             |
                                    +--------v----------+
                                    |   FastMCP Server  |
                                    |   (server.ts)     |
                                    +--------+----------+
                                             |
              +------------------+-----------+-----------+------------------+
              |                  |                       |                  |
     +--------v--------+ +------v-------+       +-------v-------+ +--------v--------+
     | Google Docs     | | Google Sheets|       | Google Drive  | | Comment System  |
     | API Helpers     | | API Helpers  |       | Operations    | | (Drive API v3)  |
     +--------+--------+ +------+-------+       +-------+-------+ +--------+--------+
              |                  |                       |                  |
              +------------------+-----------+-----------+------------------+
                                             |
                                    +--------v----------+
                                    |   Auth Module     |
                                    |   (auth.ts)       |
                                    | - OAuth2 Flow     |
                                    | - Service Account |
                                    +--------+----------+
                                             |
                                    +--------v----------+
                                    |  Google APIs      |
                                    | - Docs API v1     |
                                    | - Sheets API v4   |
                                    | - Drive API v3    |
                                    +-------------------+
```

---

## Component Breakdown

### 1. Entry Points

| File | Purpose |
|------|---------|
| `index.js` | Node.js entry point, imports compiled `dist/server.js` |
| `src/server.ts` | Main server implementation with all 42 tool definitions |

### 2. Core Modules

#### `src/server.ts` (~3000 lines)
- FastMCP server initialization
- All 42 MCP tool definitions
- Google API client initialization
- Error handling and logging
- Markdown conversion for document output

#### `src/auth.ts`
- OAuth2 authentication flow
- Service account authentication (domain-wide delegation)
- Token management (load/store/refresh)
- Credential validation

#### `src/types.ts`
- Zod schemas for parameter validation
- Type definitions for tool parameters
- Color validation (hex to RGB conversion)
- Shared parameter types (DocumentId, Range, TextStyle, etc.)

#### `src/googleDocsApiHelpers.ts`
- Text range finding and manipulation
- Batch update execution
- Style request builders (text, paragraph)
- Tab navigation helpers
- Image upload utilities
- Table creation

#### `src/googleSheetsApiHelpers.ts`
- A1 notation parsing and conversion
- Range operations (read, write, append, clear)
- Spreadsheet metadata retrieval
- Sheet/tab management
- Cell formatting

### 3. Test Suite

| File | Coverage |
|------|----------|
| `tests/helpers.test.js` | `findTextRange` function tests |
| `tests/types.test.js` | Color validation and conversion tests |

---

## Tool Catalog (42 Tools)

### Document Access & Editing (5 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `readGoogleDoc` | Read document content | documentId, format (text/json/markdown), tabId |
| `listDocumentTabs` | List all tabs in a document | documentId, includeContent |
| `appendToGoogleDoc` | Append text to document end | documentId, textToAppend, tabId |
| `insertText` | Insert text at specific position | documentId, textToInsert, index, tabId |
| `deleteRange` | Delete content in range | documentId, startIndex, endIndex, tabId |

### Formatting & Styling (3 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `applyTextStyle` | Apply character formatting | documentId, target, style (bold, italic, colors, font) |
| `applyParagraphStyle` | Apply paragraph formatting | documentId, target, style (alignment, spacing, named styles) |
| `formatMatchingText` | Find and format text | documentId, textToFind, matchInstance, formatting options |

### Document Structure (7 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `insertTable` | Create a new table | documentId, rows, columns, index |
| `editTableCell` | Edit table cell (NOT IMPLEMENTED) | - |
| `insertPageBreak` | Insert page break | documentId, index |
| `insertImageFromUrl` | Insert image from URL | documentId, imageUrl, index, width, height |
| `insertLocalImage` | Upload and insert local image | documentId, localImagePath, index |
| `fixListFormatting` | Auto-detect and format lists (EXPERIMENTAL) | documentId, range |
| `findElement` | Find elements by criteria (NOT IMPLEMENTED) | - |

### Comment Management (6 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `listComments` | List all comments | documentId |
| `getComment` | Get comment with replies | documentId, commentId |
| `addComment` | Add comment to text range | documentId, startIndex, endIndex, commentText |
| `replyToComment` | Reply to existing comment | documentId, commentId, replyText |
| `resolveComment` | Mark comment as resolved | documentId, commentId |
| `deleteComment` | Delete a comment | documentId, commentId |

### Google Sheets (8 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `readSpreadsheet` | Read data from range | spreadsheetId, range (A1 notation) |
| `writeSpreadsheet` | Write data to range | spreadsheetId, range, values (2D array) |
| `appendSpreadsheetRows` | Append rows to sheet | spreadsheetId, range, values |
| `clearSpreadsheetRange` | Clear values from range | spreadsheetId, range |
| `getSpreadsheetInfo` | Get spreadsheet metadata | spreadsheetId |
| `addSpreadsheetSheet` | Add new sheet/tab | spreadsheetId, sheetTitle |
| `createSpreadsheet` | Create new spreadsheet | title, parentFolderId, initialData |
| `listGoogleSheets` | List spreadsheets | maxResults, query, orderBy |

### Google Drive (14 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `listGoogleDocs` | List documents | maxResults, query, orderBy |
| `searchGoogleDocs` | Search documents | searchQuery, searchIn, modifiedAfter |
| `getRecentGoogleDocs` | Get recently modified docs | maxResults, daysBack |
| `getDocumentInfo` | Get document metadata | documentId |
| `createFolder` | Create new folder | name, parentFolderId |
| `listFolderContents` | List folder contents | folderId, includeSubfolders, includeFiles |
| `listAllFolders` | Recursive folder listing | folderId, maxDepth |
| `getFolderInfo` | Get folder metadata | folderId |
| `moveFile` | Move file/folder | fileId, newParentId |
| `copyFile` | Copy file | fileId, newName, parentFolderId |
| `renameFile` | Rename file/folder | fileId, newName |
| `deleteFile` | Delete file/folder | fileId, skipTrash |
| `createDocument` | Create new document | title, parentFolderId, initialContent |
| `createFromTemplate` | Create from template | templateId, newTitle, replacements |

### Enhanced Formatting (3 tools)
| Tool | Description | Key Parameters |
|------|-------------|----------------|
| `createFormattedDocument` | Create doc with structured formatting | title, content (array of formatted sections) |
| `insertFormattedContent` | Insert formatted content at index | documentId, index, content |
| `replaceDocumentContent` | Replace entire document content | documentId, content |

---

## Authentication Architecture

### Method 1: OAuth2 (User Consent)

```
1. Server starts
2. Checks for existing token.json
3. If no token:
   a. Load credentials.json
   b. Generate OAuth URL
   c. User visits URL, grants access
   d. User pastes authorization code
   e. Token saved to token.json
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
- 404: "Document/Resource not found"
- 403: "Permission denied"
- Other: Original error message
```

---

## Known Limitations

1. **Comment Anchoring**: Programmatically created comments appear in "All Comments" but aren't visibly anchored to text
2. **Resolved Status**: Comment resolution may not persist in Google Docs UI
3. **editTableCell**: Not implemented (complex cell index calculation)
4. **findElement**: Not implemented
5. **fixListFormatting**: Experimental, may not work reliably
6. **Converted Documents**: Some converted Word documents may not support all operations

---

## Dependencies

### Runtime Dependencies
| Package | Version | Purpose |
|---------|---------|---------|
| fastmcp | ^0.3.3 | MCP server framework |
| googleapis | ^148.0.0 | Google API client |
| google-auth-library | ^9.15.0 | OAuth2 and service account auth |
| zod | ^3.24.1 | Schema validation |

### Development Dependencies
| Package | Version | Purpose |
|---------|---------|---------|
| typescript | ^5.8.3 | TypeScript compiler |
| @types/node | ^22.14.1 | Node.js type definitions |

---

## Configuration Files

| File | Purpose |
|------|---------|
| `package.json` | Project metadata and dependencies |
| `tsconfig.json` | TypeScript compiler configuration |
| `credentials.json` | OAuth2 client credentials (user-provided) |
| `token.json` | OAuth2 access/refresh tokens (generated) |

---

## Build & Run

```bash
# Install dependencies
npm install

# Build TypeScript
npm run build

# Run server (first-time auth required)
node dist/server.js

# Entry point (production)
node index.js
```

---

## Security Considerations

1. **Credential Protection**: `credentials.json` and `token.json` must never be committed
2. **Token Storage**: Tokens stored in plain text (consider keychain for production)
3. **Scope Minimization**: Server requests only necessary OAuth scopes
4. **Error Sanitization**: API errors sanitized before returning to client
