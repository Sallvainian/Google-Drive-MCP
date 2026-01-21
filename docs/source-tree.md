# Source Tree Analysis

## Project: Google Docs MCP Server
**Generated**: 2026-01-20

---

## Directory Structure

```
mcp-googledocs-server/
├── src/                          # TypeScript source files
│   ├── server.ts                 # Main MCP server (42 tools)
│   ├── auth.ts                   # Authentication module
│   ├── types.ts                  # Type definitions & Zod schemas
│   ├── googleDocsApiHelpers.ts   # Google Docs API helper functions
│   ├── googleSheetsApiHelpers.ts # Google Sheets API helper functions
│   └── backup/                   # Backup files
│       ├── auth.ts.bak
│       └── server.ts.bak
│
├── tests/                        # Test files
│   ├── helpers.test.js           # Tests for helper functions
│   └── types.test.js             # Tests for type validation
│
├── dist/                         # Compiled JavaScript (generated)
│   ├── server.js
│   ├── auth.js
│   ├── types.js
│   ├── googleDocsApiHelpers.js
│   └── googleSheetsApiHelpers.js
│
├── docs/                         # Generated documentation
│   ├── index.md                  # Master index
│   ├── architecture.md           # Architecture overview
│   ├── source-tree.md            # This file
│   └── project-scan-report.json  # Scan state file
│
├── assets/                       # Static assets
│   └── google.docs.mcp.1.gif     # Demo animation
│
├── pages/                        # Additional pages
│   └── pages.md                  # (empty)
│
├── node_modules/                 # Dependencies (not tracked)
│
├── index.js                      # Node.js entry point
├── package.json                  # Project manifest
├── package-lock.json             # Dependency lock file
├── tsconfig.json                 # TypeScript configuration
├── README.md                     # Main documentation
├── SAMPLE_TASKS.md               # Example workflows
├── CLAUDE.md                     # AI assistant instructions
├── vscode.md                     # VS Code integration guide
├── credentials.json              # OAuth credentials (user-provided, gitignored)
└── token.json                    # OAuth token (generated, gitignored)
```

---

## File Analysis

### Source Files

#### `src/server.ts` (Main Server)
**Lines**: ~3000
**Purpose**: Core MCP server implementation

**Key Sections**:
1. **Imports & Setup** (Lines 1-30)
   - FastMCP, Zod, googleapis imports
   - Global client variables

2. **Initialization** (Lines 30-115)
   - `initializeGoogleClient()`: Lazy initialization
   - `getDocsClient()`, `getDriveClient()`, `getSheetsClient()`: Client getters
   - Process-level error handlers

3. **Helper Functions** (Lines 115-275)
   - `convertDocsJsonToMarkdown()`: JSON to Markdown converter
   - `convertParagraphToMarkdown()`: Paragraph processing
   - `convertTextRunToMarkdown()`: Text run processing
   - `convertTableToMarkdown()`: Table processing

4. **Document Tools** (Lines 275-680)
   - `readGoogleDoc`: Read with format options
   - `listDocumentTabs`: Tab structure listing
   - `appendToGoogleDoc`: Append text
   - `insertText`: Insert at index
   - `deleteRange`: Delete content range

5. **Formatting Tools** (Lines 680-900)
   - `applyTextStyle`: Character formatting
   - `applyParagraphStyle`: Paragraph formatting

6. **Structure Tools** (Lines 900-1050)
   - `insertTable`: Table creation
   - `editTableCell`: (Not implemented)
   - `insertPageBreak`: Page break insertion
   - `insertImageFromUrl`: URL image insertion
   - `insertLocalImage`: Local image upload

7. **Comment Tools** (Lines 1050-1360)
   - Full comment lifecycle management
   - Uses Drive API v3 for comments

8. **Drive Tools** (Lines 1430-2120)
   - Document discovery
   - Folder management
   - File operations

9. **Sheets Tools** (Lines 2260-2575)
   - Spreadsheet CRUD operations
   - Sheet/tab management

10. **Enhanced Formatting** (Lines 2575-3000)
    - Structured content creation
    - Formatted document generation

**Exported**: None (self-contained server)

---

#### `src/auth.ts` (Authentication)
**Purpose**: Google API authentication handling

**Key Functions**:
```typescript
authorize(): Promise<OAuth2Client>
// Main entry point, determines auth method

loadOAuth2Credentials(): OAuth2Credentials
// Loads credentials.json

createOAuth2Client(credentials: OAuth2Credentials): OAuth2Client
// Creates OAuth2 client instance

getAccessToken(client: OAuth2Client): Promise<OAuth2Client>
// Interactive OAuth flow

loadStoredToken(client: OAuth2Client): boolean
// Loads existing token.json

storeToken(token: Credentials): void
// Saves token to disk

authorizeServiceAccount(): Promise<OAuth2Client>
// Service account authentication
```

**Environment Variables**:
- `SERVICE_ACCOUNT_PATH`: Path to service account key
- `GOOGLE_IMPERSONATE_USER`: User email for impersonation

---

#### `src/types.ts` (Type Definitions)
**Purpose**: Zod schemas and TypeScript types

**Key Exports**:
```typescript
// Parameter Schemas
DocumentIdParameter: z.ZodObject
RangeParameters: z.ZodObject
OptionalRangeParameters: z.ZodObject
TextFindParameter: z.ZodObject
TextStyleParameters: z.ZodObject
ParagraphStyleParameters: z.ZodObject

// Composite Schemas
ApplyTextStyleToolParameters: z.ZodObject
ApplyParagraphStyleToolParameters: z.ZodObject

// Type Aliases
TextStyleArgs: z.infer<typeof TextStyleParameters>
ParagraphStyleArgs: z.infer<typeof ParagraphStyleParameters>

// Utility Functions
validateHexColor(color: string): boolean
hexToRgbColor(hex: string): RgbColor | null

// Error Classes
NotImplementedError extends UserError
```

---

#### `src/googleDocsApiHelpers.ts` (Docs Helpers)
**Purpose**: Google Docs API abstraction layer

**Key Exports**:
```typescript
// Text Operations
findTextRange(docs, documentId, text, instance): Promise<Range | null>
insertText(docs, documentId, text, index): Promise<void>
getParagraphRange(docs, documentId, index): Promise<Range | null>

// Batch Operations
executeBatchUpdate(docs, documentId, requests): Promise<BatchUpdateResponse>

// Style Builders
buildUpdateTextStyleRequest(start, end, style): RequestInfo
buildUpdateParagraphStyleRequest(start, end, style): RequestInfo

// Tab Helpers
findTabById(document, tabId): Tab | null
getAllTabs(document): TabWithLevel[]
getTabTextLength(documentTab): number

// Structure Operations
createTable(docs, documentId, rows, cols, index): Promise<void>
insertInlineImage(docs, documentId, url, index, width?, height?): Promise<void>
uploadImageToDrive(drive, localPath, parentFolderId?): Promise<string>
detectAndFormatLists(docs, documentId, start?, end?): Promise<void>
```

---

#### `src/googleSheetsApiHelpers.ts` (Sheets Helpers)
**Purpose**: Google Sheets API abstraction layer

**Key Exports**:
```typescript
// Coordinate Conversion
a1ToRowCol(a1: string): { row: number; col: number }
rowColToA1(row: number, col: number): string
normalizeRange(range: string, sheetName?: string): string

// Range Operations
readRange(sheets, spreadsheetId, range): Promise<ValueRange>
writeRange(sheets, spreadsheetId, range, values, option?): Promise<UpdateResponse>
appendValues(sheets, spreadsheetId, range, values, option?): Promise<AppendResponse>
clearRange(sheets, spreadsheetId, range): Promise<ClearResponse>

// Metadata
getSpreadsheetMetadata(sheets, spreadsheetId): Promise<Spreadsheet>
addSheet(sheets, spreadsheetId, sheetTitle): Promise<BatchUpdateResponse>

// Formatting
formatCells(sheets, spreadsheetId, range, format): Promise<BatchUpdateResponse>
hexToRgb(hex: string): RgbColor | null
```

---

### Test Files

#### `tests/helpers.test.js`
**Test Framework**: Node.js built-in test runner
**Coverage**: `findTextRange` function

**Test Cases**:
1. Find text within single text run
2. Find nth instance of text
3. Return null if text not found
4. Handle text spanning multiple text runs

**Test Method**: Mock-based testing with `mock.fn()`

---

#### `tests/types.test.js`
**Coverage**: Color validation and conversion

**Test Cases**:
1. `validateHexColor`:
   - Valid hex with/without hash
   - 3-digit and 6-digit formats
   - Invalid formats rejection

2. `hexToRgbColor`:
   - 6-digit conversion
   - 3-digit shorthand conversion
   - Without hash prefix
   - Invalid input returns null

---

### Configuration Files

#### `package.json`
```json
{
  "name": "google-docs-mcp",
  "version": "1.0.0",
  "type": "module",
  "main": "index.js",
  "scripts": {
    "build": "tsc",
    "start": "node dist/server.js",
    "test": "node --test"
  }
}
```

#### `tsconfig.json`
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "./dist",
    "rootDir": "./src",
    "strict": true,
    "esModuleInterop": true
  }
}
```

---

## Code Metrics

| File | Lines | Functions | Complexity |
|------|-------|-----------|------------|
| server.ts | ~3000 | 45 | High |
| auth.ts | ~200 | 7 | Medium |
| types.ts | ~250 | 4 | Low |
| googleDocsApiHelpers.ts | ~500 | 15 | Medium |
| googleSheetsApiHelpers.ts | ~430 | 12 | Medium |
| **Total** | **~4380** | **83** | - |

---

## Import/Export Graph

```
index.js
└── dist/server.js
    ├── fastmcp
    ├── zod
    ├── googleapis
    ├── ./auth.js
    │   └── google-auth-library
    ├── ./types.js
    │   └── zod
    ├── ./googleDocsApiHelpers.js
    │   ├── googleapis
    │   └── fastmcp (UserError)
    └── ./googleSheetsApiHelpers.js
        ├── googleapis
        └── fastmcp (UserError)
```

---

## Critical Paths

### Authentication Flow
```
server.ts:initializeGoogleClient()
  → auth.ts:authorize()
    → auth.ts:authorizeServiceAccount() OR auth.ts:loadStoredToken()
    → google-auth-library:OAuth2Client
```

### Tool Execution Flow
```
FastMCP Tool Call
  → server.ts:tool.execute()
    → server.ts:getDocsClient()
    → googleDocsApiHelpers.ts:function()
    → googleapis:docs.documents.batchUpdate()
    → Response processing
```

### Error Handling Flow
```
googleapis:API Error
  → Catch block in tool.execute()
    → Check error.code (404, 403, etc.)
    → Throw fastmcp:UserError with friendly message
```
