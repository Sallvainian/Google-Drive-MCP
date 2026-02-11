# Source Tree Analysis

## Project: Google Docs MCP Server
**Generated**: 2026-01-22

---

## Directory Structure

```
mcp-googledocs-server/
├── src/                              # TypeScript source files (9 files, 9,180 LOC)
│   ├── server.ts                     # Main MCP server (92 tools) - 5,301 lines
│   ├── auth.ts                       # Authentication module - 226 lines
│   ├── types.ts                      # Type definitions & Zod schemas - 386 lines
│   ├── googleDocsApiHelpers.ts       # Google Docs API helpers - 827 lines
│   ├── googleSheetsApiHelpers.ts     # Google Sheets API helpers - 427 lines
│   ├── googleSlidesApiHelpers.ts     # Google Slides API helpers - 595 lines
│   ├── googleGmailApiHelpers.ts      # Gmail API helpers - 802 lines
│   ├── gmailLabelManager.ts          # Gmail label management - 299 lines
│   └── gmailFilterManager.ts         # Gmail filter management - 321 lines
│
├── tests/                            # Test files
│   ├── helpers.test.js               # Tests for helper functions
│   └── types.test.js                 # Tests for type validation
│
├── dist/                             # Compiled JavaScript (generated)
│   ├── server.js
│   ├── auth.js
│   ├── types.js
│   ├── googleDocsApiHelpers.js
│   ├── googleSheetsApiHelpers.js
│   ├── googleSlidesApiHelpers.js
│   ├── googleGmailApiHelpers.js
│   ├── gmailLabelManager.js
│   └── gmailFilterManager.js
│
├── docs/                             # Generated documentation
│   ├── documentation.md              # Master index
│   ├── architecture.md               # Architecture overview
│   ├── source-tree.md                # This file
│   ├── developer-guide.md            # Developer guide
│   └── project-scan-report.json      # Scan state file
│
├── assets/                           # Static assets
│   └── google.docs.mcp.1.gif         # Demo animation
│
├── node_modules/                     # Dependencies (not tracked)
│
├── index.js                          # Node.js entry point
├── package.json                      # Project manifest
├── package-lock.json                 # Dependency lock file
├── tsconfig.json                     # TypeScript configuration
├── README.md                         # Main documentation
├── SAMPLE_TASKS.md                   # Example workflows (15 tasks)
├── CLAUDE.md                         # AI assistant instructions
├── vscode.md                         # VS Code integration guide
├── .envrc                            # Environment variables (direnv)
├── credentials.json                  # OAuth credentials (user-provided, gitignored)
└── token.json                        # OAuth token (generated, gitignored)
```

---

## File Analysis

### Source Files

#### `src/server.ts` (Main Server)
**Lines**: 5,301
**Purpose**: Core MCP server implementation with all 92 tools

**Key Sections**:
1. **Imports & Setup** (Lines 1-50)
   - FastMCP, Zod, googleapis imports
   - Global client variables for all 5 Google APIs

2. **API Client Initialization** (Lines 50-150)
   - `initializeGoogleClient()`: Lazy initialization
   - `getDocsClient()`, `getDriveClient()`, `getSheetsClient()`
   - `getSlidesClient()`, `getGmailClient()`

3. **Markdown Conversion** (Lines 150-400)
   - `convertDocsJsonToMarkdown()`: JSON to Markdown converter
   - `convertParagraphToMarkdown()`: Paragraph processing
   - `convertTableToMarkdown()`: Table processing

4. **Document Tools** (Lines 400-900)
   - readGoogleDoc, listDocumentTabs
   - appendToGoogleDoc, insertText, deleteRange

5. **Formatting Tools** (Lines 900-1200)
   - applyTextStyle, applyParagraphStyle, formatMatchingText

6. **Structure Tools** (Lines 1200-1500)
   - insertTable, insertPageBreak
   - insertImageFromUrl, insertLocalImage

7. **Comment Tools** (Lines 1500-1900)
   - Full comment lifecycle management via Drive API v3

8. **Drive Tools** (Lines 1900-2500)
   - Document/folder discovery and management
   - File operations (move, copy, rename, delete)

9. **Sheets Tools** (Lines 2500-2900)
   - Spreadsheet CRUD operations
   - Sheet/tab management

10. **Enhanced Formatting** (Lines 2900-3200)
    - createFormattedDocument, insertFormattedContent
    - replaceDocumentContent, createFromTemplate

11. **Slides Tools** (Lines 3200-4200)
    - Presentation management (16 tools)
    - Element creation (text boxes, shapes, images, tables)
    - Slide manipulation (add, duplicate, delete, move)

12. **Gmail Tools** (Lines 4200-5301)
    - Email operations (34 tools)
    - Label and filter management
    - Thread and draft handling

---

#### `src/auth.ts` (Authentication)
**Lines**: 226
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
// Interactive OAuth flow with HTTP callback server

loadStoredToken(client: OAuth2Client): boolean
// Loads existing token.json

storeToken(token: Credentials): void
// Saves token to disk

authorizeServiceAccount(): Promise<OAuth2Client>
// Service account authentication with impersonation
```

---

#### `src/types.ts` (Type Definitions)
**Lines**: 386
**Purpose**: Zod schemas and TypeScript types

**Key Exports**:
```typescript
// Document Schemas
DocumentIdParameter, RangeParameters, TextFindParameter
TextStyleParameters, ParagraphStyleParameters

// Slides Schemas
SlideElementSchema, SlidePositionSchema
PresentationIdParameter, PageObjectIdParameter

// Gmail Schemas
EmailAddressSchema, EmailComposeSchema
FilterCriteriaSchema, FilterActionSchema
LabelIdSchema, MessageIdSchema

// Utility Functions
validateHexColor(color: string): boolean
hexToRgbColor(hex: string): RgbColor | null
```

---

#### `src/googleDocsApiHelpers.ts` (Docs Helpers)
**Lines**: 827
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
```

---

#### `src/googleSheetsApiHelpers.ts` (Sheets Helpers)
**Lines**: 427
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
formatCells(sheets, spreadsheetId, range, format): Promise<BatchUpdateResponse>
```

---

#### `src/googleSlidesApiHelpers.ts` (Slides Helpers)
**Lines**: 595
**Purpose**: Google Slides API abstraction layer

**Key Exports**:
```typescript
// Unit Conversion
emuFromPoints(pt: number): number  // Points to EMU
pointsFromEmu(emu: number): number // EMU to points

// Object Management
generateObjectId(prefix?: string): string
executeBatchUpdate(slides, presentationId, requests): Promise<Response>

// Retrieval
getPresentation(slides, presentationId): Promise<Presentation>
getSlide(slides, presentationId, pageObjectId): Promise<Page>

// Element Creation
createTransform(x, y): AffineTransform
createSize(width, height): Size
createPageElementProperties(x, y, width, height): PageElementProperties

// Request Builders
buildCreateSlideRequest(insertionIndex?, layout?, objectId?)
buildCreateShapeRequest(pageId, type, x, y, width, height, objectId?)
buildCreateImageRequest(pageId, url, x, y, width, height, objectId?)
buildCreateTableRequest(pageId, rows, cols, x, y, width, height, objectId?)
buildInsertTextRequest(objectId, text, insertionIndex?)
buildUpdateTextStyleRequest(objectId, options, rangeType?, start?, end?)
buildUpdateShapePropertiesRequest(objectId, fillColor?)

// Speaker Notes
getSpeakerNotesShapeId(slide): string | null
```

---

#### `src/googleGmailApiHelpers.ts` (Gmail Helpers)
**Lines**: 802
**Purpose**: Gmail API abstraction layer

**Key Exports**:
```typescript
// Email Validation & Encoding
validateEmail(email: string): boolean
encodeEmailHeader(text: string): string  // RFC 2047 MIME encoding
base64UrlEncode(str: string): string
base64UrlDecode(str: string): string
getMimeType(filename: string): string

// Email Creation
createSimpleEmail(options): string
createEmailWithAttachments(options): Promise<string>

// Message Parsing
parseEmailHeaders(headers): Record<string, string>
extractPlainText(payload): string
extractHtmlContent(payload): string
extractAttachments(payload): AttachmentInfo[]
formatMessage(message): FormattedMessage

// API Operations
sendEmail(gmail, rawEmail, threadId?): Promise<Message>
createDraft(gmail, rawEmail, threadId?): Promise<Draft>
getMessage(gmail, messageId, format?): Promise<Message>
searchMessages(gmail, options): Promise<SearchResult>
getThread(gmail, threadId, format?): Promise<Thread>
listThreads(gmail, options): Promise<ThreadList>
modifyMessageLabels(gmail, messageId, add?, remove?): Promise<Message>
batchModifyMessages(gmail, ids, add?, remove?, batchSize?): Promise<BatchResult>
deleteMessage(gmail, messageId): Promise<void>
batchDeleteMessages(gmail, ids, batchSize?): Promise<BatchResult>
trashMessage(gmail, messageId): Promise<Message>
untrashMessage(gmail, messageId): Promise<Message>
downloadAttachment(gmail, messageId, attachmentId, savePath?, filename?): Promise<SaveResult>
getUserProfile(gmail): Promise<Profile>

// Draft Management
listDrafts(gmail, options): Promise<DraftList>
getDraft(gmail, draftId): Promise<Draft>
updateDraft(gmail, draftId, rawEmail): Promise<Draft>
deleteDraft(gmail, draftId): Promise<void>
sendDraft(gmail, draftId): Promise<Message>
```

---

#### `src/gmailLabelManager.ts` (Label Management)
**Lines**: 299
**Purpose**: Gmail label operations

**Key Exports**:
```typescript
// Constants
SYSTEM_LABELS: readonly string[]  // INBOX, SPAM, TRASH, etc.

// Types
interface LabelInfo { id, name, type, visibility, counts }

// Operations
listLabels(gmail): Promise<LabelInfo[]>
getLabel(gmail, labelId): Promise<LabelInfo>
findLabelByName(gmail, name): Promise<LabelInfo | null>
createLabel(gmail, options): Promise<LabelInfo>
updateLabel(gmail, labelId, options): Promise<LabelInfo>
deleteLabel(gmail, labelId): Promise<void>
getOrCreateLabel(gmail, options): Promise<LabelInfo>
resolveLabelIds(gmail, labelsOrIds): Promise<string[]>
formatLabelsForDisplay(labels): string
```

---

#### `src/gmailFilterManager.ts` (Filter Management)
**Lines**: 321
**Purpose**: Gmail filter operations

**Key Exports**:
```typescript
// Types
type FilterTemplateType = 'fromSender' | 'withSubject' | 'withAttachments' |
                          'largeEmails' | 'containingText' | 'mailingList'

interface FilterInfo { id, criteria, action }

// Operations
listFilters(gmail): Promise<FilterInfo[]>
getFilter(gmail, filterId): Promise<FilterInfo>
createFilter(gmail, criteria, action): Promise<FilterInfo>
deleteFilter(gmail, filterId): Promise<void>
createFilterFromTemplate(gmail, template, params): Promise<FilterInfo>
formatFilterForDisplay(filter): string
formatFiltersForDisplay(filters): string
```

---

## Code Metrics

| File | Lines | Functions | Complexity |
|------|-------|-----------|------------|
| server.ts | 5,301 | 92+ tools | High |
| auth.ts | 226 | 7 | Medium |
| types.ts | 386 | 4 | Low |
| googleDocsApiHelpers.ts | 827 | 15 | Medium |
| googleSheetsApiHelpers.ts | 427 | 12 | Medium |
| googleSlidesApiHelpers.ts | 595 | 18 | Medium |
| googleGmailApiHelpers.ts | 802 | 25 | Medium |
| gmailLabelManager.ts | 299 | 10 | Low |
| gmailFilterManager.ts | 321 | 8 | Low |
| **Total** | **9,180** | **~191** | - |

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
    ├── ./googleSheetsApiHelpers.js
    │   ├── googleapis
    │   └── fastmcp (UserError)
    ├── ./googleSlidesApiHelpers.js
    │   ├── googleapis
    │   └── fastmcp (UserError)
    ├── ./googleGmailApiHelpers.js
    │   ├── googleapis
    │   ├── fastmcp (UserError)
    │   └── fs/promises, path
    ├── ./gmailLabelManager.js
    │   └── googleapis
    └── ./gmailFilterManager.js
        ├── googleapis
        └── ./types.js
```

---

## Configuration Files

### `package.json`
```json
{
  "name": "mcp-googledocs-server",
  "version": "1.0.0",
  "type": "module",
  "main": "index.js",
  "dependencies": {
    "fastmcp": "^3.24.0",
    "google-auth-library": "^9.15.1",
    "googleapis": "^148.0.0",
    "zod": "^3.24.2"
  }
}
```

### `tsconfig.json`
```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "outDir": "./dist",
    "rootDir": "./src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true
  }
}
```

---

## Critical Paths

### Authentication Flow
```
server.ts:initializeGoogleClient()
  → auth.ts:authorize()
    → auth.ts:authorizeServiceAccount() OR auth.ts:getAccessToken()
    → google-auth-library:OAuth2Client
```

### Tool Execution Flow
```
FastMCP Tool Call
  → server.ts:tool.execute()
    → server.ts:get[API]Client()
    → [api]Helpers.ts:function()
    → googleapis:[api].batchUpdate() or get()
    → Response processing
```

### Error Handling Flow
```
googleapis:API Error
  → Catch block in tool.execute()
    → Check error.code (400, 403, 404, etc.)
    → Throw fastmcp:UserError with friendly message
```
