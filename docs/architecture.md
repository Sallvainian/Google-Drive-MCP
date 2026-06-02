# Architecture Overview

**Generated:** 2026-05-05 | **Scan Level:** Exhaustive

---

## System Design

Google-Drive-MCP follows a **monolith helper-per-domain** architecture. A single FastMCP server (`server.ts`) acts as the orchestration layer, registering all 108 tools. Each Google API domain has dedicated helper modules that encapsulate low-level API interactions.

```
                    MCP Client (Claude, VS Code, etc.)
                              |
                         MCP Protocol (stdio)
                              |
                    +---------v---------+
                    |   FastMCP Server  |
                    |   (server.ts)     |
                    |   108 tool defs   |
                    +---+---+---+---+---+
                        |   |   |   |
           +------------+   |   |   +------------+
           |            |   |   |                |
    +------v------+ +---v---v---+  +------v------+  +------v-------+
    | Docs Helpers| |Sheets Help|  |Slides Helper|  | Gmail Helpers|
    | (1,041 LOC) | | (427 LOC) |  | (619 LOC)   |  | (801 LOC)    |
    +------+------+ +-----+-----+  +------+------+  +------+-------+
           |               |               |                |
           +-------+-------+-------+-------+         +------+------+
                   |                                 | Label Mgr   |
            +------v------+                          | (298 LOC)   |
            |  types.ts   |                          +-------------+
            | Zod Schemas |                          | Filter Mgr  |
            | (386 LOC)   |                          | (320 LOC)   |
            +------+------+                          +-------------+
                   |
            +------v------+
            |   auth.ts   |
            | OAuth2/JWT  |
            | (372 LOC)   |
            +-------------+
                   |
            Google APIs (googleapis ^148.0.0)
```

## Module Breakdown

### server.ts (6,015 LOC) -- Orchestration Layer

The largest file in the codebase (~58% of total source LOC). Responsibilities:

- Creates a `FastMCP` instance with name "Ultimate Google Docs & Sheets MCP Server" (version 1.0.0)
- Holds module-level singleton state for all 5 Google API clients (`googleDocs`, `googleDrive`, `googleSheets`, `googleSlides`, `googleGmail`) and the auth client
- `initializeGoogleClient()` — lazy initialization called on first tool invocation; resets all clients to null on failure to allow retry
- Per-domain getter helpers (`getDocsClient()`, `getDriveClient()`, `getSheetsClient()`, `getSlidesClient()`, `getGmailClient()`) that throw `UserError` if the client is not initialized
- Registers all 108 tools with Zod parameter schemas and async `execute` handlers
- Tools are grouped by API domain in source order: Docs (text/format/structure) → Comments → Drive → Sheets → Formatted Docs → Slides → Gmail

**Process-level resilience handlers (lines 107–154):**

- `uncaughtException`: detects `EPIPE`/`EBADF` (broken stdio = parent process died) and exits cleanly with code 0 to avoid orphaning. Other errors are logged but do not crash the server.
- `unhandledRejection`: same EPIPE/EBADF handling, plus a **rejection storm detector**:
  - Constants: `REJECTION_LIMIT = 10`, `REJECTION_WINDOW_MS = 60_000`
  - Tracks rejection count within a 60s sliding window
  - On ≥10 rejections within 60s, exits with code 1 so the supervisor can restart with fresh state
  - Otherwise logs full stack with counter context (`[N/10 in 60s]`) and continues

### auth.ts (372 LOC) -- Authentication Layer

Authentication router supporting two methods, with an OAuth fallback chain:

**Service Account flow** (`authorizeWithServiceAccount`):
1. Read key file from `SERVICE_ACCOUNT_PATH`
2. Create JWT client with all 5 scopes
3. Optional impersonation via `GOOGLE_IMPERSONATE_USER` (domain-wide delegation)

**OAuth 2.0 fallback chain** (`loadSavedCredentialsIfExist`):
1. If `GOOGLE_REFRESH_TOKEN` env var set:
   - Try refresh; on success, persist if Google rotated the token; verify required-account match
   - On `invalid_grant`, `invalid_client`, `token has been expired`, or `token has been revoked` → fall through to file (don't trigger interactive OAuth for a stale env var)
   - On non-recoverable error → throw
2. If `token.json` exists:
   - Same refresh + verify pattern
   - Same recoverable-error fallthrough (return null → triggers interactive OAuth)
3. Interactive OAuth (`authenticate`):
   - Spin up local HTTP server on port 3000
   - Open browser to consent URL with `prompt=consent`, `access_type=offline`
   - Seeds `login_hint` from `REQUIRED_ACCOUNT_EMAIL` if set
   - 5-minute timeout
   - Calls `enforceRequiredAccount()` BEFORE persisting tokens (wrong-account auth never reaches disk)

**Token rotation:** `installTokenRefreshListener()` subscribes to `tokens` events on the OAuth2Client; whenever Google rotates the refresh token, the new one is written to disk via `saveCredentials()`.

**Account enforcement:** `enforceRequiredAccount()` calls `drive.about.get({fields: 'user'})` and throws on email mismatch with `REQUIRED_ACCOUNT_EMAIL`. Used to pin per-instance accounts in multi-account setups (e.g., `Drive-MCP-Personal` vs `Drive-MCP-Work`).

**Path overrides:** `TOKEN_PATH` and `CREDENTIALS_PATH` env vars allow custom paths for multi-account setups.

**Auth precedence summary:**
```
authorize()
  ├── SERVICE_ACCOUNT_PATH? → authorizeWithServiceAccount() → JWT
  └── No service account:
       ├── GOOGLE_REFRESH_TOKEN? → try refresh → ok? done | stale? fall through
       ├── token.json?            → try refresh → ok? done | stale? null
       └── No saved token         → authenticate() → interactive OAuth → save token
  → enforceRequiredAccount() (if REQUIRED_ACCOUNT_EMAIL set)
```

### types.ts (386 LOC) -- Schema Definitions

Centralized Zod schemas organized by domain:

- **Reusable fragments**: `DocumentIdParameter`, `RangeParameters`, `OptionalRangeParameters`, `TextFindParameter`
- **Style schemas**: `TextStyleParameters`, `ParagraphStyleParameters` with type aliases (`TextStyleArgs`, `ParagraphStyleArgs`)
- **Tool-level combinations**: `ApplyTextStyleToolParameters`, `ApplyParagraphStyleToolParameters` (use `z.union` for find-text-vs-range targeting)
- **Slides schemas**: `PresentationIdParameter`, `PageObjectIdParameter`, `SlidePositionParameter`, `ElementSizeParameter`, `ElementPositionParameter`, `ShapeTypeEnum` (24 shape types), `PredefinedLayoutEnum` (11 layouts)
- **Gmail schemas**: `MessageIdParameter`, `ThreadIdParameter`, `DraftIdParameter`, `LabelIdParameter`, `FilterIdParameter`, `EmailRecipientsParameter`, `EmailContentParameter`, `EmailAttachmentsParameter`, `EmailReplyParameter`, `GmailSearchParameter`, `LabelVisibilityParameter`, `CreateLabelParameter`, `UpdateLabelParameter`, `ModifyLabelsParameter`, `BatchModifyLabelsParameter`, `FilterCriteriaParameter`, `FilterActionParameter`, `CreateFilterParameter`, `FilterTemplateType` (6 templates), `FilterTemplateParameter`, `DownloadAttachmentParameter`, `BatchDeleteParameter`
- **Combined Gmail schemas**: `SendEmailParameter`, `DraftEmailParameter` (built via `.merge()`)

Also exports:
- `hexColorRegex`, `validateHexColor()`, `hexToRgbColor()` — color validation/conversion utilities
- `NotImplementedError` class — for stub implementations

### googleDocsApiHelpers.ts (1,041 LOC) -- Docs API

Key capabilities:

- **Batch updates**: `executeBatchUpdate()` with auto-chunking at 50 requests per call
- **Text operations**: `findTextRange()` with multi-run support (text spanning multiple `textRun` elements), `getParagraphRange()` with table recursion
- **Style builders**: `buildUpdateTextStyleRequest()`, `buildUpdateParagraphStyleRequest()` — construct Google Docs API request objects from simplified params
- **Table operations**: `getTableCellRange()`, `findTableStartIndexByText()`, `findCellBelowHeader()` — navigate and manipulate table cells
- **Image operations**: `insertInlineImage()`, `uploadImageToDrive()` — insert images from URL or local files
- **Tab management**: `getAllTabs()`, `findTabById()`, `getTabTextLength()` — recursive tab navigation
- **Section management**: `findSectionRange()` — find heading-bounded sections for content replacement
- **Stubs**: `findParagraphsMatchingStyle()`, `detectAndFormatLists()`, `addCommentHelper()` — throw `NotImplementedError`

### googleSheetsApiHelpers.ts (427 LOC) -- Sheets API

Key capabilities:

- **A1 notation**: `a1ToRowCol()`, `rowColToA1()`, `normalizeRange()` — bidirectional conversion
- **Range CRUD**: `readRange()`, `writeRange()`, `appendValues()`, `clearRange()`
- **Metadata**: `getSpreadsheetMetadata()` — get sheet info without grid data
- **Sheet management**: `addSheet()` — create new tabs
- **Formatting**: `formatCells()` — apply background color, text format, alignment via `repeatCell` request
- **Color utility**: `hexToRgb()` (local copy)

### googleSlidesApiHelpers.ts (619 LOC) -- Slides API

Key capabilities:

- **Unit conversion**: `emuFromPoints()`, `pointsFromEmu()` — EMU (English Metric Units) conversion (12,700 EMU per point)
- **Object IDs**: `generateObjectId()` — timestamp + random format
- **Batch updates**: `executeBatchUpdate()` with 50-request chunking
- **Retrieval**: `getPresentation()`, `getSlide()`
- **Element creation helpers**: `createTransform()`, `createSize()`, `createPageElementProperties()`
- **Request builders**: `buildCreateSlideRequest()`, `buildCreateShapeRequest()`, `buildCreateImageRequest()`, `buildCreateTableRequest()`, `buildInsertTextRequest()`, `buildDeleteTextRequest()`, `buildUpdateSlidesPositionRequest()`, `buildUpdateShapePropertiesRequest()`, `buildUpdateTextStyleRequest()`
- **Speaker notes**: `getSpeakerNotesShapeId()` — find notes shape in slide
- **Local color utility**: `hexToRgbColor()`

### googleGmailApiHelpers.ts (801 LOC) -- Gmail API

Key capabilities:

- **Email creation**: `createSimpleEmail()` (no attachments), `createEmailWithAttachments()` (multipart MIME with base64 attachment encoding)
- **MIME handling**: `base64UrlEncode/Decode()`, `encodeEmailHeader()` (RFC 2047), `getMimeType()` (40+ extension mappings)
- **Message parsing**: `parseEmailHeaders()`, `extractPlainText()`, `extractHtmlContent()`, `extractAttachments()`, `formatMessage()` — recursive multipart extraction
- **Core operations**: `sendEmail()`, `createDraft()`, `getMessage()`, `searchMessages()`, `getThread()`, `listThreads()`
- **Label operations**: `modifyMessageLabels()`, `batchModifyMessages()` (batch-size-aware)
- **Delete operations**: `deleteMessage()`, `batchDeleteMessages()`, `trashMessage()`, `untrashMessage()`
- **Attachments**: `downloadAttachment()` — download and save to filesystem
- **Draft management**: `listDrafts()`, `getDraft()`, `updateDraft()`, `deleteDraft()`, `sendDraft()`
- **Profile**: `getUserProfile()`

### gmailLabelManager.ts (298 LOC) -- Gmail Label Management

CRUD operations for Gmail labels:
- System label awareness (`INBOX`, `SPAM`, `TRASH`, etc.)
- Idempotent `getOrCreateLabel()`
- `resolveLabelIds()` — resolve names or IDs to canonical IDs
- Protection against modifying/deleting system labels

### gmailFilterManager.ts (320 LOC) -- Gmail Filter Management

CRUD operations for Gmail filters:
- Template-based creation: `fromSender`, `withSubject`, `withAttachments`, `largeEmails`, `containingText`, `mailingList`
- Validation that criteria and actions are non-empty
- Display formatting utilities

## Data Flow

### Tool Invocation Flow

```
MCP Client → FastMCP (stdio) → Tool Handler (server.ts)
  → initializeGoogleClient() [lazy, once per process]
  → Zod validation (automatic via FastMCP)
  → getXxxClient() helper
  → Domain helper module (e.g., GDocsHelpers)
  → Google API call (googleapis)
  → Format response → Return string to client
```

Tool return values are always strings — complex data is JSON-serialized or formatted as readable text.

### Authentication Flow at Startup

See **auth.ts** section above for the precedence chain. Key invariants:
- Stale env-var refresh tokens fall through to file (don't trigger consent UI)
- Wrong-account OAuth never reaches `token.json` (validation runs before persistence)
- Token rotation is handled transparently via the `tokens` listener

### Large File Export Path

`downloadFile` for Google-native files (Docs/Sheets/Slides) calls `drive.files.export`. Drive's export API has a hard 10MB limit and returns 403 with a size-related message for larger files. The implementation catches this and falls back to `exportViaWebUrl()`, which uses Drive's no-limit web export URLs (returns the URL rather than the content).

## Testing Strategy

- **Runner**: Node.js built-in (`node --test tests/`). NOT Jest, Vitest, or Mocha.
- **Style**: Test files are JavaScript (`.test.js`), not TypeScript. They import from compiled `dist/` directory using `.js` extensions.
- **Pre-condition**: Always `npm run build` before `npm test` — tests run against compiled output.
- **Test files (3 files, 70 cases total):**
  - `tests/types.test.js` (11 cases): Color validation (`validateHexColor`), hex-to-RGB conversion (`hexToRgbColor`)
  - `tests/helpers.test.js` (14 cases): Text range finding across single/multi text runs, table cell range finding with edge cases (empty cells, out-of-bounds)
  - `tests/slides.test.js` (45 cases): EMU conversion, object ID generation, request builders, batch update error handling (404, 403)
- **Mocking**: `node:test`'s `mock.fn()` for Google API client mocking — no external mock libraries.

## CI/CD

GitHub Actions workflows under `.github/workflows/`:
- `claude.yml`: Triggers on `@claude` mentions in issues/PRs/comments, runs Claude Code action
- `claude-code-review.yml`: Automated code review workflow

No build/test CI — tests are run locally before commits.

## Configuration & Conventions

- **Module system**: ESM only. `"type": "module"` in package.json. No `require()`, no `module.exports`.
- **Imports**: Always use `.js` extension for relative imports — even for `.ts` files (NodeNext module resolution requirement).
- **Logging**: All server output goes through `console.error()` since stdout is the MCP protocol transport. Using `console.log()` would corrupt the protocol stream.
- **TypeScript strict mode**: Enabled (`strict: true`). Target ES2022, module NodeNext.
- **No linter/formatter**: No ESLint or Prettier configured. Match surrounding code style when editing.
- **Path resolution**: Use `fileURLToPath(import.meta.url)` and `path.dirname()` instead of `__dirname`/`__filename` (see `auth.ts` for the canonical pattern).
- **Indices**: Google Docs content indices are 1-based. Slides element positions are in points (72 points = 1 inch; 1 point = 12,700 EMU).
- **Batch update ordering**: For Docs, requests in a batch apply in reverse index order — when making multiple changes, process from end of document to beginning, or indices will shift.
