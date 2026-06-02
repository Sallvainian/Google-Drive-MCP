# Source Tree Analysis

**Generated:** 2026-05-05 | **Scan Level:** Exhaustive

---

## Directory Structure

```
Google-Drive-MCP/
├── index.js                    # Entry point — imports dist/server.js
├── package.json                # npm manifest (fastmcp, googleapis, zod)
├── package-lock.json           # npm lockfile
├── tsconfig.json               # TypeScript config (ES2022, NodeNext, strict)
├── credentials.json            # Google OAuth client secrets (gitignored)
├── token.json                  # Persisted OAuth refresh token (gitignored)
├── export-docs.mjs             # Standalone script: export Google Docs as .docx
├── .envrc                      # direnv: Google OAuth env vars
├── .gitignore                  # Ignores node_modules, dist, token.json, .env
├── .mcp.json                   # MCP client configuration
├── README.md                   # Setup, scopes, usage examples
├── claude.md                   # AI assistant quick reference
├── SAMPLE_TASKS.md             # 15 example workflows
├── vscode.md                   # VS Code MCP extension setup
│
├── src/                        # TypeScript source (9 files, 10,279 LOC)
│   ├── server.ts               # [ENTRY] FastMCP server — all 108 tool definitions (6,015 LOC)
│   ├── auth.ts                 # Authentication: OAuth2 + Service Account + account enforcement (372 LOC)
│   ├── types.ts                # Zod schemas for all tool parameters (386 LOC)
│   ├── googleDocsApiHelpers.ts # Docs API: batch update, text find, styles, tables, images, tabs (1,041 LOC)
│   ├── googleSheetsApiHelpers.ts # Sheets API: A1 notation, range CRUD, formatting (427 LOC)
│   ├── googleSlidesApiHelpers.ts # Slides API: EMU conversion, element builders, batch update (619 LOC)
│   ├── googleGmailApiHelpers.ts  # Gmail API: email creation, MIME, search, threads, attachments (801 LOC)
│   ├── gmailLabelManager.ts    # Gmail labels: CRUD, resolve IDs, system labels (298 LOC)
│   ├── gmailFilterManager.ts   # Gmail filters: CRUD, template-based creation (320 LOC)
│   └── backup/                 # Backup of older code (excluded from build)
│
├── dist/                       # Compiled JavaScript output (tsc) — gitignored
│   ├── server.js               # Compiled entry point
│   ├── auth.js
│   ├── types.js
│   ├── googleDocsApiHelpers.js
│   ├── googleSheetsApiHelpers.js
│   ├── googleSlidesApiHelpers.js
│   ├── googleGmailApiHelpers.js
│   ├── gmailLabelManager.js
│   └── gmailFilterManager.js
│
├── tests/                      # Unit tests (node:test runner; ~70 cases)
│   ├── types.test.js           # Color validation and hex-to-RGB conversion
│   ├── helpers.test.js         # Text range finding and table cell range helpers
│   └── slides.test.js          # EMU conversion, object ID gen, request builders, batch updates
│
├── docs/                       # Documentation + GitHub Pages
│   ├── index.md                # Master documentation index
│   ├── project-overview.md     # Executive summary, tech stack, tool categories
│   ├── architecture.md         # Architecture documentation
│   ├── source-tree-analysis.md # This file
│   ├── development-guide.md    # Setup, commands, env vars, MCP client config
│   ├── project-scan-report.json # Workflow state file
│   ├── index.html              # GitHub Pages: OAuth app landing page
│   ├── privacy.html            # GitHub Pages: Privacy policy
│   ├── terms.html              # GitHub Pages: Terms of service
│   ├── google25e86b582c5bc5ae.html # Google site verification token
│   └── .archive/               # Archived workflow state files
│
├── pages/                      # Additional page templates
├── assets/                     # Static assets
│
├── .github/workflows/          # CI/CD
│   ├── claude.yml              # Claude Code action (issue/PR comment trigger)
│   └── claude-code-review.yml  # Claude code review workflow
│
├── .claude/                    # Claude Code configuration
├── .agents/                    # Agent configurations
├── .codex/, .cursor/           # Other editor/agent configs
├── _bmad/                      # BMAD workflow system
└── _bmad-output/               # BMAD workflow outputs
```

## Critical Files

### Entry Points

| File | Role |
|------|------|
| `index.js` | npm entry point, imports `dist/server.js` |
| `src/server.ts` | Main source: FastMCP server with all 108 tool definitions |
| `src/auth.ts` | Authentication router (OAuth2 vs Service Account) |

### Core Source Files by Domain

| File | LOC | Domain | Key Exports |
|------|-----|--------|-------------|
| `server.ts` | 6,015 | All | FastMCP server instance, 108 tool registrations, process-level error handlers (`uncaughtException`, `unhandledRejection` with rejection storm detection) |
| `googleDocsApiHelpers.ts` | 1,041 | Docs | `executeBatchUpdate`, `findTextRange`, `getParagraphRange`, `getTableCellRange`, `findSectionRange`, `buildUpdateTextStyleRequest`, `buildUpdateParagraphStyleRequest`, `insertInlineImage`, `uploadImageToDrive`, `getAllTabs`, `findTabById` |
| `googleGmailApiHelpers.ts` | 801 | Gmail | `createSimpleEmail`, `createEmailWithAttachments`, `sendEmail`, `createDraft`, `getMessage`, `searchMessages`, `getThread`, `listThreads`, `modifyMessageLabels`, `batchModifyMessages`, `deleteMessage`, `batchDeleteMessages`, `trashMessage`, `untrashMessage`, `downloadAttachment`, `getUserProfile`, `listDrafts`, `getDraft`, `updateDraft`, `deleteDraft`, `sendDraft`, `formatMessage`, `extractPlainText`, `extractHtmlContent`, `extractAttachments` |
| `googleSlidesApiHelpers.ts` | 619 | Slides | `emuFromPoints`, `pointsFromEmu`, `generateObjectId`, `executeBatchUpdate`, `getPresentation`, `getSlide`, `createTransform`, `createSize`, `hexToRgbColor`, `buildCreateSlideRequest`, `buildCreateShapeRequest`, `buildCreateImageRequest`, `buildCreateTableRequest`, `buildInsertTextRequest`, `buildDeleteTextRequest`, `buildUpdateSlidesPositionRequest`, `buildUpdateShapePropertiesRequest`, `buildUpdateTextStyleRequest`, `getSpeakerNotesShapeId` |
| `googleSheetsApiHelpers.ts` | 427 | Sheets | `a1ToRowCol`, `rowColToA1`, `normalizeRange`, `readRange`, `writeRange`, `appendValues`, `clearRange`, `getSpreadsheetMetadata`, `addSheet`, `formatCells`, `hexToRgb` |
| `auth.ts` | 372 | Auth | `authorize` (main export), `authorizeWithServiceAccount`, `loadSavedCredentialsIfExist`, `authenticate`, `enforceRequiredAccount`, `saveCredentials`, `installTokenRefreshListener`, `isRecoverableAuthError` |
| `types.ts` | 386 | Schemas | Zod schemas: `DocumentIdParameter`, `RangeParameters`, `TextStyleParameters`, `ParagraphStyleParameters`, `PresentationIdParameter`, `ShapeTypeEnum`, `PredefinedLayoutEnum`, `MessageIdParameter`, `GmailSearchParameter`, `SendEmailParameter`, `CreateFilterParameter`, `FilterTemplateParameter`, + ~30 more; `hexColorRegex`, `validateHexColor`, `hexToRgbColor`, `NotImplementedError` |
| `gmailFilterManager.ts` | 320 | Gmail Filters | `listFilters`, `getFilter`, `createFilter`, `deleteFilter`, `createFilterFromTemplate`, `formatFilterForDisplay` |
| `gmailLabelManager.ts` | 298 | Gmail Labels | `listLabels`, `getLabel`, `findLabelByName`, `createLabel`, `updateLabel`, `deleteLabel`, `getOrCreateLabel`, `resolveLabelIds`, `formatLabelsForDisplay` |

### Configuration Files

| File | Purpose |
|------|---------|
| `credentials.json` | Google OAuth client ID/secret (gitignored) |
| `token.json` | Persisted OAuth refresh token (gitignored) |
| `.envrc` | direnv config with Google OAuth env vars |
| `.mcp.json` | MCP client configuration |
| `tsconfig.json` | TypeScript compiler options (target ES2022, module NodeNext, strict) |

## Dependency Graph

```
index.js
  └── dist/server.js (compiled from src/server.ts)
        ├── src/auth.ts
        │     ├── googleapis (google.auth.OAuth2)
        │     ├── google-auth-library (OAuth2Client, JWT)
        │     └── fs/promises, http, url, child_process (stdlib)
        ├── src/types.ts
        │     ├── zod (z)
        │     └── googleapis (docs_v1 types only)
        ├── src/googleDocsApiHelpers.ts
        │     ├── googleapis (docs_v1)
        │     ├── google-auth-library
        │     ├── fastmcp (UserError)
        │     └── ./types.ts
        ├── src/googleSheetsApiHelpers.ts
        │     ├── googleapis (sheets_v4)
        │     └── fastmcp (UserError)
        ├── src/googleSlidesApiHelpers.ts
        │     ├── googleapis (slides_v1)
        │     └── fastmcp (UserError)
        ├── src/googleGmailApiHelpers.ts
        │     ├── googleapis (gmail_v1)
        │     ├── fastmcp (UserError)
        │     └── fs/promises, path (stdlib)
        ├── src/gmailLabelManager.ts
        │     ├── googleapis (gmail_v1)
        │     └── fastmcp (UserError)
        └── src/gmailFilterManager.ts
              ├── googleapis (gmail_v1)
              ├── fastmcp (UserError)
              └── ./types.ts (FilterCriteriaArgs, FilterActionArgs)
```

## Code Metrics

| Metric | Value |
|--------|-------|
| Total source lines | 10,279 |
| Source files | 9 |
| Largest file | `server.ts` (6,015 LOC — 58.5% of codebase) |
| Test files | 3 |
| Test cases | ~70 (`types.test.js`: 11, `helpers.test.js`: 14, `slides.test.js`: 45) |
| External dependencies | 4 (fastmcp, googleapis, google-auth-library, zod) |
| Dev dependencies | 3 (typescript, tsx, @types/node) |
| Tool count | 108 |
| Google API scopes | 5 (documents, drive, spreadsheets, presentations, mail.google.com) |

## File Locations Quick Reference

| Looking for... | Go to... |
|----------------|----------|
| Tool definitions (parameters, descriptions, handlers) | `src/server.ts` |
| Reusable Zod schemas | `src/types.ts` |
| Auth flow, token rotation, account enforcement | `src/auth.ts` |
| Docs batch update, text find, table helpers | `src/googleDocsApiHelpers.ts` |
| Sheets A1 conversion, range CRUD | `src/googleSheetsApiHelpers.ts` |
| Slides EMU conversion, request builders | `src/googleSlidesApiHelpers.ts` |
| Gmail MIME encoding, message parsing | `src/googleGmailApiHelpers.ts` |
| Gmail label operations | `src/gmailLabelManager.ts` |
| Gmail filter templates | `src/gmailFilterManager.ts` |
| OAuth scopes (5 scopes constant) | `src/auth.ts` line 20 |
| Process error handler / rejection storm logic | `src/server.ts` lines 107–154 |
| Lazy client initialization | `src/server.ts` `initializeGoogleClient()` |
