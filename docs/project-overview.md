# Google-Drive-MCP -- Project Overview

**Generated:** 2026-05-05 | **Scan Level:** Exhaustive | **Mode:** Full Rescan

---

## Executive Summary

Google-Drive-MCP is a comprehensive **Model Context Protocol (MCP) server** that bridges AI assistants with 5 Google Workspace APIs. Built on FastMCP 3.35.0, it exposes **108 tools** for programmatic interaction with Google Docs, Sheets, Slides, Drive, and Gmail -- enabling AI-driven document creation, email management, presentation building, and file operations.

## Project Identity

| Attribute | Value |
|-----------|-------|
| **Name** | mcp-googledocs-server |
| **Server Name** | Ultimate Google Docs & Sheets MCP Server |
| **Type** | Library / MCP Server |
| **Repository** | Monolith (single codebase) |
| **Language** | TypeScript (ES2022, NodeNext modules) |
| **Framework** | FastMCP 3.35.0 |
| **Runtime** | Node.js 18+ (ESM, `"type": "module"`) |
| **Package Manager** | npm |
| **License** | ISC |
| **Source Files** | 9 TypeScript files, 10,279 LOC |
| **Test Files** | 3 (node:test runner, 70 test cases) |

## Technology Stack

| Category | Technology | Version | Purpose |
|----------|-----------|---------|---------|
| Framework | FastMCP | ^3.35.0 | MCP server framework with tool registration |
| Schema Validation | Zod | ^3.24.2 | Runtime parameter validation for all 108 tools |
| Google APIs | googleapis | ^148.0.0 | Unified Google Workspace API client |
| Auth | google-auth-library | ^9.15.1 | OAuth2 + Service Account (JWT) auth |
| Language | TypeScript | ^5.8.3 | Strict mode, type-safe development |
| Dev Tooling | tsx | ^4.19.3 | TypeScript execution for development |
| Types | @types/node | ^22.14.1 | Node.js type definitions |

## Tool Categories (108 total)

| Category | Count | Key Tools |
|----------|-------|-----------|
| Google Docs (Core) | 7 | `readGoogleDoc`, `appendToGoogleDoc`, `insertText`, `replaceAllText`, `deleteRange`, `listDocumentTabs`, `formatMatchingText` |
| Formatting | 2 | `applyTextStyle`, `applyParagraphStyle` |
| Document Structure | 9 | `insertTable`, `editTableCell`, `insertDocTableRow`, `deleteDocTableRow`, `insertPageBreak`, `insertImageFromUrl`, `insertLocalImage`, `findElement`*, `fixListFormatting`* |
| Formatted Docs | 4 | `createFormattedDocument`, `insertFormattedContent`, `replaceDocumentContent`, `updateDocumentSection` |
| Comments | 6 | `listComments`, `getComment`, `addComment`, `replyToComment`, `resolveComment`, `deleteComment` |
| Google Sheets | 8 | `readSpreadsheet`, `writeSpreadsheet`, `appendSpreadsheetRows`, `clearSpreadsheetRange`, `createSpreadsheet`, `getSpreadsheetInfo`, `addSpreadsheetSheet`, `formatSpreadsheetCells` |
| Google Drive | 21 | `listGoogleDocs`, `listGoogleSheets`, `listGoogleSlides`, `searchGoogleDocs`, `getRecentGoogleDocs`, `getDocumentInfo`, `getFolderInfo`, `createFolder`, `listFolderContents`, `listAllFolders`, `moveFile`, `copyFile`, `renameFile`, `deleteFile`, `uploadFile`, `downloadFile`, `createDocument`, `createFromTemplate`, `shareFile`, `makeFilePublic`, `listFilePermissions` |
| Google Slides | 17 | `getPresentation`, `listSlides`, `getSlide`, `mapSlide`, `createPresentation`, `addSlide`, `duplicateSlide`, `addTextBox`, `addShape`, `addImage`, `addTable`, `editSlideTableCell`, `deleteSlide`, `deleteElement`, `updateSpeakerNotes`, `moveSlide`, `insertTextInElement` |
| Gmail | 34 | `send_email`, `draft_email`, `read_email`, `search_emails`, `reply_to_email`, `forward_email`, `get_thread`, `list_threads`, `trash_email`, `archive_email`, `mark_as_read`, `mark_as_unread`, `list_email_labels`, `create_label`, `update_label`, `delete_label`, `get_or_create_label`, `create_filter`, `list_filters`, `get_filter`, `delete_filter`, `create_filter_from_template`, `batch_modify_emails`, `batch_delete_emails`, `download_attachment`, `modify_email`, `delete_email`, `list_drafts`, `get_draft`, `update_draft`, `delete_draft`, `send_draft`, `get_user_profile` (+1) |

\* = Stubs that throw `NotImplementedError` (`findElement`) or are experimental (`fixListFormatting`)

## API Integrations

| Google API | Version | OAuth Scope | Key Capabilities |
|------------|---------|-------------|-----------------|
| Docs | v1 | `documents` | Read/write/format, tables, images, comments, tabs, sections |
| Drive | v3 | `drive` (full) | File CRUD, folders, search, upload/download, templates, permissions |
| Sheets | v4 | `spreadsheets` | Range read/write/append/clear, creation, cell formatting |
| Slides | v1 | `presentations` | Presentations, slides, shapes, images, tables, speaker notes |
| Gmail | v1 | `mail.google.com` (full) | Full email lifecycle: send, draft, read, search, labels, filters, threads, attachments |

## Authentication

Three modes, evaluated in order at startup:

1. **Service Account (JWT)** -- when `SERVICE_ACCOUNT_PATH` is set
   - Reads JSON key file, creates JWT client with all 5 scopes
   - Optional domain-wide delegation via `GOOGLE_IMPERSONATE_USER`

2. **OAuth 2.0 with persisted credentials** (default path)
   - Tries `GOOGLE_REFRESH_TOKEN` env var first
   - On stale env-var token: falls back to `token.json` rather than failing
   - On stale file token: falls back to interactive OAuth
   - Refresh tokens rotated by Google are auto-persisted to disk

3. **OAuth 2.0 interactive** -- triggered when no valid token exists
   - Spins up local HTTP server on port 3000
   - Opens browser for consent
   - 5-minute timeout
   - Web client type supported (external redirect flow)

**Account enforcement:** When `REQUIRED_ACCOUNT_EMAIL` is set, `enforceRequiredAccount()` calls `drive.about.get` after authorization and throws if the authenticated email doesn't match. The required email is also seeded as `login_hint` on the consent URL to pre-select the correct account in Google's account picker.

## Key Design Decisions

- **Single server, multiple APIs**: All 108 tools share one auth client in one process
- **Lazy initialization**: Google API clients created on first tool invocation
- **Batch update chunking**: Docs/Slides batch updates auto-split at 50 requests per call
- **UserError vs Error**: Client-facing errors use FastMCP `UserError`; internal errors use standard `Error`
- **Large file fallback**: Drive `files.export` has a hard 10MB limit; on 403/size error, `downloadFile` falls back to `exportViaWebUrl()` using no-limit web export URLs
- **Process-level resilience**:
  - `uncaughtException`: exits cleanly on EPIPE/EBADF (parent died), keeps server alive otherwise
  - `unhandledRejection`: rejection-storm detector (≥10 rejections in 60s → exit for clean supervisor restart); otherwise logs with counter context and continues
- **Stderr-only logging**: All server logs go through `console.error()` since stdout is the MCP protocol transport

## Known Limitations

- Programmatically created comments appear in "All Comments" sidebar but aren't visibly anchored in the Docs UI
- Resolved comment status may not persist (Drive API limitation)
- `findElement` throws `NotImplementedError`; `fixListFormatting` is experimental
- Drive `files.export` returns 403 for files >10MB → falls back to web URL (no direct content)
- Gmail attachment uploads may timeout for large files
- Rate limits apply per Google API quotas
- Adding Gmail scope to an existing OAuth setup requires deleting `token.json` and re-authenticating
