# Google-Drive-MCP — Project Overview

## Executive Summary

Google-Drive-MCP is a comprehensive **Model Context Protocol (MCP) server** that bridges AI assistants (Claude, etc.) with 5 Google Workspace APIs. Built on FastMCP 3.24.0, it exposes **99 tools** for programmatic interaction with Google Docs, Sheets, Slides, Drive, and Gmail — enabling AI-driven document creation, email management, presentation building, and file operations.

## Project Identity

| Attribute | Value |
|-----------|-------|
| **Name** | mcp-googledocs-server |
| **Type** | Library / MCP Server |
| **Repository** | Monolith |
| **Language** | TypeScript (ES2022, NodeNext modules) |
| **Framework** | FastMCP 3.24.0 |
| **Runtime** | Node.js 18+ |
| **Package Manager** | npm |
| **License** | ISC |

## Technology Stack

| Category | Technology | Version | Purpose |
|----------|-----------|---------|---------|
| Framework | FastMCP | ^3.24.0 | MCP server framework |
| Schema Validation | Zod | ^3.24.2 | Runtime parameter validation |
| Google APIs | googleapis | ^148.0.0 | Google Workspace API client |
| Auth | google-auth-library | ^9.15.1 | OAuth2 & Service Account auth |
| Language | TypeScript | ^5.8.3 | Type-safe development |
| Dev Tooling | tsx | ^4.19.3 | TypeScript execution |
| Types | @types/node | ^22.14.1 | Node.js type definitions |

## Tool Categories (99 total)

| Category | Count | Key Tools |
|----------|-------|-----------|
| Google Docs (Core) | 6 | `readGoogleDoc`, `appendToGoogleDoc`, `insertText`, `insertHyperlink`, `deleteRange`, `listDocumentTabs` |
| Formatting | 3 | `applyTextStyle`, `applyParagraphStyle`, `formatMatchingText` |
| Document Structure | 7 | `insertTable`, `editTableCell`, `insertPageBreak`, `insertImageFromUrl`, `insertLocalImage`, `findElement`*, `fixListFormatting`* |
| Comments | 6 | `listComments`, `getComment`, `addComment`, `replyToComment`, `resolveComment`, `deleteComment` |
| Google Sheets | 8 | `readSpreadsheet`, `writeSpreadsheet`, `appendSpreadsheetRows`, `clearSpreadsheetRange`, `createSpreadsheet`, `getSpreadsheetInfo`, `addSpreadsheetSheet`, `listGoogleSheets` |
| Google Drive | 15 | `listGoogleDocs`, `searchGoogleDocs`, `getRecentGoogleDocs`, `getDocumentInfo`, `createFolder`, `listFolderContents`, `moveFile`, `copyFile`, `renameFile`, `deleteFile`, `uploadFile`, `downloadFile`, `createDocument`, `createFromTemplate`, `listAllFolders`, `getFolderInfo` |
| Google Slides | 17 | `getPresentation`, `listSlides`, `getSlide`, `mapSlide`, `createPresentation`, `addSlide`, `duplicateSlide`, `addTextBox`, `addShape`, `addImage`, `addTable`, `editSlideTableCell`, `deleteSlide`, `deleteElement`, `updateSpeakerNotes`, `moveSlide`, `insertTextInElement` |
| Gmail | 34 | `send_email`, `draft_email`, `read_email`, `search_emails`, `reply_to_email`, `forward_email`, `get_thread`, `list_threads`, `trash_email`, `archive_email`, `mark_as_read`, `list_email_labels`, `create_label`, `create_filter`, `batch_modify_emails`, `download_attachment`, + 18 more |
| Formatted Docs | 3 | `createFormattedDocument`, `insertFormattedContent`, `replaceDocumentContent`, `updateDocumentSection` |

\* = Not fully implemented

## API Integrations

1. **Google Docs API v1** — Document reading, writing, formatting, table/image operations, comment management
2. **Google Drive API v3** — File discovery, folder management, file operations, uploads/downloads
3. **Google Sheets API v4** — Range read/write, append, clear, spreadsheet creation, sheet management
4. **Google Slides API v1** — Presentation CRUD, slide management, element creation (shapes, images, tables, text boxes)
5. **Gmail API v1** — Full email lifecycle (send, draft, read, search, labels, filters, threads, attachments)

## Authentication

Supports two authentication methods:
1. **OAuth 2.0 (Desktop App)** — Interactive browser-based consent flow with automatic token refresh and rotation handling
2. **Service Account** — Non-interactive authentication via `SERVICE_ACCOUNT_PATH` env var, with optional domain-wide delegation via `GOOGLE_IMPERSONATE_USER`

## Known Limitations

- Comment anchoring may not visibly appear in Google Docs UI (appears in "All Comments" sidebar only)
- `fixListFormatting` and `findElement` are experimental stubs
- Gmail attachment uploads may timeout for large files
- Rate limits apply per Google API quotas
