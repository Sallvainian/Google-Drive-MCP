# Google Docs MCP Server

FastMCP server with 99 tools for Google Docs, Sheets, Slides, Drive, and Gmail.

## Tool Categories

| Category | Count | Examples |
|----------|-------|----------|
| Docs | 6 | `readGoogleDoc`, `appendToGoogleDoc`, `insertText`, `insertHyperlink`, `deleteRange`, `listDocumentTabs` |
| Formatting | 3 | `applyTextStyle`, `applyParagraphStyle`, `formatMatchingText` |
| Structure | 7 | `insertTable`, `insertPageBreak`, `insertImageFromUrl`, `insertLocalImage`, `editTableCell`, `findElement`*, `fixListFormatting`* |
| Comments | 6 | `listComments`, `getComment`, `addComment`, `replyToComment`, `resolveComment`, `deleteComment` |
| Sheets | 8 | `readSpreadsheet`, `writeSpreadsheet`, `appendSpreadsheetRows`, `clearSpreadsheetRange`, `createSpreadsheet`, `listGoogleSheets` |
| Drive | 15 | `listGoogleDocs`, `searchGoogleDocs`, `getRecentGoogleDocs`, `getDocumentInfo`, `createFolder`, `moveFile`, `copyFile`, `uploadFile`, `createDocument`, `createFromTemplate` |
| Slides | 17 | `getPresentation`, `listSlides`, `getSlide`, `mapSlide`, `createPresentation`, `addSlide`, `duplicateSlide`, `addTextBox`, `addShape`, `addImage`, `addTable`, `editSlideTableCell`, `deleteSlide`, `deleteElement`, `updateSpeakerNotes`, `moveSlide`, `insertTextInElement` |
| Gmail | 34 | `send_email`, `read_email`, `search_emails`, `draft_email`, `reply_to_email`, `forward_email`, `get_thread`, `list_threads`, `trash_email`, `archive_email`, `mark_as_read`, `list_email_labels`, `create_label`, `create_filter`, `batch_modify_emails` |

*Not fully implemented

## Gmail Tools (34 total)

### Core Email Operations (7)
- `send_email` - Send email with attachments (plain text, HTML, multipart)
- `draft_email` - Create draft with attachments
- `read_email` - Get message content, headers, attachments
- `search_emails` - Gmail search syntax (`from:`, `is:unread`, `subject:`)
- `modify_email` - Add/remove labels from email
- `delete_email` - Permanently delete (irreversible)
- `download_attachment` - Save attachment to filesystem

### Label Management (5)
- `list_email_labels` - Get all system + user labels
- `create_label` - Create new label
- `update_label` - Rename or change visibility
- `delete_label` - Remove user label
- `get_or_create_label` - Idempotent label creation

### Batch Operations (2)
- `batch_modify_emails` - Bulk label changes
- `batch_delete_emails` - Bulk permanent delete

### Filter Management (5)
- `create_filter` - Create custom filter with criteria/actions
- `list_filters` - Get all filters
- `get_filter` - Get filter details
- `delete_filter` - Remove filter
- `create_filter_from_template` - Template-based (`fromSender`, `withSubject`, `withAttachments`, `largeEmails`, `containingText`, `mailingList`)

### Thread & Conversation (4)
- `get_thread` - Full conversation with all messages
- `list_threads` - Search/list threads
- `reply_to_email` - Reply maintaining thread
- `forward_email` - Forward to new recipients

### Message Actions (5)
- `trash_email` - Move to trash (recoverable)
- `untrash_email` - Restore from trash
- `archive_email` - Remove from inbox only
- `mark_as_read` - Mark message read
- `mark_as_unread` - Mark message unread

### Draft Management (5)
- `list_drafts` - Get all drafts
- `get_draft` - Get draft content
- `update_draft` - Modify draft
- `delete_draft` - Delete draft
- `send_draft` - Send existing draft

### Profile (1)
- `get_user_profile` - Get authenticated email address

## Known Limitations

### Google Docs
- **Comment anchoring:** Programmatically created comments appear in "All Comments" but aren't visibly anchored to text in the UI
- **Resolved status:** May not persist in Google Docs UI (Drive API limitation)
- **fixListFormatting:** Experimental, may not work reliably

### Gmail
- **First use:** After adding Gmail scope, delete existing tokens and re-authenticate
- **Attachment size:** Large attachments may timeout
- **Rate limits:** Batch operations respect Gmail API quotas

## Parameter Patterns

### Documents & Slides
- **Document ID:** Extract from URL: `docs.google.com/document/d/DOCUMENT_ID/edit`
- **Presentation ID:** Extract from URL: `docs.google.com/presentation/d/PRESENTATION_ID/edit`
- **Text targeting:** Use `textToFind` + `matchInstance` OR `startIndex`/`endIndex`
- **Colors:** Hex format `#RRGGBB` or `#RGB`
- **Alignment:** `START`, `END`, `CENTER`, `JUSTIFIED` (not LEFT/RIGHT)
- **Indices:** 1-based for Docs, 0-based for Slides positions
- **Tabs:** Optional `tabId` parameter (defaults to first tab)
- **Slides positioning:** All positions and sizes in points (1 inch = 72 points)
- **Slides workflow:** Use `mapSlide` BEFORE adding elements to see dimensions and available space
- **Text box styling:** `addTextBox` supports `fontSize` (points), `fontFamily`, `bold`, `italic`, `textColor` (hex)
- **getSlide modes:** Single slide by `pageObjectId`, OR range by `startIndex`/`endIndex` (0-based, exclusive end)

### Gmail
- **Message ID:** From `search_emails` or `list_threads` results
- **Thread ID:** From message results, used for `reply_to_email`
- **Search syntax:** Standard Gmail operators (`from:`, `to:`, `subject:`, `is:`, `has:`, `in:`, `label:`)
- **Labels:** Use label ID (from `list_email_labels`) or name
- **MIME types:** `text/plain`, `text/html`, `multipart/alternative`

## Source Files (for implementation details)

| File | Contains |
|------|----------|
| `src/types.ts` | Zod schemas, hex color validation, style parameter definitions, Gmail schemas |
| `src/googleDocsApiHelpers.ts` | `findTextRange`, `executeBatchUpdate`, style request builders |
| `src/googleSheetsApiHelpers.ts` | A1 notation parsing, range operations |
| `src/googleSlidesApiHelpers.ts` | EMU conversion, shape/table builders, batch update helpers |
| `src/googleGmailApiHelpers.ts` | Email creation, MIME handling, message formatting, batch operations |
| `src/gmailLabelManager.ts` | Label CRUD operations, system labels |
| `src/gmailFilterManager.ts` | Filter CRUD, template-based filter creation |
| `src/server.ts` | All 99 tool definitions with full parameter schemas |

## See Also

- `README.md` - Setup instructions and usage examples
- `SAMPLE_TASKS.md` - 15 example workflows
