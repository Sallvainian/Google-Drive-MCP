# Developer Guide

## Project: Google Docs MCP Server
**Generated**: 2026-01-22

---

## Quick Start for Contributors

### Prerequisites
- Node.js 18+
- npm or yarn
- Google Cloud Console access
- TypeScript knowledge

### Initial Setup

```bash
# Clone repository
git clone https://github.com/a-bonus/google-docs-mcp.git
cd mcp-googledocs-server

# Install dependencies
npm install

# Build TypeScript
npm run build

# First run (requires OAuth setup)
node dist/server.js
```

---

## Development Workflow

### Build Commands

```bash
# Compile TypeScript to JavaScript
npm run build

# Run tests
npm test

# Start server directly
npm start

# Watch mode (requires ts-node-dev)
npm run dev
```

### Project Structure

```
src/
├── server.ts                    # 92 tool definitions (main entry point)
├── auth.ts                      # OAuth2 and service account authentication
├── types.ts                     # Zod schemas and TypeScript types
├── googleDocsApiHelpers.ts      # Docs API helpers (text, tables, images)
├── googleSheetsApiHelpers.ts    # Sheets API helpers (A1 notation, ranges)
├── googleSlidesApiHelpers.ts    # Slides API helpers (EMU conversion, elements)
├── googleGmailApiHelpers.ts     # Gmail API helpers (MIME, attachments)
├── gmailLabelManager.ts         # Label CRUD operations
└── gmailFilterManager.ts        # Filter management and templates
```

---

## Adding a New Tool

### Step 1: Define Parameters in `types.ts`

```typescript
// src/types.ts

// Add new parameter schema
export const MyNewToolParameters = z.object({
  documentId: z.string().describe('The Google Document ID'),
  myParam: z.string().min(1).describe('Description of parameter'),
  optionalParam: z.number().optional().describe('Optional parameter'),
});

// Export type inference
export type MyNewToolArgs = z.infer<typeof MyNewToolParameters>;
```

### Step 2: Add Helper Function (if needed)

```typescript
// src/googleDocsApiHelpers.ts

export async function myNewHelper(
  docs: docs_v1.Docs,
  documentId: string,
  myParam: string
): Promise<void> {
  try {
    const requests: docs_v1.Schema$Request[] = [
      {
        // Build your request here
      }
    ];

    await executeBatchUpdate(docs, documentId, requests);
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Document not found (ID: ${documentId}).`);
    }
    throw new UserError(`Operation failed: ${error.message}`);
  }
}
```

### Step 3: Register Tool in `server.ts`

```typescript
// src/server.ts

server.addTool({
  name: 'myNewTool',
  description: 'Clear description of what the tool does. Include examples if helpful.',
  parameters: MyNewToolParameters,
  execute: async (args: MyNewToolArgs, { log }) => {
    const docs = await getDocsClient();
    log.info(`Executing myNewTool for doc: ${args.documentId}`);

    try {
      await myNewHelper(docs, args.documentId, args.myParam);
      return `Successfully completed operation on document ${args.documentId}.`;
    } catch (error: any) {
      log.error(`Error in myNewTool: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed: ${error.message || 'Unknown error'}`);
    }
  }
});
```

### Step 4: Add Tests

```javascript
// tests/myNewTool.test.js

import { myNewHelper } from '../dist/googleDocsApiHelpers.js';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

describe('myNewHelper', () => {
  it('should perform expected operation', async () => {
    const mockDocs = {
      documents: {
        batchUpdate: mock.fn(async () => ({ data: {} }))
      }
    };

    await myNewHelper(mockDocs, 'doc123', 'testParam');

    assert.strictEqual(mockDocs.documents.batchUpdate.mock.calls.length, 1);
  });
});
```

---

## Error Handling Patterns

### Standard Error Pattern

```typescript
try {
  // API call
} catch (error: any) {
  log.error(`Error: ${error.message}`);

  // Re-throw UserError as-is
  if (error instanceof UserError) throw error;

  // Map common API errors
  if (error.code === 404) {
    throw new UserError(`Resource not found (ID: ${args.id}).`);
  }
  if (error.code === 403) {
    throw new UserError(`Permission denied for resource (ID: ${args.id}).`);
  }

  // Generic fallback
  throw new UserError(`Operation failed: ${error.message || 'Unknown error'}`);
}
```

### NotImplementedError for Stubs

```typescript
import { NotImplementedError } from './types.js';

execute: async (args, { log }) => {
  log.warn('This tool is not yet implemented.');
  throw new NotImplementedError('Feature not yet implemented.');
}
```

---

## Working with Google Docs API

### Batch Updates

The Docs API uses batch updates for modifications:

```typescript
const requests: docs_v1.Schema$Request[] = [
  {
    insertText: {
      location: { index: 1 },
      text: 'Hello World'
    }
  },
  {
    updateTextStyle: {
      range: { startIndex: 1, endIndex: 12 },
      textStyle: { bold: true },
      fields: 'bold'
    }
  }
];

await docs.documents.batchUpdate({
  documentId: 'your-doc-id',
  requestBody: { requests }
});
```

### Index Handling

- Google Docs uses 1-based indexing
- Ranges are `[start, end)` (end is exclusive)
- The document always ends with a newline at `endIndex`

```typescript
// To insert at the very beginning
const location = { index: 1 };

// To insert at the end (before final newline)
const endIndex = lastElement.endIndex - 1;
```

### Tab Support

Multi-tab documents require special handling:

```typescript
// Include tabs content when reading
const doc = await docs.documents.get({
  documentId: args.documentId,
  includeTabsContent: true
});

// Find specific tab
const tab = findTabById(doc.data, args.tabId);

// Insert with tab context
const request = {
  insertText: {
    location: {
      index: args.index,
      tabId: args.tabId  // Include tabId in location
    },
    text: args.text
  }
};
```

---

## Working with Google Sheets API

### A1 Notation

```typescript
// Single cell
'A1'

// Range
'A1:B10'

// With sheet name
'Sheet1!A1:B10'

// Entire column
'A:A'

// Entire row
'1:1'
```

### Value Input Options

```typescript
// RAW: Store values exactly as provided
await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: 'A1',
  valueInputOption: 'RAW',
  requestBody: { values: [['=SUM(A1:A10)']] }  // Stored as text
});

// USER_ENTERED: Parse as if user typed
await sheets.spreadsheets.values.update({
  spreadsheetId,
  range: 'A1',
  valueInputOption: 'USER_ENTERED',
  requestBody: { values: [['=SUM(A1:A10)']] }  // Evaluated as formula
});
```

---

## Working with Google Slides API

### EMU (English Metric Units)

Google Slides uses EMUs for precise positioning. The helper functions handle conversion:

```typescript
import { emuFromPoints, pointsFromEmu } from './googleSlidesApiHelpers.js';

// Convert points to EMU (1 point = 12700 EMU)
const emu = emuFromPoints(72);  // 72 points = 1 inch = 914400 EMU

// Convert EMU to points
const pts = pointsFromEmu(914400);  // = 72 points
```

### Standard Slide Dimensions

```typescript
// Default presentation size in points
const STANDARD_WIDTH = 720;   // 10 inches
const STANDARD_HEIGHT = 540;  // 7.5 inches

// Common element sizes
const FULL_WIDTH_BOX = { width: 680, height: 60, x: 20, y: 20 };
const CENTERED_CONTENT = { width: 600, height: 400, x: 60, y: 100 };
```

### Adding Elements to Slides

```typescript
import {
  buildCreateShapeRequest,
  buildInsertTextRequest,
  buildUpdateTextStyleRequest,
  executeBatchUpdate
} from './googleSlidesApiHelpers.js';

// Create a text box
const createShapeRequest = buildCreateShapeRequest(
  pageId,           // Slide ID
  'TEXT_BOX',       // Shape type
  100,              // x position (points)
  100,              // y position (points)
  300,              // width (points)
  50,               // height (points)
  'myTextBox_001'   // Object ID (optional)
);

// Insert text into the shape
const insertTextRequest = buildInsertTextRequest(
  'myTextBox_001',
  'Hello, Slides!',
  0                 // Insertion index
);

// Style the text
const styleRequest = buildUpdateTextStyleRequest(
  'myTextBox_001',
  {
    bold: true,
    fontSize: 24,
    foregroundColor: '#FF0000'
  },
  'ALL'             // Range type: ALL, FIXED_RANGE, FROM_START_INDEX
);

await executeBatchUpdate(slides, presentationId, [
  createShapeRequest,
  insertTextRequest,
  styleRequest
]);
```

### Shape Types

Supported shape types for `addShape` tool:

```typescript
const SHAPE_TYPES = [
  'RECTANGLE', 'ROUND_RECTANGLE', 'ELLIPSE',
  'TRIANGLE', 'RIGHT_TRIANGLE', 'PARALLELOGRAM', 'TRAPEZOID',
  'PENTAGON', 'HEXAGON', 'HEPTAGON', 'OCTAGON',
  'STAR_4', 'STAR_5', 'STAR_6', 'STAR_8', 'STAR_10', 'STAR_12',
  'ARROW_EAST', 'ARROW_NORTH', 'ARROW_NORTH_EAST',
  'CLOUD', 'HEART', 'PLUS'
];
```

### Working with Tables

```typescript
import { buildCreateTableRequest } from './googleSlidesApiHelpers.js';

// Create a 3x4 table
const tableRequest = buildCreateTableRequest(
  pageId,
  3,      // rows
  4,      // columns
  50,     // x position
  150,    // y position
  620,    // width
  300,    // height
  'myTable_001'
);
```

### Speaker Notes

```typescript
import { getSpeakerNotesShapeId } from './googleSlidesApiHelpers.js';

// Get the speaker notes shape ID for a slide
const slide = await getSlide(slides, presentationId, pageObjectId);
const notesShapeId = getSpeakerNotesShapeId(slide);

if (notesShapeId) {
  // Update speaker notes
  await executeBatchUpdate(slides, presentationId, [
    { deleteText: { objectId: notesShapeId, textRange: { type: 'ALL' } } },
    { insertText: { objectId: notesShapeId, text: 'New speaker notes' } }
  ]);
}
```

---

## Working with Gmail API

### Email Creation

```typescript
import {
  createSimpleEmail,
  createEmailWithAttachments,
  sendEmail
} from './googleGmailApiHelpers.js';

// Simple email
const rawEmail = createSimpleEmail({
  to: ['recipient@example.com'],
  subject: 'Test Email',
  body: 'Hello from the MCP server!',
  mimeType: 'text/plain'
});

await sendEmail(gmail, rawEmail);

// Email with attachments
const rawEmailWithAttachments = await createEmailWithAttachments({
  to: ['recipient@example.com'],
  subject: 'Report Attached',
  body: '<h1>Please see attachment</h1>',
  mimeType: 'text/html',
  attachments: ['/path/to/report.pdf']
});

await sendEmail(gmail, rawEmailWithAttachments);
```

### MIME Type Support

```typescript
// Plain text
{ mimeType: 'text/plain', body: 'Plain text content' }

// HTML
{ mimeType: 'text/html', body: '<h1>HTML content</h1>' }

// Multipart (both plain and HTML)
{
  mimeType: 'multipart/alternative',
  body: 'Plain text version',
  htmlBody: '<h1>HTML version</h1>'
}
```

### Search Queries

Gmail search uses standard Gmail query syntax:

```typescript
import { searchMessages } from './googleGmailApiHelpers.js';

// Search examples
const results = await searchMessages(gmail, {
  query: 'from:sender@example.com is:unread',
  maxResults: 50,
  includeSpamTrash: false
});

// Common query operators
// from:email      - Messages from sender
// to:email        - Messages to recipient
// subject:text    - Subject contains text
// is:unread       - Unread messages
// is:starred      - Starred messages
// has:attachment  - Has attachments
// in:inbox        - In inbox
// label:name      - Has label
// after:YYYY/MM/DD - After date
// before:YYYY/MM/DD - Before date
// larger:size     - Larger than (e.g., larger:5M)
```

### Thread Handling

```typescript
import { getThread, getMessage } from './googleGmailApiHelpers.js';

// Get full conversation
const thread = await getThread(gmail, threadId);
console.log(`Thread has ${thread.messages.length} messages`);

// Reply to maintain thread
const replyEmail = createSimpleEmail({
  to: ['recipient@example.com'],
  subject: 'Re: Original Subject',
  body: 'Reply content',
  inReplyTo: originalMessageId,  // Required for threading
  threadId: threadId             // Required for threading
});
```

### Label Management

```typescript
import {
  listLabels,
  createLabel,
  getOrCreateLabel,
  resolveLabelIds
} from './gmailLabelManager.js';

// List all labels
const labels = await listLabels(gmail);

// Create a new label
const newLabel = await createLabel(gmail, {
  name: 'My Custom Label',
  labelListVisibility: 'labelShow',
  messageListVisibility: 'show'
});

// Idempotent label creation (returns existing if found)
const label = await getOrCreateLabel(gmail, { name: 'Project X' });

// Resolve label names to IDs
const ids = await resolveLabelIds(gmail, ['INBOX', 'My Custom Label']);
```

### System Labels

```typescript
import { SYSTEM_LABELS } from './gmailLabelManager.js';

// System labels (cannot be deleted)
// INBOX, SPAM, TRASH, UNREAD, STARRED, IMPORTANT,
// SENT, DRAFT, CATEGORY_PERSONAL, CATEGORY_SOCIAL,
// CATEGORY_PROMOTIONS, CATEGORY_UPDATES, CATEGORY_FORUMS
```

### Filter Management

```typescript
import {
  createFilter,
  createFilterFromTemplate,
  listFilters
} from './gmailFilterManager.js';

// Create custom filter
const filter = await createFilter(gmail, {
  criteria: {
    from: 'newsletter@example.com',
    hasAttachment: false
  },
  action: {
    addLabelIds: ['Label_123'],
    removeLabelIds: ['INBOX']  // Archive
  }
});

// Use template for common scenarios
const templateFilter = await createFilterFromTemplate(
  gmail,
  'fromSender',
  {
    senderEmail: 'important@client.com',
    labelIds: ['Label_Priority'],
    markImportant: true
  }
);

// Available templates:
// - fromSender: Filter by sender email
// - withSubject: Filter by subject text
// - withAttachments: Filter messages with attachments
// - largeEmails: Filter by size threshold
// - containingText: Filter by body content
// - mailingList: Filter mailing list messages
```

### Batch Operations

```typescript
import {
  batchModifyMessages,
  batchDeleteMessages
} from './googleGmailApiHelpers.js';

// Bulk label changes (processed in batches of 50 by default)
const result = await batchModifyMessages(
  gmail,
  messageIds,
  ['Label_Processed'],    // Labels to add
  ['INBOX'],              // Labels to remove
  50                      // Batch size
);

console.log(`Processed: ${result.successful}, Failed: ${result.failed}`);

// Bulk delete (irreversible!)
const deleteResult = await batchDeleteMessages(gmail, messageIds);
```

---

## Authentication Development

### Testing OAuth Flow

```bash
# Delete existing token to force re-auth
rm token.json

# Run server - will prompt for authorization
node dist/server.js
```

### Testing Service Account

```bash
# Set environment variables
export SERVICE_ACCOUNT_PATH="/path/to/service-account.json"
export GOOGLE_IMPERSONATE_USER="user@domain.com"

# Run server
node dist/server.js
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

### Debugging Auth Issues

```typescript
// Add debug logging to auth.ts
console.error('Loading credentials from:', CREDENTIALS_PATH);
console.error('Token path:', TOKEN_PATH);
console.error('Service account:', process.env.SERVICE_ACCOUNT_PATH);
```

---

## Common Development Tasks

### Adding a New API Scope

1. Update OAuth consent screen in Google Cloud Console
2. Add scope to `SCOPES` array in `auth.ts`
3. Delete `token.json` and re-authorize

### Adding a New Google API

1. Import API in `server.ts`:
   ```typescript
   import { newapi_v1 } from 'googleapis';
   ```

2. Add client variable and getter:
   ```typescript
   let newApiClient: newapi_v1.NewApi | null = null;

   async function getNewApiClient(): Promise<newapi_v1.NewApi> {
     const auth = await initializeGoogleClient();
     if (!newApiClient) {
       newApiClient = google.newapi({ version: 'v1', auth });
     }
     return newApiClient;
   }
   ```

3. Create helper file `src/googleNewApiHelpers.ts`
4. Add required scopes to `auth.ts`
5. Implement tools using the new API

### Debugging Tool Execution

```typescript
execute: async (args, { log }) => {
  log.info(`Starting tool with args: ${JSON.stringify(args)}`);

  try {
    // ... operation
    log.info(`Operation completed successfully`);
  } catch (error) {
    log.error(`Error details: ${JSON.stringify(error)}`);
    log.error(`Stack: ${error.stack}`);
    throw error;
  }
}
```

### Testing with Real Documents

1. Create a test document in Google Docs
2. Get the document ID from the URL: `docs.google.com/document/d/DOCUMENT_ID/edit`
3. Run specific tool using MCP inspector or Claude Desktop

---

## Code Style Guidelines

### Parameter Descriptions

Always include clear descriptions in Zod schemas:

```typescript
// Good
documentId: z.string().describe('The unique identifier of the Google Document (found in the URL).')

// Bad
documentId: z.string()
```

### Error Messages

Use actionable error messages:

```typescript
// Good
throw new UserError(`Document not found (ID: ${id}). Check the document ID and try again.`);

// Bad
throw new UserError('Error');
```

### Logging

Use appropriate log levels:

```typescript
log.info('Starting operation');       // Normal flow
log.warn('Experimental feature');     // Caution
log.error('Operation failed');        // Errors
```

---

## Testing Strategy

### Unit Tests (Mocked)

```javascript
// Mock Google API clients
const mockDocs = {
  documents: {
    get: mock.fn(async () => ({ data: mockDocument })),
    batchUpdate: mock.fn(async () => ({ data: {} }))
  }
};
```

### Integration Tests (Real API)

```javascript
// Use real credentials for integration tests
// Run sparingly to avoid API quota issues
describe.skip('Integration Tests', () => {
  it('should read real document', async () => {
    // Test with actual API
  });
});
```

---

## Performance Considerations

### Batch Operations

Combine multiple operations into single batch updates:

```typescript
// Good: Single batch update
const requests = [
  { insertText: { ... } },
  { updateTextStyle: { ... } },
  { updateParagraphStyle: { ... } }
];
await executeBatchUpdate(docs, documentId, requests);

// Bad: Multiple API calls
await docs.documents.batchUpdate({ ... });  // Call 1
await docs.documents.batchUpdate({ ... });  // Call 2
await docs.documents.batchUpdate({ ... });  // Call 3
```

### Minimize Document Reads

Cache document content when multiple operations need it:

```typescript
// Good: Single read, multiple uses
const docInfo = await docs.documents.get({ documentId });
const endIndex = docInfo.data.body.content[...].endIndex;
// Use endIndex for multiple operations

// Bad: Multiple reads
const doc1 = await docs.documents.get({ documentId });
const doc2 = await docs.documents.get({ documentId });  // Redundant
```

### Gmail Batch Limits

Respect Gmail API quotas:

```typescript
// Good: Process in batches
await batchModifyMessages(gmail, messageIds, labels, [], 50);

// Bad: Process all at once (may hit rate limits)
for (const id of messageIds) {
  await modifyMessage(gmail, id, ...);  // One call per message
}
```

---

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| "Invalid grant" | Expired/revoked token | Delete `token.json`, re-authorize |
| "Quota exceeded" | Too many API calls | Add delays, use batch operations |
| "Permission denied" | Missing scope or access | Check OAuth scopes, document sharing |
| "Invalid range" | Wrong index values | Use 1-based indexing for Docs, check document length |
| "Not implemented" | Feature stub | Check `NotImplementedError` in tool definition |
| "Invalid EMU" | Wrong unit conversion | Use `emuFromPoints()` helper |
| "Thread not found" | Missing threadId | Ensure `threadId` from original message |
| "Label not found" | Using name instead of ID | Use `resolveLabelIds()` to convert |

### Debug Mode

Enable verbose logging:

```typescript
// In server.ts
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack:', error.stack);
});
```

---

## API Quick Reference

### Docs Tools (24)
- readGoogleDoc, listDocumentTabs, appendToGoogleDoc, insertText, deleteRange
- applyTextStyle, applyParagraphStyle, formatMatchingText
- insertTable, insertPageBreak, insertImageFromUrl, insertLocalImage
- listComments, getComment, addComment, replyToComment, resolveComment, deleteComment
- createFormattedDocument, insertFormattedContent, replaceDocumentContent, createFromTemplate

### Sheets Tools (8)
- readSpreadsheet, writeSpreadsheet, appendSpreadsheetRows, clearSpreadsheetRange
- getSpreadsheetInfo, addSpreadsheetSheet, createSpreadsheet, listGoogleSheets

### Drive Tools (13)
- listGoogleDocs, searchGoogleDocs, getRecentGoogleDocs, getDocumentInfo
- createFolder, listFolderContents, listAllFolders, getFolderInfo
- moveFile, copyFile, renameFile, deleteFile, createDocument

### Slides Tools (16)
- getPresentation, listSlides, getSlide, mapSlide, createPresentation
- addSlide, duplicateSlide, addTextBox, addShape, addImage, addTable
- deleteSlide, deleteElement, updateSpeakerNotes, moveSlide, insertTextInElement

### Gmail Tools (34)
- send_email, draft_email, read_email, search_emails, modify_email, delete_email, download_attachment
- list_email_labels, create_label, update_label, delete_label, get_or_create_label
- batch_modify_emails, batch_delete_emails
- create_filter, list_filters, get_filter, delete_filter, create_filter_from_template
- get_thread, list_threads, reply_to_email, forward_email
- trash_email, untrash_email, archive_email, mark_as_read, mark_as_unread
- list_drafts, get_draft, update_draft, delete_draft, send_draft
- get_user_profile
