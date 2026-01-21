# Developer Guide

## Project: Google Docs MCP Server
**Generated**: 2026-01-20

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
├── server.ts          # Add new tools here
├── auth.ts            # Authentication logic
├── types.ts           # Add new parameter schemas here
├── googleDocsApiHelpers.ts    # Docs-specific helpers
└── googleSheetsApiHelpers.ts  # Sheets-specific helpers
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
2. Add scope to `SCOPES` array in `auth.ts`:
   ```typescript
   const SCOPES = [
     'https://www.googleapis.com/auth/documents',
     'https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive.file',
     'https://www.googleapis.com/auth/NEW_SCOPE'  // Add here
   ];
   ```
3. Delete `token.json` and re-authorize

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
3. Run specific tool:
   ```bash
   # Using MCP inspector or Claude Desktop
   # Tool: readGoogleDoc
   # Args: { "documentId": "your-test-doc-id", "format": "text" }
   ```

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

---

## Troubleshooting

### Common Issues

| Issue | Cause | Solution |
|-------|-------|----------|
| "Invalid grant" | Expired/revoked token | Delete `token.json`, re-authorize |
| "Quota exceeded" | Too many API calls | Add delays, use batch operations |
| "Permission denied" | Missing scope or access | Check OAuth scopes, document sharing |
| "Invalid range" | Wrong index values | Use 1-based indexing, check document length |
| "Not implemented" | Feature stub | Check `NotImplementedError` in tool definition |

### Debug Mode

Enable verbose logging:

```typescript
// In server.ts
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  console.error('Stack:', error.stack);
});
```
