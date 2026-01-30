// src/server.ts
import { FastMCP, UserError } from 'fastmcp';
import { z } from 'zod';
import { google, docs_v1, drive_v3, sheets_v4, slides_v1, gmail_v1 } from 'googleapis';
import { authorize } from './auth.js';
import { OAuth2Client } from 'google-auth-library';

// Import types and helpers
import {
DocumentIdParameter,
RangeParameters,
OptionalRangeParameters,
TextFindParameter,
TextStyleParameters,
TextStyleArgs,
ParagraphStyleParameters,
ParagraphStyleArgs,
ApplyTextStyleToolParameters, ApplyTextStyleToolArgs,
ApplyParagraphStyleToolParameters, ApplyParagraphStyleToolArgs,
NotImplementedError,
// Gmail types
MessageIdParameter,
ThreadIdParameter,
DraftIdParameter,
LabelIdParameter,
FilterIdParameter,
EmailRecipientsParameter,
EmailContentParameter,
EmailAttachmentsParameter,
GmailSearchParameter,
CreateLabelParameter,
UpdateLabelParameter,
ModifyLabelsParameter,
BatchModifyLabelsParameter,
CreateFilterParameter,
FilterTemplateParameter,
DownloadAttachmentParameter,
BatchDeleteParameter,
SendEmailParameter,
DraftEmailParameter,
LabelVisibilityParameter,
} from './types.js';
import * as GDocsHelpers from './googleDocsApiHelpers.js';
import * as SheetsHelpers from './googleSheetsApiHelpers.js';
import * as SlidesHelpers from './googleSlidesApiHelpers.js';
import * as GmailHelpers from './googleGmailApiHelpers.js';
import * as LabelManager from './gmailLabelManager.js';
import * as FilterManager from './gmailFilterManager.js';

let authClient: OAuth2Client | null = null;
let googleDocs: docs_v1.Docs | null = null;
let googleDrive: drive_v3.Drive | null = null;
let googleSheets: sheets_v4.Sheets | null = null;
let googleSlides: slides_v1.Slides | null = null;
let googleGmail: gmail_v1.Gmail | null = null;

// --- Initialization ---
async function initializeGoogleClient() {
if (googleDocs && googleDrive && googleSheets && googleSlides && googleGmail) return { authClient, googleDocs, googleDrive, googleSheets, googleSlides, googleGmail };
if (!authClient) { // Check authClient instead of googleDocs to allow re-attempt
try {
console.error("Attempting to authorize Google API client...");
const client = await authorize();
authClient = client; // Assign client here
googleDocs = google.docs({ version: 'v1', auth: authClient });
googleDrive = google.drive({ version: 'v3', auth: authClient });
googleSheets = google.sheets({ version: 'v4', auth: authClient });
googleSlides = google.slides({ version: 'v1', auth: authClient });
googleGmail = google.gmail({ version: 'v1', auth: authClient });
console.error("Google API client authorized successfully.");
} catch (error) {
console.error("FATAL: Failed to initialize Google API client:", error);
authClient = null; // Reset on failure
googleDocs = null;
googleDrive = null;
googleSheets = null;
googleSlides = null;
googleGmail = null;
// Decide if server should exit or just fail tools
throw new Error("Google client initialization failed. Cannot start server tools.");
}
}
// Ensure googleDocs, googleDrive, googleSheets, googleSlides, and googleGmail are set if authClient is valid
if (authClient && !googleDocs) {
googleDocs = google.docs({ version: 'v1', auth: authClient });
}
if (authClient && !googleDrive) {
googleDrive = google.drive({ version: 'v3', auth: authClient });
}
if (authClient && !googleSheets) {
googleSheets = google.sheets({ version: 'v4', auth: authClient });
}
if (authClient && !googleSlides) {
googleSlides = google.slides({ version: 'v1', auth: authClient });
}
if (authClient && !googleGmail) {
googleGmail = google.gmail({ version: 'v1', auth: authClient });
}

if (!googleDocs || !googleDrive || !googleSheets || !googleSlides || !googleGmail) {
throw new Error("Google Docs, Drive, Sheets, Slides, and Gmail clients could not be initialized.");
}

return { authClient, googleDocs, googleDrive, googleSheets, googleSlides, googleGmail };
}

// Set up process-level unhandled error/rejection handlers to prevent crashes
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
  // Don't exit process, just log the error and continue
  // This will catch timeout errors that might otherwise crash the server
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Promise Rejection:', reason);
  // Don't exit process, just log the error and continue
});

const server = new FastMCP({
  name: 'Ultimate Google Docs & Sheets MCP Server',
  version: '1.0.0'
});

// --- Helper to get Docs client within tools ---
async function getDocsClient() {
const { googleDocs: docs } = await initializeGoogleClient();
if (!docs) {
throw new UserError("Google Docs client is not initialized. Authentication might have failed during startup or lost connection.");
}
return docs;
}

// --- Helper to get Drive client within tools ---
async function getDriveClient() {
const { googleDrive: drive } = await initializeGoogleClient();
if (!drive) {
throw new UserError("Google Drive client is not initialized. Authentication might have failed during startup or lost connection.");
}
return drive;
}

// --- Helper to get Sheets client within tools ---
async function getSheetsClient() {
const { googleSheets: sheets } = await initializeGoogleClient();
if (!sheets) {
throw new UserError("Google Sheets client is not initialized. Authentication might have failed during startup or lost connection.");
}
return sheets;
}

// --- Helper to get Slides client within tools ---
async function getSlidesClient() {
const { googleSlides: slides } = await initializeGoogleClient();
if (!slides) {
throw new UserError("Google Slides client is not initialized. Authentication might have failed during startup or lost connection.");
}
return slides;
}

// --- Helper to get Gmail client within tools ---
async function getGmailClient() {
const { googleGmail: gmail } = await initializeGoogleClient();
if (!gmail) {
throw new UserError("Gmail client is not initialized. Authentication might have failed during startup or lost connection.");
}
return gmail;
}

// === HELPER FUNCTIONS ===

/**
 * Converts Google Docs JSON structure to Markdown format
 */
function convertDocsJsonToMarkdown(docData: any): string {
    let markdown = '';

    if (!docData.body?.content) {
        return 'Document appears to be empty.';
    }

    docData.body.content.forEach((element: any) => {
        if (element.paragraph) {
            markdown += convertParagraphToMarkdown(element.paragraph);
        } else if (element.table) {
            markdown += convertTableToMarkdown(element.table);
        } else if (element.sectionBreak) {
            markdown += '\n---\n\n'; // Section break as horizontal rule
        }
    });

    return markdown.trim();
}

/**
 * Converts a paragraph element to markdown
 */
function convertParagraphToMarkdown(paragraph: any): string {
    let text = '';
    let isHeading = false;
    let headingLevel = 0;
    let isList = false;
    let listType = '';

    // Check paragraph style for headings and lists
    if (paragraph.paragraphStyle?.namedStyleType) {
        const styleType = paragraph.paragraphStyle.namedStyleType;
        if (styleType.startsWith('HEADING_')) {
            isHeading = true;
            headingLevel = parseInt(styleType.replace('HEADING_', ''));
        } else if (styleType === 'TITLE') {
            isHeading = true;
            headingLevel = 1;
        } else if (styleType === 'SUBTITLE') {
            isHeading = true;
            headingLevel = 2;
        }
    }

    // Check for bullet lists
    if (paragraph.bullet) {
        isList = true;
        listType = paragraph.bullet.listId ? 'bullet' : 'bullet';
    }

    // Process text elements
    if (paragraph.elements) {
        paragraph.elements.forEach((element: any) => {
            if (element.textRun) {
                text += convertTextRunToMarkdown(element.textRun);
            }
        });
    }

    // Format based on style
    if (isHeading && text.trim()) {
        const hashes = '#'.repeat(Math.min(headingLevel, 6));
        return `${hashes} ${text.trim()}\n\n`;
    } else if (isList && text.trim()) {
        return `- ${text.trim()}\n`;
    } else if (text.trim()) {
        return `${text.trim()}\n\n`;
    }

    return '\n'; // Empty paragraph
}

/**
 * Converts a text run to markdown with formatting
 */
function convertTextRunToMarkdown(textRun: any): string {
    let text = textRun.content || '';

    if (textRun.textStyle) {
        const style = textRun.textStyle;

        // Apply formatting
        if (style.bold && style.italic) {
            text = `***${text}***`;
        } else if (style.bold) {
            text = `**${text}**`;
        } else if (style.italic) {
            text = `*${text}*`;
        }

        if (style.underline && !style.link) {
            // Markdown doesn't have native underline, use HTML
            text = `<u>${text}</u>`;
        }

        if (style.strikethrough) {
            text = `~~${text}~~`;
        }

        if (style.link?.url) {
            text = `[${text}](${style.link.url})`;
        }
    }

    return text;
}

/**
 * Converts a table to markdown format
 */
function convertTableToMarkdown(table: any): string {
    if (!table.tableRows || table.tableRows.length === 0) {
        return '';
    }

    let markdown = '\n';
    let isFirstRow = true;

    table.tableRows.forEach((row: any) => {
        if (!row.tableCells) return;

        let rowText = '|';
        row.tableCells.forEach((cell: any) => {
            let cellText = '';
            if (cell.content) {
                cell.content.forEach((element: any) => {
                    if (element.paragraph?.elements) {
                        element.paragraph.elements.forEach((pe: any) => {
                            if (pe.textRun?.content) {
                                cellText += pe.textRun.content.replace(/\n/g, ' ').trim();
                            }
                        });
                    }
                });
            }
            rowText += ` ${cellText} |`;
        });

        markdown += rowText + '\n';

        // Add header separator after first row
        if (isFirstRow) {
            let separator = '|';
            for (let i = 0; i < row.tableCells.length; i++) {
                separator += ' --- |';
            }
            markdown += separator + '\n';
            isFirstRow = false;
        }
    });

    return markdown + '\n';
}

// === TOOL DEFINITIONS ===

// --- Foundational Tools ---

server.addTool({
name: 'readGoogleDoc',
description: 'Reads the content of a specific Google Document, optionally returning structured data.',
parameters: DocumentIdParameter.extend({
format: z.enum(['text', 'json', 'markdown']).optional().default('text')
.describe("Output format: 'text' (plain text), 'json' (raw API structure, complex), 'markdown' (experimental conversion)."),
maxLength: z.number().optional().describe('Maximum character limit for text output. If not specified, returns full document content. Use this to limit very large documents.'),
tabId: z.string().optional().describe('The ID of the specific tab to read. If not specified, reads the first tab (or legacy document.body for documents without tabs).')
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Reading Google Doc: ${args.documentId}, Format: ${args.format}${args.tabId ? `, Tab: ${args.tabId}` : ''}`);

    try {
        // Determine if we need tabs content
        const needsTabsContent = !!args.tabId;

        const fields = args.format === 'json' || args.format === 'markdown'
            ? '*' // Get everything for structure analysis
            : 'body(content(paragraph(elements(textRun(content)))))'; // Just text content

        const res = await docs.documents.get({
            documentId: args.documentId,
            includeTabsContent: needsTabsContent,
            fields: needsTabsContent ? '*' : fields, // Get full document if using tabs
        });
        log.info(`Fetched doc: ${args.documentId}${args.tabId ? ` (tab: ${args.tabId})` : ''}`);

        // If tabId is specified, find the specific tab
        let contentSource: any;
        if (args.tabId) {
            const targetTab = GDocsHelpers.findTabById(res.data, args.tabId);
            if (!targetTab) {
                throw new UserError(`Tab with ID "${args.tabId}" not found in document.`);
            }
            if (!targetTab.documentTab) {
                throw new UserError(`Tab "${args.tabId}" does not have content (may not be a document tab).`);
            }
            contentSource = { body: targetTab.documentTab.body };
            log.info(`Using content from tab: ${targetTab.tabProperties?.title || 'Untitled'}`);
        } else {
            // Use the document body (backward compatible)
            contentSource = res.data;
        }

        if (args.format === 'json') {
            const jsonContent = JSON.stringify(contentSource, null, 2);
            // Apply length limit to JSON if specified
            if (args.maxLength && jsonContent.length > args.maxLength) {
                return jsonContent.substring(0, args.maxLength) + `\n... [JSON truncated: ${jsonContent.length} total chars]`;
            }
            return jsonContent;
        }

        if (args.format === 'markdown') {
            const markdownContent = convertDocsJsonToMarkdown(contentSource);
            const totalLength = markdownContent.length;
            log.info(`Generated markdown: ${totalLength} characters`);

            // Apply length limit to markdown if specified
            if (args.maxLength && totalLength > args.maxLength) {
                const truncatedContent = markdownContent.substring(0, args.maxLength);
                return `${truncatedContent}\n\n... [Markdown truncated to ${args.maxLength} chars of ${totalLength} total. Use maxLength parameter to adjust limit or remove it to get full content.]`;
            }

            return markdownContent;
        }

        // Default: Text format - extract all text content
        let textContent = '';
        let elementCount = 0;

        // Process all content elements from contentSource
        contentSource.body?.content?.forEach((element: any) => {
            elementCount++;

            // Handle paragraphs
            if (element.paragraph?.elements) {
                element.paragraph.elements.forEach((pe: any) => {
                    if (pe.textRun?.content) {
                        textContent += pe.textRun.content;
                    }
                });
            }

            // Handle tables
            if (element.table?.tableRows) {
                element.table.tableRows.forEach((row: any) => {
                    row.tableCells?.forEach((cell: any) => {
                        cell.content?.forEach((cellElement: any) => {
                            cellElement.paragraph?.elements?.forEach((pe: any) => {
                                if (pe.textRun?.content) {
                                    textContent += pe.textRun.content;
                                }
                            });
                        });
                    });
                });
            }
        });

        if (!textContent.trim()) return "Document found, but appears empty.";

        const totalLength = textContent.length;
        log.info(`Document contains ${totalLength} characters across ${elementCount} elements`);
        log.info(`maxLength parameter: ${args.maxLength || 'not specified'}`);

        // Apply length limit only if specified
        if (args.maxLength && totalLength > args.maxLength) {
            const truncatedContent = textContent.substring(0, args.maxLength);
            log.info(`Truncating content from ${totalLength} to ${args.maxLength} characters`);
            return `Content (truncated to ${args.maxLength} chars of ${totalLength} total):\n---\n${truncatedContent}\n\n... [Document continues for ${totalLength - args.maxLength} more characters. Use maxLength parameter to adjust limit or remove it to get full content.]`;
        }

        // Return full content
        const fullResponse = `Content (${totalLength} characters):\n---\n${textContent}`;
        const responseLength = fullResponse.length;
        log.info(`Returning full content: ${responseLength} characters in response (${totalLength} content + ${responseLength - totalLength} metadata)`);

        return fullResponse;

    } catch (error: any) {
         log.error(`Error reading doc ${args.documentId}: ${error.message || error}`);
         log.error(`Error details: ${JSON.stringify(error.response?.data || error)}`);
         // Handle errors thrown by helpers or API directly
         if (error instanceof UserError) throw error;
         if (error instanceof NotImplementedError) throw error;
         // Generic fallback for API errors not caught by helpers
          if (error.code === 404) throw new UserError(`Doc not found (ID: ${args.documentId}).`);
          if (error.code === 403) throw new UserError(`Permission denied for doc (ID: ${args.documentId}).`);
         // Extract detailed error information from Google API response
         const errorDetails = error.response?.data?.error?.message || error.message || 'Unknown error';
         const errorCode = error.response?.data?.error?.code || error.code;
         throw new UserError(`Failed to read doc: ${errorDetails}${errorCode ? ` (Code: ${errorCode})` : ''}`);
    }

},
});

server.addTool({
name: 'listDocumentTabs',
description: 'Lists all tabs in a Google Document, including their hierarchy, IDs, and structure.',
parameters: DocumentIdParameter.extend({
  includeContent: z.boolean().optional().default(false)
    .describe('Whether to include a content summary for each tab (character count).')
}),
execute: async (args, { log }) => {
  const docs = await getDocsClient();
  log.info(`Listing tabs for document: ${args.documentId}`);

  try {
    // Get document with tabs structure
    const res = await docs.documents.get({
      documentId: args.documentId,
      includeTabsContent: true,
      // Only get essential fields for tab listing
      fields: args.includeContent
        ? 'title,tabs'  // Get all tab data if we need content summary
        : 'title,tabs(tabProperties,childTabs)'  // Otherwise just structure
    });

    const docTitle = res.data.title || 'Untitled Document';

    // Get all tabs in a flat list with hierarchy info
    const allTabs = GDocsHelpers.getAllTabs(res.data);

    if (allTabs.length === 0) {
      // Shouldn't happen with new structure, but handle edge case
      return `Document "${docTitle}" appears to have no tabs (unexpected).`;
    }

    // Check if it's a single-tab or multi-tab document
    const isSingleTab = allTabs.length === 1;

    // Format the output
    let result = `**Document:** "${docTitle}"\n`;
    result += `**Total tabs:** ${allTabs.length}`;
    result += isSingleTab ? ' (single-tab document)\n\n' : '\n\n';

    if (!isSingleTab) {
      result += `**Tab Structure:**\n`;
      result += `${'─'.repeat(50)}\n\n`;
    }

    allTabs.forEach((tab: GDocsHelpers.TabWithLevel, index: number) => {
      const level = tab.level;
      const tabProperties = tab.tabProperties || {};
      const indent = '  '.repeat(level);

      // For single tab documents, show simplified info
      if (isSingleTab) {
        result += `**Default Tab:**\n`;
        result += `- Tab ID: ${tabProperties.tabId || 'Unknown'}\n`;
        result += `- Title: ${tabProperties.title || '(Untitled)'}\n`;
      } else {
        // For multi-tab documents, show hierarchy
        const prefix = level > 0 ? '└─ ' : '';
        result += `${indent}${prefix}**Tab ${index + 1}:** "${tabProperties.title || 'Untitled Tab'}"\n`;
        result += `${indent}   - ID: ${tabProperties.tabId || 'Unknown'}\n`;
        result += `${indent}   - Index: ${tabProperties.index !== undefined ? tabProperties.index : 'N/A'}\n`;

        if (tabProperties.parentTabId) {
          result += `${indent}   - Parent Tab ID: ${tabProperties.parentTabId}\n`;
        }
      }

      // Optionally include content summary
      if (args.includeContent && tab.documentTab) {
        const textLength = GDocsHelpers.getTabTextLength(tab.documentTab);
        const contentInfo = textLength > 0
          ? `${textLength.toLocaleString()} characters`
          : 'Empty';
        result += `${indent}   - Content: ${contentInfo}\n`;
      }

      if (!isSingleTab) {
        result += '\n';
      }
    });

    // Add usage hint for multi-tab documents
    if (!isSingleTab) {
      result += `\n💡 **Tip:** Use tab IDs with other tools to target specific tabs.`;
    }

    return result;

  } catch (error: any) {
    log.error(`Error listing tabs for doc ${args.documentId}: ${error.message || error}`);
    if (error.code === 404) throw new UserError(`Document not found (ID: ${args.documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${args.documentId}).`);
    throw new UserError(`Failed to list tabs: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'appendToGoogleDoc',
description: 'Appends text to the very end of a specific Google Document or tab.',
parameters: DocumentIdParameter.extend({
textToAppend: z.string().min(1).describe('The text to add to the end.'),
addNewlineIfNeeded: z.boolean().optional().default(true).describe("Automatically add a newline before the appended text if the doc doesn't end with one."),
tabId: z.string().optional().describe('The ID of the specific tab to append to. If not specified, appends to the first tab (or legacy document.body for documents without tabs).')
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Appending to Google Doc: ${args.documentId}${args.tabId ? ` (tab: ${args.tabId})` : ''}`);

    try {
        // Determine if we need tabs content
        const needsTabsContent = !!args.tabId;

        // Get the current end index
        const docInfo = await docs.documents.get({
            documentId: args.documentId,
            includeTabsContent: needsTabsContent,
            fields: needsTabsContent ? 'tabs' : 'body(content(endIndex)),documentStyle(pageSize)'
        });

        let endIndex = 1;
        let bodyContent: any;

        // If tabId is specified, find the specific tab
        if (args.tabId) {
            const targetTab = GDocsHelpers.findTabById(docInfo.data, args.tabId);
            if (!targetTab) {
                throw new UserError(`Tab with ID "${args.tabId}" not found in document.`);
            }
            if (!targetTab.documentTab) {
                throw new UserError(`Tab "${args.tabId}" does not have content (may not be a document tab).`);
            }
            bodyContent = targetTab.documentTab.body?.content;
        } else {
            bodyContent = docInfo.data.body?.content;
        }

        if (bodyContent) {
            const lastElement = bodyContent[bodyContent.length - 1];
            if (lastElement?.endIndex) {
                endIndex = lastElement.endIndex - 1; // Insert *before* the final newline of the doc typically
            }
        }

        // Simpler approach: Always assume insertion is needed unless explicitly told not to add newline
        const textToInsert = (args.addNewlineIfNeeded && endIndex > 1 ? '\n' : '') + args.textToAppend;

        if (!textToInsert) return "Nothing to append.";

        const location: any = { index: endIndex };
        if (args.tabId) {
            location.tabId = args.tabId;
        }

        const request: docs_v1.Schema$Request = { insertText: { location, text: textToInsert } };
        await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);

        log.info(`Successfully appended to doc: ${args.documentId}${args.tabId ? ` (tab: ${args.tabId})` : ''}`);
        return `Successfully appended text to ${args.tabId ? `tab ${args.tabId} in ` : ''}document ${args.documentId}.`;
    } catch (error: any) {
         log.error(`Error appending to doc ${args.documentId}: ${error.message || error}`);
         if (error instanceof UserError) throw error;
         if (error instanceof NotImplementedError) throw error;
         throw new UserError(`Failed to append to doc: ${error.message || 'Unknown error'}`);
    }

},
});

server.addTool({
name: 'insertText',
description: 'Inserts text at a specific index within the document body or a specific tab.',
parameters: DocumentIdParameter.extend({
textToInsert: z.string().min(1).describe('The text to insert.'),
index: z.number().int().min(1).describe('The index (1-based) where the text should be inserted.'),
tabId: z.string().optional().describe('The ID of the specific tab to insert into. If not specified, inserts into the first tab (or legacy document.body for documents without tabs).')
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting text in doc ${args.documentId} at index ${args.index}${args.tabId ? ` (tab: ${args.tabId})` : ''}`);
try {
    if (args.tabId) {
        // For tab-specific inserts, we need to verify the tab exists first
        const docInfo = await docs.documents.get({
            documentId: args.documentId,
            includeTabsContent: true,
            fields: 'tabs(tabProperties,documentTab)'
        });
        const targetTab = GDocsHelpers.findTabById(docInfo.data, args.tabId);
        if (!targetTab) {
            throw new UserError(`Tab with ID "${args.tabId}" not found in document.`);
        }
        if (!targetTab.documentTab) {
            throw new UserError(`Tab "${args.tabId}" does not have content (may not be a document tab).`);
        }

        // Insert with tabId
        const location: any = { index: args.index, tabId: args.tabId };
        const request: docs_v1.Schema$Request = { insertText: { location, text: args.textToInsert } };
        await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);
    } else {
        // Use existing helper for backward compatibility
        await GDocsHelpers.insertText(docs, args.documentId, args.textToInsert, args.index);
    }
    return `Successfully inserted text at index ${args.index}${args.tabId ? ` in tab ${args.tabId}` : ''}.`;
} catch (error: any) {
log.error(`Error inserting text in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert text: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'deleteRange',
description: 'Deletes content within a specified range (start index inclusive, end index exclusive) from the document or a specific tab.',
parameters: DocumentIdParameter.extend({
  startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
  endIndex: z.number().int().min(1).describe('The ending index of the text range (exclusive).'),
  tabId: z.string().optional().describe('The ID of the specific tab to delete from. If not specified, deletes from the first tab (or legacy document.body for documents without tabs).')
}).refine(data => data.endIndex > data.startIndex, {
  message: "endIndex must be greater than startIndex",
  path: ["endIndex"],
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Deleting range ${args.startIndex}-${args.endIndex} in doc ${args.documentId}${args.tabId ? ` (tab: ${args.tabId})` : ''}`);
if (args.endIndex <= args.startIndex) {
throw new UserError("End index must be greater than start index for deletion.");
}
try {
    // If tabId is specified, verify the tab exists
    if (args.tabId) {
        const docInfo = await docs.documents.get({
            documentId: args.documentId,
            includeTabsContent: true,
            fields: 'tabs(tabProperties,documentTab)'
        });
        const targetTab = GDocsHelpers.findTabById(docInfo.data, args.tabId);
        if (!targetTab) {
            throw new UserError(`Tab with ID "${args.tabId}" not found in document.`);
        }
        if (!targetTab.documentTab) {
            throw new UserError(`Tab "${args.tabId}" does not have content (may not be a document tab).`);
        }
    }

    const range: any = { startIndex: args.startIndex, endIndex: args.endIndex };
    if (args.tabId) {
        range.tabId = args.tabId;
    }

    const request: docs_v1.Schema$Request = {
        deleteContentRange: { range }
    };
    await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);
    return `Successfully deleted content in range ${args.startIndex}-${args.endIndex}${args.tabId ? ` in tab ${args.tabId}` : ''}.`;
} catch (error: any) {
    log.error(`Error deleting range in doc ${args.documentId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to delete range: ${error.message || 'Unknown error'}`);
}
}
});

// --- Advanced Formatting & Styling Tools ---

server.addTool({
name: 'applyTextStyle',
description: 'Applies character-level formatting (bold, color, font, etc.) to a specific range or found text.',
parameters: ApplyTextStyleToolParameters,
execute: async (args: ApplyTextStyleToolArgs, { log }) => {
const docs = await getDocsClient();
let { startIndex, endIndex } = args.target as any; // Will be updated if target is text

        log.info(`Applying text style in doc ${args.documentId}. Target: ${JSON.stringify(args.target)}, Style: ${JSON.stringify(args.style)}`);

        try {
            // Determine target range
            if ('textToFind' in args.target) {
                const range = await GDocsHelpers.findTextRange(docs, args.documentId, args.target.textToFind, args.target.matchInstance);
                if (!range) {
                    throw new UserError(`Could not find instance ${args.target.matchInstance} of text "${args.target.textToFind}".`);
                }
                startIndex = range.startIndex;
                endIndex = range.endIndex;
                log.info(`Found text "${args.target.textToFind}" (instance ${args.target.matchInstance}) at range ${startIndex}-${endIndex}`);
            }

            if (startIndex === undefined || endIndex === undefined) {
                 throw new UserError("Target range could not be determined.");
            }
             if (endIndex <= startIndex) {
                 throw new UserError("End index must be greater than start index for styling.");
            }

            // Build the request
            const requestInfo = GDocsHelpers.buildUpdateTextStyleRequest(startIndex, endIndex, args.style);
            if (!requestInfo) {
                 return "No valid text styling options were provided.";
            }

            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);
            return `Successfully applied text style (${requestInfo.fields.join(', ')}) to range ${startIndex}-${endIndex}.`;

        } catch (error: any) {
            log.error(`Error applying text style in doc ${args.documentId}: ${error.message || error}`);
            if (error instanceof UserError) throw error;
            if (error instanceof NotImplementedError) throw error; // Should not happen here
            throw new UserError(`Failed to apply text style: ${error.message || 'Unknown error'}`);
        }
    }

});

server.addTool({
name: 'applyParagraphStyle',
description: 'Applies paragraph-level formatting (alignment, spacing, named styles like Heading 1) to the paragraph(s) containing specific text, an index, or a range.',
parameters: ApplyParagraphStyleToolParameters,
execute: async (args: ApplyParagraphStyleToolArgs, { log }) => {
const docs = await getDocsClient();
let startIndex: number | undefined;
let endIndex: number | undefined;

        log.info(`Applying paragraph style to document ${args.documentId}`);
        log.info(`Style options: ${JSON.stringify(args.style)}`);
        log.info(`Target specification: ${JSON.stringify(args.target)}`);

        try {
            // STEP 1: Determine the target paragraph's range based on the targeting method
            if ('textToFind' in args.target) {
                // Find the text first
                log.info(`Finding text "${args.target.textToFind}" (instance ${args.target.matchInstance || 1})`);
                const textRange = await GDocsHelpers.findTextRange(
                    docs,
                    args.documentId,
                    args.target.textToFind,
                    args.target.matchInstance || 1
                );

                if (!textRange) {
                    throw new UserError(`Could not find "${args.target.textToFind}" in the document.`);
                }

                log.info(`Found text at range ${textRange.startIndex}-${textRange.endIndex}, now locating containing paragraph`);

                // Then find the paragraph containing this text
                const paragraphRange = await GDocsHelpers.getParagraphRange(
                    docs,
                    args.documentId,
                    textRange.startIndex
                );

                if (!paragraphRange) {
                    throw new UserError(`Found the text but could not determine the paragraph boundaries.`);
                }

                startIndex = paragraphRange.startIndex;
                endIndex = paragraphRange.endIndex;
                log.info(`Text is contained within paragraph at range ${startIndex}-${endIndex}`);

            } else if ('indexWithinParagraph' in args.target) {
                // Find paragraph containing the specified index
                log.info(`Finding paragraph containing index ${args.target.indexWithinParagraph}`);
                const paragraphRange = await GDocsHelpers.getParagraphRange(
                    docs,
                    args.documentId,
                    args.target.indexWithinParagraph
                );

                if (!paragraphRange) {
                    throw new UserError(`Could not find paragraph containing index ${args.target.indexWithinParagraph}.`);
                }

                startIndex = paragraphRange.startIndex;
                endIndex = paragraphRange.endIndex;
                log.info(`Located paragraph at range ${startIndex}-${endIndex}`);

            } else if ('startIndex' in args.target && 'endIndex' in args.target) {
                // Use directly provided range
                startIndex = args.target.startIndex;
                endIndex = args.target.endIndex;
                log.info(`Using provided paragraph range ${startIndex}-${endIndex}`);
            }

            // Verify that we have a valid range
            if (startIndex === undefined || endIndex === undefined) {
                throw new UserError("Could not determine target paragraph range from the provided information.");
            }

            if (endIndex <= startIndex) {
                throw new UserError(`Invalid paragraph range: end index (${endIndex}) must be greater than start index (${startIndex}).`);
            }

            // STEP 2: Build and apply the paragraph style request
            log.info(`Building paragraph style request for range ${startIndex}-${endIndex}`);
            const requestInfo = GDocsHelpers.buildUpdateParagraphStyleRequest(startIndex, endIndex, args.style);

            if (!requestInfo) {
                return "No valid paragraph styling options were provided.";
            }

            log.info(`Applying styles: ${requestInfo.fields.join(', ')}`);
            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);

            return `Successfully applied paragraph styles (${requestInfo.fields.join(', ')}) to the paragraph.`;

        } catch (error: any) {
            // Detailed error logging
            log.error(`Error applying paragraph style in doc ${args.documentId}:`);
            log.error(error.stack || error.message || error);

            if (error instanceof UserError) throw error;
            if (error instanceof NotImplementedError) throw error;

            // Provide a more helpful error message
            throw new UserError(`Failed to apply paragraph style: ${error.message || 'Unknown error'}`);
        }
    }
});

// --- Structure & Content Tools ---

server.addTool({
name: 'insertTable',
description: 'Inserts a new table with the specified dimensions at a given index.',
parameters: DocumentIdParameter.extend({
rows: z.number().int().min(1).describe('Number of rows for the new table.'),
columns: z.number().int().min(1).describe('Number of columns for the new table.'),
index: z.number().int().min(1).describe('The index (1-based) where the table should be inserted.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting ${args.rows}x${args.columns} table in doc ${args.documentId} at index ${args.index}`);
try {
await GDocsHelpers.createTable(docs, args.documentId, args.rows, args.columns, args.index);
// The API response contains info about the created table, but might be too complex to return here.
return `Successfully inserted a ${args.rows}x${args.columns} table at index ${args.index}.`;
} catch (error: any) {
log.error(`Error inserting table in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert table: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'editTableCell',
description: 'Edits the content and/or basic style of a specific table cell. Requires knowing table start index.',
parameters: DocumentIdParameter.extend({
tableStartIndex: z.number().int().min(1).describe("The starting index of the TABLE element itself (tricky to find, may require reading structure first)."),
rowIndex: z.number().int().min(0).describe("Row index (0-based)."),
columnIndex: z.number().int().min(0).describe("Column index (0-based)."),
textContent: z.string().optional().describe("Optional: New text content for the cell. Replaces existing content."),
// Combine basic styles for simplicity here. More advanced cell styling might need separate tools.
textStyle: TextStyleParameters.optional().describe("Optional: Text styles to apply."),
paragraphStyle: ParagraphStyleParameters.optional().describe("Optional: Paragraph styles (like alignment) to apply."),
// cellBackgroundColor: z.string().optional()... // Cell-specific styles are complex
}),
execute: async (args, { log }) => {
        const docs = await getDocsClient();
        log.info(`Editing cell (${args.rowIndex}, ${args.columnIndex}) in table starting at ${args.tableStartIndex}, doc ${args.documentId}`);

        try {
            // Step 1: Get cell range
            const cellRange = await GDocsHelpers.getTableCellRange(
                docs, args.documentId, args.tableStartIndex, args.rowIndex, args.columnIndex
            );

            if (!cellRange) {
                throw new UserError(`Could not find cell at row ${args.rowIndex}, column ${args.columnIndex}`);
            }

            const requests: docs_v1.Schema$Request[] = [];
            let newTextEndIndex = cellRange.contentStartIndex;

            // Step 2: Replace content if provided
            if (args.textContent !== undefined) {
                // Delete existing content (if any exists beyond the start)
                if (cellRange.contentEndIndex > cellRange.contentStartIndex) {
                    requests.push({
                        deleteContentRange: {
                            range: {
                                startIndex: cellRange.contentStartIndex,
                                endIndex: cellRange.contentEndIndex
                            }
                        }
                    });
                }
                // Insert new text
                if (args.textContent.length > 0) {
                    requests.push({
                        insertText: {
                            location: { index: cellRange.contentStartIndex },
                            text: args.textContent
                        }
                    });
                    newTextEndIndex = cellRange.contentStartIndex + args.textContent.length;
                }
            } else {
                newTextEndIndex = cellRange.contentEndIndex;
            }

            // Step 3: Apply text style
            if (args.textStyle && newTextEndIndex > cellRange.contentStartIndex) {
                const styleReq = GDocsHelpers.buildUpdateTextStyleRequest(
                    cellRange.contentStartIndex, newTextEndIndex, args.textStyle
                );
                if (styleReq) requests.push(styleReq.request);
            }

            // Step 4: Apply paragraph style
            if (args.paragraphStyle && newTextEndIndex > cellRange.contentStartIndex) {
                const paraReq = GDocsHelpers.buildUpdateParagraphStyleRequest(
                    cellRange.contentStartIndex, newTextEndIndex, args.paragraphStyle
                );
                if (paraReq) requests.push(paraReq.request);
            }

            // Step 5: Execute batch
            if (requests.length === 0) {
                return "No changes specified for the cell.";
            }

            await GDocsHelpers.executeBatchUpdate(docs, args.documentId, requests);
            return `Successfully edited cell (${args.rowIndex}, ${args.columnIndex}).`;
        } catch (error: any) {
            log.error(`Error editing cell (${args.rowIndex}, ${args.columnIndex}): ${error.message || error}`);
            if (error instanceof UserError) throw error;
            throw new UserError(`Failed to edit table cell: ${error.message || 'Unknown error'}`);
        }
    }
});

server.addTool({
name: 'insertPageBreak',
description: 'Inserts a page break at the specified index.',
parameters: DocumentIdParameter.extend({
index: z.number().int().min(1).describe('The index (1-based) where the page break should be inserted.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting page break in doc ${args.documentId} at index ${args.index}`);
try {
const request: docs_v1.Schema$Request = {
insertPageBreak: {
location: { index: args.index }
}
};
await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [request]);
return `Successfully inserted page break at index ${args.index}.`;
} catch (error: any) {
log.error(`Error inserting page break in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert page break: ${error.message || 'Unknown error'}`);
}
}
});

// --- Image Insertion Tools ---

server.addTool({
name: 'insertImageFromUrl',
description: 'Inserts an inline image into a Google Document from a publicly accessible URL.',
parameters: DocumentIdParameter.extend({
imageUrl: z.string().url().describe('Publicly accessible URL to the image (must be http:// or https://).'),
index: z.number().int().min(1).describe('The index (1-based) where the image should be inserted.'),
width: z.number().min(1).optional().describe('Optional: Width of the image in points.'),
height: z.number().min(1).optional().describe('Optional: Height of the image in points.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.info(`Inserting image from URL ${args.imageUrl} at index ${args.index} in doc ${args.documentId}`);

try {
await GDocsHelpers.insertInlineImage(
docs,
args.documentId,
args.imageUrl,
args.index,
args.width,
args.height
);

let sizeInfo = '';
if (args.width && args.height) {
sizeInfo = ` with size ${args.width}x${args.height}pt`;
}

return `Successfully inserted image from URL at index ${args.index}${sizeInfo}.`;
} catch (error: any) {
log.error(`Error inserting image in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to insert image: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'insertLocalImage',
description: 'Uploads a local image file to Google Drive and inserts it into a Google Document. The image will be uploaded to the same folder as the document (or optionally to a specified folder).',
parameters: DocumentIdParameter.extend({
localImagePath: z.string().describe('Absolute path to the local image file (supports .jpg, .jpeg, .png, .gif, .bmp, .webp, .svg).'),
index: z.number().int().min(1).describe('The index (1-based) where the image should be inserted in the document.'),
width: z.number().min(1).optional().describe('Optional: Width of the image in points.'),
height: z.number().min(1).optional().describe('Optional: Height of the image in points.'),
uploadToSameFolder: z.boolean().optional().default(true).describe('If true, uploads the image to the same folder as the document. If false, uploads to Drive root.'),
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
const drive = await getDriveClient();
log.info(`Uploading local image ${args.localImagePath} and inserting at index ${args.index} in doc ${args.documentId}`);

try {
// Get the document's parent folder if requested
let parentFolderId: string | undefined;
if (args.uploadToSameFolder) {
try {
const docInfo = await drive.files.get({
fileId: args.documentId,
fields: 'parents'
});
if (docInfo.data.parents && docInfo.data.parents.length > 0) {
parentFolderId = docInfo.data.parents[0];
log.info(`Will upload image to document's parent folder: ${parentFolderId}`);
}
} catch (folderError) {
log.warn(`Could not determine document's parent folder, using Drive root: ${folderError}`);
}
}

// Upload the image to Drive
log.info(`Uploading image to Drive...`);
const imageUrl = await GDocsHelpers.uploadImageToDrive(
drive,
args.localImagePath,
parentFolderId
);
log.info(`Image uploaded successfully, public URL: ${imageUrl}`);

// Insert the image into the document
await GDocsHelpers.insertInlineImage(
docs,
args.documentId,
imageUrl,
args.index,
args.width,
args.height
);

let sizeInfo = '';
if (args.width && args.height) {
sizeInfo = ` with size ${args.width}x${args.height}pt`;
}

return `Successfully uploaded image to Drive and inserted it at index ${args.index}${sizeInfo}.\nImage URL: ${imageUrl}`;
} catch (error: any) {
log.error(`Error uploading/inserting local image in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
throw new UserError(`Failed to upload/insert local image: ${error.message || 'Unknown error'}`);
}
}
});

// --- Intelligent Assistance Tools (Examples/Stubs) ---

server.addTool({
name: 'fixListFormatting',
description: 'EXPERIMENTAL: Attempts to detect paragraphs that look like lists (e.g., starting with -, *, 1.) and convert them to proper Google Docs bulleted or numbered lists. Best used on specific sections.',
parameters: DocumentIdParameter.extend({
// Optional range to limit the scope, otherwise scans whole doc (potentially slow/risky)
range: OptionalRangeParameters.optional().describe("Optional: Limit the fixing process to a specific range.")
}),
execute: async (args, { log }) => {
const docs = await getDocsClient();
log.warn(`Executing EXPERIMENTAL fixListFormatting for doc ${args.documentId}. Range: ${JSON.stringify(args.range)}`);
try {
await GDocsHelpers.detectAndFormatLists(docs, args.documentId, args.range?.startIndex, args.range?.endIndex);
return `Attempted to fix list formatting. Please review the document for accuracy.`;
} catch (error: any) {
log.error(`Error fixing list formatting in doc ${args.documentId}: ${error.message || error}`);
if (error instanceof UserError) throw error;
if (error instanceof NotImplementedError) throw error; // Expected if helper not implemented
throw new UserError(`Failed to fix list formatting: ${error.message || 'Unknown error'}`);
}
}
});

// === COMMENT TOOLS ===

server.addTool({
  name: 'listComments',
  description: 'Lists all comments in a Google Document.',
  parameters: DocumentIdParameter,
  execute: async (args, { log }) => {
    log.info(`Listing comments for document ${args.documentId}`);
    const docsClient = await getDocsClient();
    const driveClient = await getDriveClient();

    try {
      // First get the document to have context
      const doc = await docsClient.documents.get({ documentId: args.documentId });

      // Use Drive API v3 with proper fields to get quoted content
      const drive = google.drive({ version: 'v3', auth: authClient! });
      const response = await drive.comments.list({
        fileId: args.documentId,
        fields: 'comments(id,content,quotedFileContent,author,createdTime,resolved)',
        pageSize: 100
      });

      const comments = response.data.comments || [];

      if (comments.length === 0) {
        return 'No comments found in this document.';
      }

      // Format comments for display
      const formattedComments = comments.map((comment: any, index: number) => {
        const replies = comment.replies?.length || 0;
        const status = comment.resolved ? ' [RESOLVED]' : '';
        const author = comment.author?.displayName || 'Unknown';
        const date = comment.createdTime ? new Date(comment.createdTime).toLocaleDateString() : 'Unknown date';

        // Get the actual quoted text content
        const quotedText = comment.quotedFileContent?.value || 'No quoted text';
        const anchor = quotedText !== 'No quoted text' ? ` (anchored to: "${quotedText.substring(0, 100)}${quotedText.length > 100 ? '...' : ''}")` : '';

        let result = `\n${index + 1}. **${author}** (${date})${status}${anchor}\n   ${comment.content}`;

        if (replies > 0) {
          result += `\n   └─ ${replies} ${replies === 1 ? 'reply' : 'replies'}`;
        }

        result += `\n   Comment ID: ${comment.id}`;

        return result;
      }).join('\n');

      return `Found ${comments.length} comment${comments.length === 1 ? '' : 's'}:\n${formattedComments}`;

    } catch (error: any) {
      log.error(`Error listing comments: ${error.message || error}`);
      throw new UserError(`Failed to list comments: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'getComment',
  description: 'Gets a specific comment with its full thread of replies.',
  parameters: DocumentIdParameter.extend({
    commentId: z.string().describe('The ID of the comment to retrieve')
  }),
  execute: async (args, { log }) => {
    log.info(`Getting comment ${args.commentId} from document ${args.documentId}`);

    try {
      const drive = google.drive({ version: 'v3', auth: authClient! });
      const response = await drive.comments.get({
        fileId: args.documentId,
        commentId: args.commentId,
        fields: 'id,content,quotedFileContent,author,createdTime,resolved,replies(id,content,author,createdTime)'
      });

      const comment = response.data;
      const author = comment.author?.displayName || 'Unknown';
      const date = comment.createdTime ? new Date(comment.createdTime).toLocaleDateString() : 'Unknown date';
      const status = comment.resolved ? ' [RESOLVED]' : '';
      const quotedText = comment.quotedFileContent?.value || 'No quoted text';
      const anchor = quotedText !== 'No quoted text' ? `\nAnchored to: "${quotedText}"` : '';

      let result = `**${author}** (${date})${status}${anchor}\n${comment.content}`;

      // Add replies if any
      if (comment.replies && comment.replies.length > 0) {
        result += '\n\n**Replies:**';
        comment.replies.forEach((reply: any, index: number) => {
          const replyAuthor = reply.author?.displayName || 'Unknown';
          const replyDate = reply.createdTime ? new Date(reply.createdTime).toLocaleDateString() : 'Unknown date';
          result += `\n${index + 1}. **${replyAuthor}** (${replyDate})\n   ${reply.content}`;
        });
      }

      return result;

    } catch (error: any) {
      log.error(`Error getting comment: ${error.message || error}`);
      throw new UserError(`Failed to get comment: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'addComment',
  description: 'Adds a comment anchored to a specific text range in the document. NOTE: Due to Google API limitations, comments created programmatically appear in the "All Comments" list but are not visibly anchored to text in the document UI (they show "original content deleted"). However, replies, resolve, and delete operations work on all comments including manually-created ones.',
  parameters: DocumentIdParameter.extend({
    startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
    endIndex: z.number().int().min(1).describe('The ending index of the text range (exclusive).'),
    commentText: z.string().min(1).describe('The content of the comment.'),
  }).refine(data => data.endIndex > data.startIndex, {
    message: 'endIndex must be greater than startIndex',
    path: ['endIndex'],
  }),
  execute: async (args, { log }) => {
    log.info(`Adding comment to range ${args.startIndex}-${args.endIndex} in doc ${args.documentId}`);

    try {
      // First, get the text content that will be quoted
      const docsClient = await getDocsClient();
      const doc = await docsClient.documents.get({ documentId: args.documentId });

      // Extract the quoted text from the document
      let quotedText = '';
      const content = doc.data.body?.content || [];

      for (const element of content) {
        if (element.paragraph) {
          const elements = element.paragraph.elements || [];
          for (const textElement of elements) {
            if (textElement.textRun) {
              const elementStart = textElement.startIndex || 0;
              const elementEnd = textElement.endIndex || 0;

              // Check if this element overlaps with our range
              if (elementEnd > args.startIndex && elementStart < args.endIndex) {
                const text = textElement.textRun.content || '';
                const startOffset = Math.max(0, args.startIndex - elementStart);
                const endOffset = Math.min(text.length, args.endIndex - elementStart);
                quotedText += text.substring(startOffset, endOffset);
              }
            }
          }
        }
      }

      // Use Drive API v3 for comments
      const drive = google.drive({ version: 'v3', auth: authClient! });

      const response = await drive.comments.create({
        fileId: args.documentId,
        fields: 'id,content,quotedFileContent,author,createdTime,resolved',
        requestBody: {
          content: args.commentText,
          quotedFileContent: {
            value: quotedText,
            mimeType: 'text/html'
          },
          anchor: JSON.stringify({
            r: args.documentId,
            a: [{
              txt: {
                o: args.startIndex - 1,  // Drive API uses 0-based indexing
                l: args.endIndex - args.startIndex,
                ml: args.endIndex - args.startIndex
              }
            }]
          })
        }
      });

      return `Comment added successfully. Comment ID: ${response.data.id}`;

    } catch (error: any) {
      log.error(`Error adding comment: ${error.message || error}`);
      throw new UserError(`Failed to add comment: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'replyToComment',
  description: 'Adds a reply to an existing comment.',
  parameters: DocumentIdParameter.extend({
    commentId: z.string().describe('The ID of the comment to reply to'),
    replyText: z.string().min(1).describe('The content of the reply')
  }),
  execute: async (args, { log }) => {
    log.info(`Adding reply to comment ${args.commentId} in doc ${args.documentId}`);

    try {
      const drive = google.drive({ version: 'v3', auth: authClient! });

      const response = await drive.replies.create({
        fileId: args.documentId,
        commentId: args.commentId,
        fields: 'id,content,author,createdTime',
        requestBody: {
          content: args.replyText
        }
      });

      return `Reply added successfully. Reply ID: ${response.data.id}`;

    } catch (error: any) {
      log.error(`Error adding reply: ${error.message || error}`);
      throw new UserError(`Failed to add reply: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'resolveComment',
  description: 'Marks a comment as resolved. NOTE: Due to Google API limitations, the Drive API does not support resolving comments on Google Docs files. This operation will attempt to update the comment but the resolved status may not persist in the UI. Comments can be resolved manually in the Google Docs interface.',
  parameters: DocumentIdParameter.extend({
    commentId: z.string().describe('The ID of the comment to resolve')
  }),
  execute: async (args, { log }) => {
    log.info(`Resolving comment ${args.commentId} in doc ${args.documentId}`);

    try {
      const drive = google.drive({ version: 'v3', auth: authClient! });

      // First, get the current comment content (required by the API)
      const currentComment = await drive.comments.get({
        fileId: args.documentId,
        commentId: args.commentId,
        fields: 'content'
      });

      // Update with both content and resolved status
      await drive.comments.update({
        fileId: args.documentId,
        commentId: args.commentId,
        fields: 'id,resolved',
        requestBody: {
          content: currentComment.data.content,
          resolved: true
        }
      });

      // Verify the resolved status was set
      const verifyComment = await drive.comments.get({
        fileId: args.documentId,
        commentId: args.commentId,
        fields: 'resolved'
      });

      if (verifyComment.data.resolved) {
        return `Comment ${args.commentId} has been marked as resolved.`;
      } else {
        return `Attempted to resolve comment ${args.commentId}, but the resolved status may not persist in the Google Docs UI due to API limitations. The comment can be resolved manually in the Google Docs interface.`;
      }

    } catch (error: any) {
      log.error(`Error resolving comment: ${error.message || error}`);
      const errorDetails = error.response?.data?.error?.message || error.message || 'Unknown error';
      const errorCode = error.response?.data?.error?.code;
      throw new UserError(`Failed to resolve comment: ${errorDetails}${errorCode ? ` (Code: ${errorCode})` : ''}`);
    }
  }
});

server.addTool({
  name: 'deleteComment',
  description: 'Deletes a comment from the document.',
  parameters: DocumentIdParameter.extend({
    commentId: z.string().describe('The ID of the comment to delete')
  }),
  execute: async (args, { log }) => {
    log.info(`Deleting comment ${args.commentId} from doc ${args.documentId}`);

    try {
      const drive = google.drive({ version: 'v3', auth: authClient! });

      await drive.comments.delete({
        fileId: args.documentId,
        commentId: args.commentId
      });

      return `Comment ${args.commentId} has been deleted.`;

    } catch (error: any) {
      log.error(`Error deleting comment: ${error.message || error}`);
      throw new UserError(`Failed to delete comment: ${error.message || 'Unknown error'}`);
    }
  }
});

// --- Add Stubs for other advanced features ---
// (findElement, getDocumentMetadata, replaceText, list management, image handling, section breaks, footnotes, etc.)
// Example Stub:
server.addTool({
name: 'findElement',
description: 'Finds elements (paragraphs, tables, etc.) based on various criteria. (Not Implemented)',
parameters: DocumentIdParameter.extend({
// Define complex query parameters...
textQuery: z.string().optional(),
elementType: z.enum(['paragraph', 'table', 'list', 'image']).optional(),
// styleQuery...
}),
execute: async (args, { log }) => {
log.warn("findElement tool called but is not implemented.");
throw new NotImplementedError("Finding elements by complex criteria is not yet implemented.");
}
});

// --- Preserve the existing formatMatchingText tool for backward compatibility ---
server.addTool({
name: 'formatMatchingText',
description: 'Finds specific text within a Google Document and applies character formatting (bold, italics, color, etc.) to the specified instance.',
parameters: z.object({
  documentId: z.string().describe('The ID of the Google Document.'),
  textToFind: z.string().min(1).describe('The exact text string to find and format.'),
  matchInstance: z.number().int().min(1).optional().default(1).describe('Which instance of the text to format (1st, 2nd, etc.). Defaults to 1.'),
  // Re-use optional Formatting Parameters (SHARED)
  bold: z.boolean().optional().describe('Apply bold formatting.'),
  italic: z.boolean().optional().describe('Apply italic formatting.'),
  underline: z.boolean().optional().describe('Apply underline formatting.'),
  strikethrough: z.boolean().optional().describe('Apply strikethrough formatting.'),
  fontSize: z.number().min(1).optional().describe('Set font size (in points, e.g., 12).'),
  fontFamily: z.string().optional().describe('Set font family (e.g., "Arial", "Times New Roman").'),
  foregroundColor: z.string()
    .refine((color) => /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(color), {
      message: "Invalid hex color format (e.g., #FF0000 or #F00)"
    })
    .optional()
    .describe('Set text color using hex format (e.g., "#FF0000").'),
  backgroundColor: z.string()
    .refine((color) => /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/.test(color), {
      message: "Invalid hex color format (e.g., #00FF00 or #0F0)"
    })
    .optional()
    .describe('Set text background color using hex format (e.g., "#FFFF00").'),
  linkUrl: z.string().url().optional().describe('Make the text a hyperlink pointing to this URL.')
})
.refine(data => Object.keys(data).some(key => !['documentId', 'textToFind', 'matchInstance'].includes(key) && data[key as keyof typeof data] !== undefined), {
    message: "At least one formatting option (bold, italic, fontSize, etc.) must be provided."
}),
execute: async (args, { log }) => {
  // Adapt to use the new applyTextStyle implementation under the hood
  const docs = await getDocsClient();
  log.info(`Using formatMatchingText (legacy) for doc ${args.documentId}, target: "${args.textToFind}" (instance ${args.matchInstance})`);

  try {
    // Extract the style parameters
    const styleParams: TextStyleArgs = {};
    if (args.bold !== undefined) styleParams.bold = args.bold;
    if (args.italic !== undefined) styleParams.italic = args.italic;
    if (args.underline !== undefined) styleParams.underline = args.underline;
    if (args.strikethrough !== undefined) styleParams.strikethrough = args.strikethrough;
    if (args.fontSize !== undefined) styleParams.fontSize = args.fontSize;
    if (args.fontFamily !== undefined) styleParams.fontFamily = args.fontFamily;
    if (args.foregroundColor !== undefined) styleParams.foregroundColor = args.foregroundColor;
    if (args.backgroundColor !== undefined) styleParams.backgroundColor = args.backgroundColor;
    if (args.linkUrl !== undefined) styleParams.linkUrl = args.linkUrl;

    // Find the text range
    const range = await GDocsHelpers.findTextRange(docs, args.documentId, args.textToFind, args.matchInstance);
    if (!range) {
      throw new UserError(`Could not find instance ${args.matchInstance} of text "${args.textToFind}".`);
    }

    // Build and execute the request
    const requestInfo = GDocsHelpers.buildUpdateTextStyleRequest(range.startIndex, range.endIndex, styleParams);
    if (!requestInfo) {
      return "No valid text styling options were provided.";
    }

    await GDocsHelpers.executeBatchUpdate(docs, args.documentId, [requestInfo.request]);
    return `Successfully applied formatting to instance ${args.matchInstance} of "${args.textToFind}".`;
  } catch (error: any) {
    log.error(`Error in formatMatchingText for doc ${args.documentId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to format text: ${error.message || 'Unknown error'}`);
  }
}
});

// === GOOGLE DRIVE TOOLS ===

server.addTool({
name: 'listGoogleDocs',
description: 'Lists Google Documents from your Google Drive with optional filtering.',
parameters: z.object({
  maxResults: z.number().int().min(1).max(100).optional().default(20).describe('Maximum number of documents to return (1-100).'),
  query: z.string().optional().describe('Search query to filter documents by name or content.'),
  orderBy: z.enum(['name', 'modifiedTime', 'createdTime']).optional().default('modifiedTime').describe('Sort order for results.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Listing Google Docs. Query: ${args.query || 'none'}, Max: ${args.maxResults}, Order: ${args.orderBy}`);

try {
  // Build the query string for Google Drive API
  let queryString = "mimeType='application/vnd.google-apps.document' and trashed=false";
  if (args.query) {
    // Use name-only search to allow sorting - fullText search doesn't support orderBy
    queryString += ` and name contains '${args.query}'`;
  }

  const response = await drive.files.list({
    q: queryString,
    pageSize: args.maxResults,
    orderBy: args.orderBy === 'name' ? 'name' : args.orderBy,
    fields: 'files(id,name,modifiedTime,createdTime,size,webViewLink,owners(displayName,emailAddress))',
  });

  const files = response.data.files || [];

  if (files.length === 0) {
    return "No Google Docs found matching your criteria.";
  }

  let result = `Found ${files.length} Google Document(s):\n\n`;
  files.forEach((file, index) => {
    const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'Unknown';
    const owner = file.owners?.[0]?.displayName || 'Unknown';
    result += `${index + 1}. **${file.name}**\n`;
    result += `   ID: ${file.id}\n`;
    result += `   Modified: ${modifiedDate}\n`;
    result += `   Owner: ${owner}\n`;
    result += `   Link: ${file.webViewLink}\n\n`;
  });

  return result;
} catch (error: any) {
  log.error(`Error listing Google Docs: ${error.message || error}`);
  if (error.code === 403) {
    // Check if it's an API limitation vs permission error
    const msg = error.message?.toLowerCase() || '';
    if (msg.includes('sorting') || msg.includes('fulltext')) {
      throw new UserError("Google Drive API doesn't support sorting with full-text search. Try searching by name only.");
    }
    throw new UserError("Permission denied. Make sure you have granted Google Drive access to the application.");
  }
  throw new UserError(`Failed to list documents: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'searchGoogleDocs',
description: 'Searches for Google Documents by name, content, or other criteria.',
parameters: z.object({
  searchQuery: z.string().min(1).describe('Search term to find in document names or content.'),
  searchIn: z.enum(['name', 'content', 'both']).optional().default('both').describe('Where to search: document names, content, or both.'),
  maxResults: z.number().int().min(1).max(50).optional().default(10).describe('Maximum number of results to return.'),
  modifiedAfter: z.string().optional().describe('Only return documents modified after this date (ISO 8601 format, e.g., "2024-01-01").'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Searching Google Docs for: "${args.searchQuery}" in ${args.searchIn}`);

try {
  let queryString = "mimeType='application/vnd.google-apps.document' and trashed=false";

  // Add search criteria
  if (args.searchIn === 'name') {
    queryString += ` and name contains '${args.searchQuery}'`;
  } else if (args.searchIn === 'content') {
    queryString += ` and fullText contains '${args.searchQuery}'`;
  } else {
    queryString += ` and (name contains '${args.searchQuery}' or fullText contains '${args.searchQuery}')`;
  }

  // Add date filter if provided
  if (args.modifiedAfter) {
    queryString += ` and modifiedTime > '${args.modifiedAfter}'`;
  }

  // fullText search doesn't support orderBy - only apply sorting for name-only search
  const usesFullText = args.searchIn === 'content' || args.searchIn === 'both';

  const response = await drive.files.list({
    q: queryString,
    pageSize: args.maxResults,
    ...(usesFullText ? {} : { orderBy: 'modifiedTime desc' }),
    fields: 'files(id,name,modifiedTime,createdTime,webViewLink,owners(displayName),parents)',
  });

  const files = response.data.files || [];

  if (files.length === 0) {
    return `No Google Docs found containing "${args.searchQuery}".`;
  }

  let result = `Found ${files.length} document(s) matching "${args.searchQuery}":\n\n`;
  files.forEach((file, index) => {
    const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'Unknown';
    const owner = file.owners?.[0]?.displayName || 'Unknown';
    result += `${index + 1}. **${file.name}**\n`;
    result += `   ID: ${file.id}\n`;
    result += `   Modified: ${modifiedDate}\n`;
    result += `   Owner: ${owner}\n`;
    result += `   Link: ${file.webViewLink}\n\n`;
  });

  return result;
} catch (error: any) {
  log.error(`Error searching Google Docs: ${error.message || error}`);
  if (error.code === 403) {
    // Check if it's an API limitation vs permission error
    const msg = error.message?.toLowerCase() || '';
    if (msg.includes('sorting') || msg.includes('fulltext')) {
      throw new UserError("Google Drive API doesn't support sorting with full-text search. Results will be returned without specific ordering.");
    }
    throw new UserError("Permission denied. Make sure you have granted Google Drive access to the application.");
  }
  throw new UserError(`Failed to search documents: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'getRecentGoogleDocs',
description: 'Gets the most recently modified Google Documents.',
parameters: z.object({
  maxResults: z.number().int().min(1).max(50).optional().default(10).describe('Maximum number of recent documents to return.'),
  daysBack: z.number().int().min(1).max(365).optional().default(30).describe('Only show documents modified within this many days.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Getting recent Google Docs: ${args.maxResults} results, ${args.daysBack} days back`);

try {
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - args.daysBack);
  const cutoffDateStr = cutoffDate.toISOString();

  const queryString = `mimeType='application/vnd.google-apps.document' and trashed=false and modifiedTime > '${cutoffDateStr}'`;

  const response = await drive.files.list({
    q: queryString,
    pageSize: args.maxResults,
    orderBy: 'modifiedTime desc',
    fields: 'files(id,name,modifiedTime,createdTime,webViewLink,owners(displayName),lastModifyingUser(displayName))',
  });

  const files = response.data.files || [];

  if (files.length === 0) {
    return `No Google Docs found that were modified in the last ${args.daysBack} days.`;
  }

  let result = `${files.length} recently modified Google Document(s) (last ${args.daysBack} days):\n\n`;
  files.forEach((file, index) => {
    const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleString() : 'Unknown';
    const lastModifier = file.lastModifyingUser?.displayName || 'Unknown';
    const owner = file.owners?.[0]?.displayName || 'Unknown';

    result += `${index + 1}. **${file.name}**\n`;
    result += `   ID: ${file.id}\n`;
    result += `   Last Modified: ${modifiedDate} by ${lastModifier}\n`;
    result += `   Owner: ${owner}\n`;
    result += `   Link: ${file.webViewLink}\n\n`;
  });

  return result;
} catch (error: any) {
  log.error(`Error getting recent Google Docs: ${error.message || error}`);
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have granted Google Drive access to the application.");
  throw new UserError(`Failed to get recent documents: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'getDocumentInfo',
description: 'Gets detailed information about a specific Google Document.',
parameters: DocumentIdParameter,
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Getting info for document: ${args.documentId}`);

try {
  const response = await drive.files.get({
    fileId: args.documentId,
    // Note: 'permissions' and 'alternateLink' fields removed - they cause
    // "Invalid field selection" errors for Google Docs files
    fields: 'id,name,description,mimeType,size,createdTime,modifiedTime,webViewLink,owners(displayName,emailAddress),lastModifyingUser(displayName,emailAddress),shared,parents,version',
  });

  const file = response.data;

  if (!file) {
    throw new UserError(`Document with ID ${args.documentId} not found.`);
  }

  const createdDate = file.createdTime ? new Date(file.createdTime).toLocaleString() : 'Unknown';
  const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleString() : 'Unknown';
  const owner = file.owners?.[0];
  const lastModifier = file.lastModifyingUser;

  let result = `**Document Information:**\n\n`;
  result += `**Name:** ${file.name}\n`;
  result += `**ID:** ${file.id}\n`;
  result += `**Type:** Google Document\n`;
  result += `**Created:** ${createdDate}\n`;
  result += `**Last Modified:** ${modifiedDate}\n`;

  if (owner) {
    result += `**Owner:** ${owner.displayName} (${owner.emailAddress})\n`;
  }

  if (lastModifier) {
    result += `**Last Modified By:** ${lastModifier.displayName} (${lastModifier.emailAddress})\n`;
  }

  result += `**Shared:** ${file.shared ? 'Yes' : 'No'}\n`;
  result += `**View Link:** ${file.webViewLink}\n`;

  if (file.description) {
    result += `**Description:** ${file.description}\n`;
  }

  return result;
} catch (error: any) {
  log.error(`Error getting document info: ${error.message || error}`);
  if (error.code === 404) throw new UserError(`Document not found (ID: ${args.documentId}).`);
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have access to this document.");
  throw new UserError(`Failed to get document info: ${error.message || 'Unknown error'}`);
}
}
});

// === GOOGLE DRIVE FILE MANAGEMENT TOOLS ===

// --- Folder Management Tools ---

server.addTool({
name: 'createFolder',
description: 'Creates a new folder in Google Drive.',
parameters: z.object({
  name: z.string().min(1).describe('Name for the new folder.'),
  parentFolderId: z.string().optional().describe('Parent folder ID. If not provided, creates folder in Drive root.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Creating folder "${args.name}" ${args.parentFolderId ? `in parent ${args.parentFolderId}` : 'in root'}`);

try {
  const folderMetadata: drive_v3.Schema$File = {
    name: args.name,
    mimeType: 'application/vnd.google-apps.folder',
  };

  if (args.parentFolderId) {
    folderMetadata.parents = [args.parentFolderId];
  }

  const response = await drive.files.create({
    requestBody: folderMetadata,
    fields: 'id,name,parents,webViewLink',
  });

  const folder = response.data;
  return `Successfully created folder "${folder.name}" (ID: ${folder.id})\nLink: ${folder.webViewLink}`;
} catch (error: any) {
  log.error(`Error creating folder: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Parent folder not found. Check the parent folder ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to the parent folder.");
  throw new UserError(`Failed to create folder: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'listFolderContents',
description: 'Lists the contents of a specific folder in Google Drive.',
parameters: z.object({
  folderId: z.string().describe('ID of the folder to list contents of. Use "root" for the root Drive folder.'),
  includeSubfolders: z.boolean().optional().default(true).describe('Whether to include subfolders in results.'),
  includeFiles: z.boolean().optional().default(true).describe('Whether to include files in results.'),
  maxResults: z.number().int().min(1).max(100).optional().default(50).describe('Maximum number of items to return.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Listing contents of folder: ${args.folderId}`);

try {
  let queryString = `'${args.folderId}' in parents and trashed=false`;

  // Filter by type if specified
  if (!args.includeSubfolders && !args.includeFiles) {
    throw new UserError("At least one of includeSubfolders or includeFiles must be true.");
  }

  if (!args.includeSubfolders) {
    queryString += ` and mimeType!='application/vnd.google-apps.folder'`;
  } else if (!args.includeFiles) {
    queryString += ` and mimeType='application/vnd.google-apps.folder'`;
  }

  const response = await drive.files.list({
    q: queryString,
    pageSize: args.maxResults,
    orderBy: 'folder,name',
    fields: 'files(id,name,mimeType,size,modifiedTime,webViewLink,owners(displayName))',
  });

  const items = response.data.files || [];

  if (items.length === 0) {
    return "The folder is empty or you don't have permission to view its contents.";
  }

  let result = `Contents of folder (${items.length} item${items.length !== 1 ? 's' : ''}):\n\n`;

  // Separate folders and files
  const folders = items.filter(item => item.mimeType === 'application/vnd.google-apps.folder');
  const files = items.filter(item => item.mimeType !== 'application/vnd.google-apps.folder');

  // List folders first
  if (folders.length > 0 && args.includeSubfolders) {
    result += `**Folders (${folders.length}):**\n`;
    folders.forEach(folder => {
      result += `📁 ${folder.name} (ID: ${folder.id})\n`;
    });
    result += '\n';
  }

  // Then list files
  if (files.length > 0 && args.includeFiles) {
    result += `**Files (${files.length}):\n`;
    files.forEach(file => {
      const fileType = file.mimeType === 'application/vnd.google-apps.document' ? '📄' :
                      file.mimeType === 'application/vnd.google-apps.spreadsheet' ? '📊' :
                      file.mimeType === 'application/vnd.google-apps.presentation' ? '📈' : '📎';
      const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'Unknown';
      const owner = file.owners?.[0]?.displayName || 'Unknown';

      result += `${fileType} ${file.name}\n`;
      result += `   ID: ${file.id}\n`;
      result += `   Modified: ${modifiedDate} by ${owner}\n`;
      result += `   Link: ${file.webViewLink}\n\n`;
    });
  }

  return result;
} catch (error: any) {
  log.error(`Error listing folder contents: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Folder not found. Check the folder ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have access to this folder.");
  throw new UserError(`Failed to list folder contents: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'listAllFolders',
description: 'Recursively lists all folders in Google Drive as a tree structure. Only returns folders, not files.',
parameters: z.object({
  folderId: z.string().optional().default('root').describe('ID of the folder to start from. Defaults to "root".'),
  maxDepth: z.number().int().min(1).max(10).optional().default(5).describe('Maximum depth to traverse (1-10). Defaults to 5.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Listing all folders recursively from: ${args.folderId} with max depth: ${args.maxDepth}`);

interface FolderNode {
  id: string;
  name: string;
  children: FolderNode[];
}

async function getFoldersInFolder(parentId: string): Promise<FolderNode[]> {
  const response = await drive.files.list({
    q: `'${parentId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
    pageSize: 100,
    orderBy: 'name',
    fields: 'files(id,name)',
  });

  return (response.data.files || []).map(f => ({
    id: f.id!,
    name: f.name!,
    children: [],
  }));
}

async function buildTree(folderId: string, currentDepth: number): Promise<FolderNode[]> {
  if (currentDepth > args.maxDepth) {
    return [];
  }

  const folders = await getFoldersInFolder(folderId);

  for (const folder of folders) {
    folder.children = await buildTree(folder.id, currentDepth + 1);
  }

  return folders;
}

function renderTree(nodes: FolderNode[], prefix: string = '', isLast: boolean[] = []): string {
  let result = '';

  nodes.forEach((node, index) => {
    const isLastNode = index === nodes.length - 1;
    const connector = isLastNode ? '└── ' : '├── ';
    const childPrefix = isLastNode ? '    ' : '│   ';

    // Build the prefix based on parent levels
    let linePrefix = '';
    for (let i = 0; i < isLast.length; i++) {
      linePrefix += isLast[i] ? '    ' : '│   ';
    }

    result += `${linePrefix}${connector}📁 ${node.name} (${node.id})\n`;

    if (node.children.length > 0) {
      result += renderTree(node.children, prefix + childPrefix, [...isLast, isLastNode]);
    }
  });

  return result;
}

function countFolders(nodes: FolderNode[]): number {
  return nodes.reduce((count, node) => count + 1 + countFolders(node.children), 0);
}

try {
  // Get the root folder name if not "root"
  let rootName = 'My Drive';
  if (args.folderId !== 'root') {
    const rootFolder = await drive.files.get({
      fileId: args.folderId,
      fields: 'name',
    });
    rootName = rootFolder.data.name || 'Unknown Folder';
  }

  const tree = await buildTree(args.folderId, 1);
  const totalFolders = countFolders(tree);

  if (tree.length === 0) {
    return `📁 ${rootName}\n   (no subfolders found)`;
  }

  let result = `📁 ${rootName}\n`;
  result += renderTree(tree);
  result += `\n---\nTotal: ${totalFolders} folder${totalFolders !== 1 ? 's' : ''} (max depth: ${args.maxDepth})`;

  return result;
} catch (error: any) {
  log.error(`Error listing folders recursively: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Folder not found. Check the folder ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have access to this folder.");
  throw new UserError(`Failed to list folders: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'getFolderInfo',
description: 'Gets detailed information about a specific folder in Google Drive.',
parameters: z.object({
  folderId: z.string().describe('ID of the folder to get information about.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Getting folder info: ${args.folderId}`);

try {
  const response = await drive.files.get({
    fileId: args.folderId,
    fields: 'id,name,description,createdTime,modifiedTime,webViewLink,owners(displayName,emailAddress),lastModifyingUser(displayName),shared,parents',
  });

  const folder = response.data;

  if (folder.mimeType !== 'application/vnd.google-apps.folder') {
    throw new UserError("The specified ID does not belong to a folder.");
  }

  const createdDate = folder.createdTime ? new Date(folder.createdTime).toLocaleString() : 'Unknown';
  const modifiedDate = folder.modifiedTime ? new Date(folder.modifiedTime).toLocaleString() : 'Unknown';
  const owner = folder.owners?.[0];
  const lastModifier = folder.lastModifyingUser;

  let result = `**Folder Information:**\n\n`;
  result += `**Name:** ${folder.name}\n`;
  result += `**ID:** ${folder.id}\n`;
  result += `**Created:** ${createdDate}\n`;
  result += `**Last Modified:** ${modifiedDate}\n`;

  if (owner) {
    result += `**Owner:** ${owner.displayName} (${owner.emailAddress})\n`;
  }

  if (lastModifier) {
    result += `**Last Modified By:** ${lastModifier.displayName}\n`;
  }

  result += `**Shared:** ${folder.shared ? 'Yes' : 'No'}\n`;
  result += `**View Link:** ${folder.webViewLink}\n`;

  if (folder.description) {
    result += `**Description:** ${folder.description}\n`;
  }

  if (folder.parents && folder.parents.length > 0) {
    result += `**Parent Folder ID:** ${folder.parents[0]}\n`;
  }

  return result;
} catch (error: any) {
  log.error(`Error getting folder info: ${error.message || error}`);
  if (error.code === 404) throw new UserError(`Folder not found (ID: ${args.folderId}).`);
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have access to this folder.");
  throw new UserError(`Failed to get folder info: ${error.message || 'Unknown error'}`);
}
}
});

// --- File Operation Tools ---

server.addTool({
name: 'moveFile',
description: 'Moves a file or folder to a different location in Google Drive.',
parameters: z.object({
  fileId: z.string().describe('ID of the file or folder to move.'),
  newParentId: z.string().describe('ID of the destination folder. Use "root" for Drive root.'),
  removeFromAllParents: z.boolean().optional().default(false).describe('If true, removes from all current parents. If false, adds to new parent while keeping existing parents.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Moving file ${args.fileId} to folder ${args.newParentId}`);

try {
  // First get the current parents
  const fileInfo = await drive.files.get({
    fileId: args.fileId,
    fields: 'name,parents',
  });

  const fileName = fileInfo.data.name;
  const currentParents = fileInfo.data.parents || [];

  let updateParams: any = {
    fileId: args.fileId,
    addParents: args.newParentId,
    fields: 'id,name,parents',
  };

  if (args.removeFromAllParents && currentParents.length > 0) {
    updateParams.removeParents = currentParents.join(',');
  }

  const response = await drive.files.update(updateParams);

  const action = args.removeFromAllParents ? 'moved' : 'copied';
  return `Successfully ${action} "${fileName}" to new location.\nFile ID: ${response.data.id}`;
} catch (error: any) {
  log.error(`Error moving file: ${error.message || error}`);
  if (error.code === 404) throw new UserError("File or destination folder not found. Check the IDs.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to both source and destination.");
  throw new UserError(`Failed to move file: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'copyFile',
description: 'Creates a copy of a Google Drive file or document.',
parameters: z.object({
  fileId: z.string().describe('ID of the file to copy.'),
  newName: z.string().optional().describe('Name for the copied file. If not provided, will use "Copy of [original name]".'),
  parentFolderId: z.string().optional().describe('ID of folder where copy should be placed. If not provided, places in same location as original.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Copying file ${args.fileId} ${args.newName ? `as "${args.newName}"` : ''}`);

try {
  // Get original file info
  const originalFile = await drive.files.get({
    fileId: args.fileId,
    fields: 'name,parents',
  });

  const copyMetadata: drive_v3.Schema$File = {
    name: args.newName || `Copy of ${originalFile.data.name}`,
  };

  if (args.parentFolderId) {
    copyMetadata.parents = [args.parentFolderId];
  } else if (originalFile.data.parents) {
    copyMetadata.parents = originalFile.data.parents;
  }

  const response = await drive.files.copy({
    fileId: args.fileId,
    requestBody: copyMetadata,
    fields: 'id,name,webViewLink',
  });

  const copiedFile = response.data;
  return `Successfully created copy "${copiedFile.name}" (ID: ${copiedFile.id})\nLink: ${copiedFile.webViewLink}`;
} catch (error: any) {
  log.error(`Error copying file: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Original file or destination folder not found. Check the IDs.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have read access to the original file and write access to the destination.");
  throw new UserError(`Failed to copy file: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'renameFile',
description: 'Renames a file or folder in Google Drive.',
parameters: z.object({
  fileId: z.string().describe('ID of the file or folder to rename.'),
  newName: z.string().min(1).describe('New name for the file or folder.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Renaming file ${args.fileId} to "${args.newName}"`);

try {
  const response = await drive.files.update({
    fileId: args.fileId,
    requestBody: {
      name: args.newName,
    },
    fields: 'id,name,webViewLink',
  });

  const file = response.data;
  return `Successfully renamed to "${file.name}" (ID: ${file.id})\nLink: ${file.webViewLink}`;
} catch (error: any) {
  log.error(`Error renaming file: ${error.message || error}`);
  if (error.code === 404) throw new UserError("File not found. Check the file ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to this file.");
  throw new UserError(`Failed to rename file: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'deleteFile',
description: 'Permanently deletes a file or folder from Google Drive.',
parameters: z.object({
  fileId: z.string().describe('ID of the file or folder to delete.'),
  skipTrash: z.boolean().optional().default(false).describe('If true, permanently deletes the file. If false, moves to trash (can be restored).'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Deleting file ${args.fileId} ${args.skipTrash ? '(permanent)' : '(to trash)'}`);

try {
  // Get file info before deletion
  const fileInfo = await drive.files.get({
    fileId: args.fileId,
    fields: 'name,mimeType',
  });

  const fileName = fileInfo.data.name;
  const isFolder = fileInfo.data.mimeType === 'application/vnd.google-apps.folder';

  if (args.skipTrash) {
    await drive.files.delete({
      fileId: args.fileId,
    });
    return `Permanently deleted ${isFolder ? 'folder' : 'file'} "${fileName}".`;
  } else {
    await drive.files.update({
      fileId: args.fileId,
      requestBody: {
        trashed: true,
      },
    });
    return `Moved ${isFolder ? 'folder' : 'file'} "${fileName}" to trash. It can be restored from the trash.`;
  }
} catch (error: any) {
  log.error(`Error deleting file: ${error.message || error}`);
  if (error.code === 404) throw new UserError("File not found. Check the file ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have delete access to this file.");
  throw new UserError(`Failed to delete file: ${error.message || 'Unknown error'}`);
}
}
});

// --- Document Creation Tools ---

server.addTool({
name: 'createDocument',
description: 'Creates a new Google Document.',
parameters: z.object({
  title: z.string().min(1).describe('Title for the new document.'),
  parentFolderId: z.string().optional().describe('ID of folder where document should be created. If not provided, creates in Drive root.'),
  initialContent: z.string().optional().describe('Initial text content to add to the document.'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Creating new document "${args.title}"`);

try {
  const documentMetadata: drive_v3.Schema$File = {
    name: args.title,
    mimeType: 'application/vnd.google-apps.document',
  };

  if (args.parentFolderId) {
    documentMetadata.parents = [args.parentFolderId];
  }

  const response = await drive.files.create({
    requestBody: documentMetadata,
    fields: 'id,name,webViewLink',
  });

  const document = response.data;
  let result = `Successfully created document "${document.name}" (ID: ${document.id})\nView Link: ${document.webViewLink}`;

  // Add initial content if provided
  if (args.initialContent) {
    try {
      const docs = await getDocsClient();
      await docs.documents.batchUpdate({
        documentId: document.id!,
        requestBody: {
          requests: [{
            insertText: {
              location: { index: 1 },
              text: args.initialContent,
            },
          }],
        },
      });
      result += `\n\nInitial content added to document.`;
    } catch (contentError: any) {
      log.warn(`Document created but failed to add initial content: ${contentError.message}`);
      result += `\n\nDocument created but failed to add initial content. You can add content manually.`;
    }
  }

  return result;
} catch (error: any) {
  log.error(`Error creating document: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Parent folder not found. Check the folder ID.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to the destination folder.");
  throw new UserError(`Failed to create document: ${error.message || 'Unknown error'}`);
}
}
});

server.addTool({
name: 'createFromTemplate',
description: 'Creates a new Google Document from an existing document template.',
parameters: z.object({
  templateId: z.string().describe('ID of the template document to copy from.'),
  newTitle: z.string().min(1).describe('Title for the new document.'),
  parentFolderId: z.string().optional().describe('ID of folder where document should be created. If not provided, creates in Drive root.'),
  replacements: z.record(z.string()).optional().describe('Key-value pairs for text replacements in the template (e.g., {"{{NAME}}": "John Doe", "{{DATE}}": "2024-01-01"}).'),
}),
execute: async (args, { log }) => {
const drive = await getDriveClient();
log.info(`Creating document from template ${args.templateId} with title "${args.newTitle}"`);

try {
  // First copy the template
  const copyMetadata: drive_v3.Schema$File = {
    name: args.newTitle,
  };

  if (args.parentFolderId) {
    copyMetadata.parents = [args.parentFolderId];
  }

  const response = await drive.files.copy({
    fileId: args.templateId,
    requestBody: copyMetadata,
    fields: 'id,name,webViewLink',
  });

  const document = response.data;
  let result = `Successfully created document "${document.name}" from template (ID: ${document.id})\nView Link: ${document.webViewLink}`;

  // Apply text replacements if provided
  if (args.replacements && Object.keys(args.replacements).length > 0) {
    try {
      const docs = await getDocsClient();
      const requests: docs_v1.Schema$Request[] = [];

      // Create replace requests for each replacement
      for (const [searchText, replaceText] of Object.entries(args.replacements)) {
        requests.push({
          replaceAllText: {
            containsText: {
              text: searchText,
              matchCase: false,
            },
            replaceText: replaceText,
          },
        });
      }

      if (requests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: document.id!,
          requestBody: { requests },
        });

        const replacementCount = Object.keys(args.replacements).length;
        result += `\n\nApplied ${replacementCount} text replacement${replacementCount !== 1 ? 's' : ''} to the document.`;
      }
    } catch (replacementError: any) {
      log.warn(`Document created but failed to apply replacements: ${replacementError.message}`);
      result += `\n\nDocument created but failed to apply text replacements. You can make changes manually.`;
    }
  }

  return result;
} catch (error: any) {
  log.error(`Error creating document from template: ${error.message || error}`);
  if (error.code === 404) throw new UserError("Template document or parent folder not found. Check the IDs.");
  if (error.code === 403) throw new UserError("Permission denied. Make sure you have read access to the template and write access to the destination folder.");
  throw new UserError(`Failed to create document from template: ${error.message || 'Unknown error'}`);
}
}
});

// === GOOGLE SHEETS TOOLS ===

server.addTool({
name: 'readSpreadsheet',
description: 'Reads data from a specific range in a Google Spreadsheet.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
  range: z.string().describe('A1 notation range to read (e.g., "A1:B10" or "Sheet1!A1:B10").'),
  valueRenderOption: z.enum(['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']).optional().default('FORMATTED_VALUE')
    .describe('How values should be rendered in the output.'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Reading spreadsheet ${args.spreadsheetId}, range: ${args.range}`);

  try {
    const response = await SheetsHelpers.readRange(sheets, args.spreadsheetId, args.range);
    const values = response.values || [];

    if (values.length === 0) {
      return `Range ${args.range} is empty or does not exist.`;
    }

    // Format as a readable table
    let result = `**Spreadsheet Range:** ${args.range}\n\n`;
    values.forEach((row, index) => {
      result += `Row ${index + 1}: ${JSON.stringify(row)}\n`;
    });

    return result;
  } catch (error: any) {
    log.error(`Error reading spreadsheet ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to read spreadsheet: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'writeSpreadsheet',
description: 'Writes data to a specific range in a Google Spreadsheet. Overwrites existing data in the range.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
  range: z.string().describe('A1 notation range to write to (e.g., "A1:B2" or "Sheet1!A1:B2").'),
  values: z.array(z.array(z.any())).describe('2D array of values to write. Each inner array represents a row.'),
  valueInputOption: z.enum(['RAW', 'USER_ENTERED']).optional().default('USER_ENTERED')
    .describe('How input data should be interpreted. RAW: values are stored as-is. USER_ENTERED: values are parsed as if typed by a user.'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Writing to spreadsheet ${args.spreadsheetId}, range: ${args.range}`);

  try {
    const response = await SheetsHelpers.writeRange(
      sheets,
      args.spreadsheetId,
      args.range,
      args.values,
      args.valueInputOption
    );

    const updatedCells = response.updatedCells || 0;
    const updatedRows = response.updatedRows || 0;
    const updatedColumns = response.updatedColumns || 0;

    return `Successfully wrote ${updatedCells} cells (${updatedRows} rows, ${updatedColumns} columns) to range ${args.range}.`;
  } catch (error: any) {
    log.error(`Error writing to spreadsheet ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to write to spreadsheet: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'appendSpreadsheetRows',
description: 'Appends rows of data to the end of a sheet in a Google Spreadsheet.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
  range: z.string().describe('A1 notation range indicating where to append (e.g., "A1" or "Sheet1!A1"). Data will be appended starting from this range.'),
  values: z.array(z.array(z.any())).describe('2D array of values to append. Each inner array represents a row.'),
  valueInputOption: z.enum(['RAW', 'USER_ENTERED']).optional().default('USER_ENTERED')
    .describe('How input data should be interpreted. RAW: values are stored as-is. USER_ENTERED: values are parsed as if typed by a user.'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Appending rows to spreadsheet ${args.spreadsheetId}, starting at: ${args.range}`);

  try {
    const response = await SheetsHelpers.appendValues(
      sheets,
      args.spreadsheetId,
      args.range,
      args.values,
      args.valueInputOption
    );

    const updatedCells = response.updates?.updatedCells || 0;
    const updatedRows = response.updates?.updatedRows || 0;
    const updatedRange = response.updates?.updatedRange || args.range;

    return `Successfully appended ${updatedRows} row(s) (${updatedCells} cells) to spreadsheet. Updated range: ${updatedRange}`;
  } catch (error: any) {
    log.error(`Error appending to spreadsheet ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to append to spreadsheet: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'clearSpreadsheetRange',
description: 'Clears all values from a specific range in a Google Spreadsheet.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
  range: z.string().describe('A1 notation range to clear (e.g., "A1:B10" or "Sheet1!A1:B10").'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Clearing range ${args.range} in spreadsheet ${args.spreadsheetId}`);

  try {
    const response = await SheetsHelpers.clearRange(sheets, args.spreadsheetId, args.range);
    const clearedRange = response.clearedRange || args.range;

    return `Successfully cleared range ${clearedRange}.`;
  } catch (error: any) {
    log.error(`Error clearing range in spreadsheet ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to clear range: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'getSpreadsheetInfo',
description: 'Gets detailed information about a Google Spreadsheet including all sheets/tabs.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Getting info for spreadsheet: ${args.spreadsheetId}`);

  try {
    const metadata = await SheetsHelpers.getSpreadsheetMetadata(sheets, args.spreadsheetId);

    let result = `**Spreadsheet Information:**\n\n`;
    result += `**Title:** ${metadata.properties?.title || 'Untitled'}\n`;
    result += `**ID:** ${metadata.spreadsheetId}\n`;
    result += `**URL:** https://docs.google.com/spreadsheets/d/${metadata.spreadsheetId}\n\n`;

    const sheetList = metadata.sheets || [];
    result += `**Sheets (${sheetList.length}):**\n`;
    sheetList.forEach((sheet, index) => {
      const props = sheet.properties;
      result += `${index + 1}. **${props?.title || 'Untitled'}**\n`;
      result += `   - Sheet ID: ${props?.sheetId}\n`;
      result += `   - Grid: ${props?.gridProperties?.rowCount || 0} rows × ${props?.gridProperties?.columnCount || 0} columns\n`;
      if (props?.hidden) {
        result += `   - Status: Hidden\n`;
      }
      result += `\n`;
    });

    return result;
  } catch (error: any) {
    log.error(`Error getting spreadsheet info ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to get spreadsheet info: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'addSpreadsheetSheet',
description: 'Adds a new sheet/tab to an existing Google Spreadsheet.',
parameters: z.object({
  spreadsheetId: z.string().describe('The ID of the Google Spreadsheet (from the URL).'),
  sheetTitle: z.string().min(1).describe('Title for the new sheet/tab.'),
}),
execute: async (args, { log }) => {
  const sheets = await getSheetsClient();
  log.info(`Adding sheet "${args.sheetTitle}" to spreadsheet ${args.spreadsheetId}`);

  try {
    const response = await SheetsHelpers.addSheet(sheets, args.spreadsheetId, args.sheetTitle);
    const addedSheet = response.replies?.[0]?.addSheet?.properties;

    if (!addedSheet) {
      throw new UserError('Failed to add sheet - no sheet properties returned.');
    }

    return `Successfully added sheet "${addedSheet.title}" (Sheet ID: ${addedSheet.sheetId}) to spreadsheet.`;
  } catch (error: any) {
    log.error(`Error adding sheet to spreadsheet ${args.spreadsheetId}: ${error.message || error}`);
    if (error instanceof UserError) throw error;
    throw new UserError(`Failed to add sheet: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'createSpreadsheet',
description: 'Creates a new Google Spreadsheet.',
parameters: z.object({
  title: z.string().min(1).describe('Title for the new spreadsheet.'),
  parentFolderId: z.string().optional().describe('ID of folder where spreadsheet should be created. If not provided, creates in Drive root.'),
  initialData: z.array(z.array(z.any())).optional().describe('Optional initial data to populate in the first sheet. Each inner array represents a row.'),
}),
execute: async (args, { log }) => {
  const drive = await getDriveClient();
  const sheets = await getSheetsClient();
  log.info(`Creating new spreadsheet "${args.title}"`);

  try {
    // Create the spreadsheet file in Drive
    const spreadsheetMetadata: drive_v3.Schema$File = {
      name: args.title,
      mimeType: 'application/vnd.google-apps.spreadsheet',
    };

    if (args.parentFolderId) {
      spreadsheetMetadata.parents = [args.parentFolderId];
    }

    const driveResponse = await drive.files.create({
      requestBody: spreadsheetMetadata,
      fields: 'id,name,webViewLink',
    });

    const spreadsheetId = driveResponse.data.id;
    if (!spreadsheetId) {
      throw new UserError('Failed to create spreadsheet - no ID returned.');
    }

    let result = `Successfully created spreadsheet "${driveResponse.data.name}" (ID: ${spreadsheetId})\nView Link: ${driveResponse.data.webViewLink}`;

    // Add initial data if provided
    if (args.initialData && args.initialData.length > 0) {
      try {
        await SheetsHelpers.writeRange(
          sheets,
          spreadsheetId,
          'A1',
          args.initialData,
          'USER_ENTERED'
        );
        result += `\n\nInitial data added to the spreadsheet.`;
      } catch (contentError: any) {
        log.warn(`Spreadsheet created but failed to add initial data: ${contentError.message}`);
        result += `\n\nSpreadsheet created but failed to add initial data. You can add data manually.`;
      }
    }

    return result;
  } catch (error: any) {
    log.error(`Error creating spreadsheet: ${error.message || error}`);
    if (error.code === 404) throw new UserError("Parent folder not found. Check the folder ID.");
    if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to the destination folder.");
    throw new UserError(`Failed to create spreadsheet: ${error.message || 'Unknown error'}`);
  }
}
});

server.addTool({
name: 'listGoogleSheets',
description: 'Lists Google Spreadsheets from your Google Drive with optional filtering.',
parameters: z.object({
  maxResults: z.number().int().min(1).max(100).optional().default(20).describe('Maximum number of spreadsheets to return (1-100).'),
  query: z.string().optional().describe('Search query to filter spreadsheets by name or content.'),
  orderBy: z.enum(['name', 'modifiedTime', 'createdTime']).optional().default('modifiedTime').describe('Sort order for results.'),
}),
execute: async (args, { log }) => {
  const drive = await getDriveClient();
  log.info(`Listing Google Sheets. Query: ${args.query || 'none'}, Max: ${args.maxResults}, Order: ${args.orderBy}`);

  try {
    // Build the query string for Google Drive API
    let queryString = "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
    let usesFullText = false;
    if (args.query) {
      // Use name-only search to allow sorting, fullText search doesn't support orderBy
      queryString += ` and name contains '${args.query}'`;
    }

    const response = await drive.files.list({
      q: queryString,
      pageSize: args.maxResults,
      orderBy: args.orderBy === 'name' ? 'name' : args.orderBy,
      fields: 'files(id,name,modifiedTime,createdTime,size,webViewLink,owners(displayName,emailAddress))',
    });

    const files = response.data.files || [];

    if (files.length === 0) {
      return "No Google Spreadsheets found matching your criteria.";
    }

    let result = `Found ${files.length} Google Spreadsheet(s):\n\n`;
    files.forEach((file, index) => {
      const modifiedDate = file.modifiedTime ? new Date(file.modifiedTime).toLocaleDateString() : 'Unknown';
      const owner = file.owners?.[0]?.displayName || 'Unknown';
      result += `${index + 1}. **${file.name}**\n`;
      result += `   ID: ${file.id}\n`;
      result += `   Modified: ${modifiedDate}\n`;
      result += `   Owner: ${owner}\n`;
      result += `   Link: ${file.webViewLink}\n\n`;
    });

    return result;
  } catch (error: any) {
    log.error(`Error listing Google Sheets: ${error.message || error}`);
    if (error.code === 403) {
      // Check if this is an actual permission error or an API limitation
      const errorReason = error.response?.data?.error?.errors?.[0]?.reason;
      if (errorReason === 'forbidden' && error.message?.includes('Sorting')) {
        throw new UserError(`API limitation: ${error.message}`);
      }
      throw new UserError("Permission denied. Make sure you have granted Google Drive access to the application.");
    }
    throw new UserError(`Failed to list spreadsheets: ${error.message || 'Unknown error'}`);
  }
}
});

// === ENHANCED FORMATTING TOOLS ===

// Schema for structured content sections
const FormattedSectionSchema = z.object({
  type: z.enum(['heading1', 'heading2', 'heading3', 'heading4', 'title', 'subtitle', 'normal', 'bullet', 'numbered']).describe('The type of content section'),
  text: z.string().describe('The text content for this section'),
  bold: z.boolean().optional().describe('Apply bold to the entire section'),
  italic: z.boolean().optional().describe('Apply italic to the entire section'),
  color: z.string().optional().describe('Text color in hex format (e.g., "#FF0000" for red)'),
});

const FormattedContentSchema = z.array(FormattedSectionSchema).describe('Array of formatted content sections to insert');

server.addTool({
  name: 'createFormattedDocument',
  description: `Creates a new Google Document with properly formatted content. Instead of raw text, you provide structured sections with formatting types (headings, bullets, etc.) and the tool applies proper Google Docs formatting automatically. This prevents auto-list conversion issues and ensures clean, professional documents.

IMPORTANT: Use this instead of createDocument + insertText when you need formatted content.

Example content array:
[
  { "type": "title", "text": "My Document Title" },
  { "type": "heading1", "text": "Introduction", "color": "#1565C0" },
  { "type": "normal", "text": "This is regular paragraph text." },
  { "type": "bullet", "text": "First bullet point" },
  { "type": "bullet", "text": "Second bullet point" },
  { "type": "heading2", "text": "Section Two" },
  { "type": "normal", "text": "More content here.", "bold": true }
]`,
  parameters: z.object({
    title: z.string().min(1).describe('Title for the new document (appears in Drive).'),
    content: FormattedContentSchema.describe('Array of formatted content sections'),
    parentFolderId: z.string().optional().describe('ID of folder where document should be created. If not provided, creates in Drive root.'),
  }),
  execute: async (args, { log }) => {
    const drive = await getDriveClient();
    const docs = await getDocsClient();
    log.info(`Creating formatted document "${args.title}" with ${args.content.length} sections`);

    try {
      // Step 1: Create the document
      const documentMetadata: drive_v3.Schema$File = {
        name: args.title,
        mimeType: 'application/vnd.google-apps.document',
      };

      if (args.parentFolderId) {
        documentMetadata.parents = [args.parentFolderId];
      }

      const createResponse = await drive.files.create({
        requestBody: documentMetadata,
        fields: 'id,name,webViewLink',
      });

      const document = createResponse.data;
      const documentId = document.id!;
      log.info(`Document created: ${documentId}`);

      // Step 2: Build batch update requests for all content
      const requests: docs_v1.Schema$Request[] = [];
      let currentIndex = 1; // Google Docs starts at index 1
      const styleRequests: docs_v1.Schema$Request[] = [];

      // Track bullet list state
      let inBulletList = false;
      let bulletListId: string | null = null;
      let inNumberedList = false;
      let numberedListId: string | null = null;

      // Process each content section
      for (let i = 0; i < args.content.length; i++) {
        const section = args.content[i];
        const text = section.text + '\n'; // Add newline after each section
        const textLength = text.length;
        const startIndex = currentIndex;
        const endIndex = currentIndex + textLength;

        // Insert the text
        requests.push({
          insertText: {
            location: { index: currentIndex },
            text: text,
          },
        });

        // Determine paragraph style based on type
        let namedStyleType: string | null = null;
        let isBullet = false;
        let isNumbered = false;

        switch (section.type) {
          case 'title':
            namedStyleType = 'TITLE';
            break;
          case 'subtitle':
            namedStyleType = 'SUBTITLE';
            break;
          case 'heading1':
            namedStyleType = 'HEADING_1';
            break;
          case 'heading2':
            namedStyleType = 'HEADING_2';
            break;
          case 'heading3':
            namedStyleType = 'HEADING_3';
            break;
          case 'heading4':
            namedStyleType = 'HEADING_4';
            break;
          case 'bullet':
            isBullet = true;
            break;
          case 'numbered':
            isNumbered = true;
            break;
          case 'normal':
          default:
            namedStyleType = 'NORMAL_TEXT';
            break;
        }

        // Apply paragraph style if it's a named style
        if (namedStyleType) {
          styleRequests.push({
            updateParagraphStyle: {
              range: { startIndex, endIndex: endIndex - 1 }, // Exclude the newline for paragraph style
              paragraphStyle: {
                namedStyleType: namedStyleType,
              },
              fields: 'namedStyleType',
            },
          });
        }

        // Handle bullet lists
        if (isBullet) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIndex - 1 },
              bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
            },
          });
        }

        // Handle numbered lists
        if (isNumbered) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIndex - 1 },
              bulletPreset: 'NUMBERED_DECIMAL_NESTED',
            },
          });
        }

        // Apply text formatting (bold, italic, color)
        const textStyleFields: string[] = [];
        const textStyle: docs_v1.Schema$TextStyle = {};

        if (section.bold) {
          textStyle.bold = true;
          textStyleFields.push('bold');
        }
        if (section.italic) {
          textStyle.italic = true;
          textStyleFields.push('italic');
        }
        if (section.color) {
          // Parse hex color
          const hex = section.color.replace('#', '');
          const r = parseInt(hex.substring(0, 2), 16) / 255;
          const g = parseInt(hex.substring(2, 4), 16) / 255;
          const b = parseInt(hex.substring(4, 6), 16) / 255;
          textStyle.foregroundColor = {
            color: { rgbColor: { red: r, green: g, blue: b } },
          };
          textStyleFields.push('foregroundColor');
        }

        if (textStyleFields.length > 0) {
          styleRequests.push({
            updateTextStyle: {
              range: { startIndex, endIndex: endIndex - 1 }, // Exclude newline
              textStyle: textStyle,
              fields: textStyleFields.join(','),
            },
          });
        }

        currentIndex = endIndex;
      }

      // Step 3: Execute batch update - first insert all text, then apply styles
      if (requests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: documentId,
          requestBody: { requests },
        });
        log.info(`Inserted ${requests.length} text sections`);
      }

      // Apply styles in a separate batch (after text is inserted)
      if (styleRequests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: documentId,
          requestBody: { requests: styleRequests },
        });
        log.info(`Applied ${styleRequests.length} style updates`);
      }

      return `Successfully created formatted document "${document.name}" (ID: ${document.id})\nView Link: ${document.webViewLink}\n\nAdded ${args.content.length} formatted sections.`;

    } catch (error: any) {
      log.error(`Error creating formatted document: ${error.message || error}`);
      if (error.code === 404) throw new UserError("Parent folder not found. Check the folder ID.");
      if (error.code === 403) throw new UserError("Permission denied. Make sure you have write access to the destination folder.");
      throw new UserError(`Failed to create formatted document: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'insertFormattedContent',
  description: `Inserts properly formatted content into an existing Google Document at a specified index. Uses structured sections with formatting types (headings, bullets, etc.) instead of raw text, ensuring clean formatting without auto-list conversion issues.

IMPORTANT: Use this instead of insertText when you need formatted content.

Example content array:
[
  { "type": "heading1", "text": "New Section", "color": "#1565C0" },
  { "type": "normal", "text": "This is paragraph text." },
  { "type": "bullet", "text": "Bullet point one" },
  { "type": "bullet", "text": "Bullet point two" }
]`,
  parameters: DocumentIdParameter.extend({
    index: z.number().int().min(1).describe('The index (1-based) where the content should be inserted.'),
    content: FormattedContentSchema.describe('Array of formatted content sections to insert'),
  }),
  execute: async (args, { log }) => {
    const docs = await getDocsClient();
    log.info(`Inserting ${args.content.length} formatted sections into document ${args.documentId} at index ${args.index}`);

    try {
      const requests: docs_v1.Schema$Request[] = [];
      let currentIndex = args.index;
      const styleRequests: docs_v1.Schema$Request[] = [];

      // Process each content section
      for (let i = 0; i < args.content.length; i++) {
        const section = args.content[i];
        const text = section.text + '\n';
        const textLength = text.length;
        const startIndex = currentIndex;
        const endIndex = currentIndex + textLength;

        // Insert the text
        requests.push({
          insertText: {
            location: { index: currentIndex },
            text: text,
          },
        });

        // Determine paragraph style based on type
        let namedStyleType: string | null = null;
        let isBullet = false;
        let isNumbered = false;

        switch (section.type) {
          case 'title':
            namedStyleType = 'TITLE';
            break;
          case 'subtitle':
            namedStyleType = 'SUBTITLE';
            break;
          case 'heading1':
            namedStyleType = 'HEADING_1';
            break;
          case 'heading2':
            namedStyleType = 'HEADING_2';
            break;
          case 'heading3':
            namedStyleType = 'HEADING_3';
            break;
          case 'heading4':
            namedStyleType = 'HEADING_4';
            break;
          case 'bullet':
            isBullet = true;
            break;
          case 'numbered':
            isNumbered = true;
            break;
          case 'normal':
          default:
            namedStyleType = 'NORMAL_TEXT';
            break;
        }

        if (namedStyleType) {
          styleRequests.push({
            updateParagraphStyle: {
              range: { startIndex, endIndex: endIndex - 1 },
              paragraphStyle: {
                namedStyleType: namedStyleType,
              },
              fields: 'namedStyleType',
            },
          });
        }

        if (isBullet) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIndex - 1 },
              bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
            },
          });
        }

        if (isNumbered) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIndex - 1 },
              bulletPreset: 'NUMBERED_DECIMAL_NESTED',
            },
          });
        }

        // Apply text formatting
        const textStyleFields: string[] = [];
        const textStyle: docs_v1.Schema$TextStyle = {};

        if (section.bold) {
          textStyle.bold = true;
          textStyleFields.push('bold');
        }
        if (section.italic) {
          textStyle.italic = true;
          textStyleFields.push('italic');
        }
        if (section.color) {
          const hex = section.color.replace('#', '');
          const r = parseInt(hex.substring(0, 2), 16) / 255;
          const g = parseInt(hex.substring(2, 4), 16) / 255;
          const b = parseInt(hex.substring(4, 6), 16) / 255;
          textStyle.foregroundColor = {
            color: { rgbColor: { red: r, green: g, blue: b } },
          };
          textStyleFields.push('foregroundColor');
        }

        if (textStyleFields.length > 0) {
          styleRequests.push({
            updateTextStyle: {
              range: { startIndex, endIndex: endIndex - 1 },
              textStyle: textStyle,
              fields: textStyleFields.join(','),
            },
          });
        }

        currentIndex = endIndex;
      }

      // Execute batch update - first insert text, then apply styles
      if (requests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: args.documentId,
          requestBody: { requests },
        });
        log.info(`Inserted ${requests.length} text sections`);
      }

      if (styleRequests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: args.documentId,
          requestBody: { requests: styleRequests },
        });
        log.info(`Applied ${styleRequests.length} style updates`);
      }

      return `Successfully inserted ${args.content.length} formatted sections at index ${args.index}.`;

    } catch (error: any) {
      log.error(`Error inserting formatted content: ${error.message || error}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to insert formatted content: ${error.message || 'Unknown error'}`);
    }
  }
});

server.addTool({
  name: 'replaceDocumentContent',
  description: `Completely replaces the content of an existing Google Document with properly formatted content. Useful for regenerating a document with clean formatting.

IMPORTANT: This will DELETE all existing content and replace it with the new formatted content.`,
  parameters: DocumentIdParameter.extend({
    content: FormattedContentSchema.describe('Array of formatted content sections to replace the document with'),
  }),
  execute: async (args, { log }) => {
    const docs = await getDocsClient();
    log.info(`Replacing content in document ${args.documentId} with ${args.content.length} formatted sections`);

    try {
      // Step 1: Get current document to find its length
      const docResponse = await docs.documents.get({
        documentId: args.documentId,
        fields: 'body(content(endIndex))',
      });

      // Find the end index of the document
      let endIndex = 1;
      if (docResponse.data.body?.content) {
        const lastElement = docResponse.data.body.content[docResponse.data.body.content.length - 1];
        if (lastElement?.endIndex) {
          endIndex = lastElement.endIndex;
        }
      }

      // Step 2: Delete all existing content (if any)
      if (endIndex > 2) {
        await docs.documents.batchUpdate({
          documentId: args.documentId,
          requestBody: {
            requests: [{
              deleteContentRange: {
                range: { startIndex: 1, endIndex: endIndex - 1 },
              },
            }],
          },
        });
        log.info(`Deleted existing content (indices 1-${endIndex - 1})`);
      }

      // Step 3: Insert new formatted content
      const requests: docs_v1.Schema$Request[] = [];
      let currentIndex = 1;
      const styleRequests: docs_v1.Schema$Request[] = [];

      for (let i = 0; i < args.content.length; i++) {
        const section = args.content[i];
        const text = section.text + '\n';
        const textLength = text.length;
        const startIndex = currentIndex;
        const endIdx = currentIndex + textLength;

        requests.push({
          insertText: {
            location: { index: currentIndex },
            text: text,
          },
        });

        let namedStyleType: string | null = null;
        let isBullet = false;
        let isNumbered = false;

        switch (section.type) {
          case 'title': namedStyleType = 'TITLE'; break;
          case 'subtitle': namedStyleType = 'SUBTITLE'; break;
          case 'heading1': namedStyleType = 'HEADING_1'; break;
          case 'heading2': namedStyleType = 'HEADING_2'; break;
          case 'heading3': namedStyleType = 'HEADING_3'; break;
          case 'heading4': namedStyleType = 'HEADING_4'; break;
          case 'bullet': isBullet = true; break;
          case 'numbered': isNumbered = true; break;
          default: namedStyleType = 'NORMAL_TEXT'; break;
        }

        if (namedStyleType) {
          styleRequests.push({
            updateParagraphStyle: {
              range: { startIndex, endIndex: endIdx - 1 },
              paragraphStyle: { namedStyleType },
              fields: 'namedStyleType',
            },
          });
        }

        if (isBullet) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIdx - 1 },
              bulletPreset: 'BULLET_DISC_CIRCLE_SQUARE',
            },
          });
        }

        if (isNumbered) {
          styleRequests.push({
            createParagraphBullets: {
              range: { startIndex, endIndex: endIdx - 1 },
              bulletPreset: 'NUMBERED_DECIMAL_NESTED',
            },
          });
        }

        const textStyleFields: string[] = [];
        const textStyle: docs_v1.Schema$TextStyle = {};

        if (section.bold) { textStyle.bold = true; textStyleFields.push('bold'); }
        if (section.italic) { textStyle.italic = true; textStyleFields.push('italic'); }
        if (section.color) {
          const hex = section.color.replace('#', '');
          const r = parseInt(hex.substring(0, 2), 16) / 255;
          const g = parseInt(hex.substring(2, 4), 16) / 255;
          const b = parseInt(hex.substring(4, 6), 16) / 255;
          textStyle.foregroundColor = { color: { rgbColor: { red: r, green: g, blue: b } } };
          textStyleFields.push('foregroundColor');
        }

        if (textStyleFields.length > 0) {
          styleRequests.push({
            updateTextStyle: {
              range: { startIndex, endIndex: endIdx - 1 },
              textStyle,
              fields: textStyleFields.join(','),
            },
          });
        }

        currentIndex = endIdx;
      }

      if (requests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: args.documentId,
          requestBody: { requests },
        });
      }

      if (styleRequests.length > 0) {
        await docs.documents.batchUpdate({
          documentId: args.documentId,
          requestBody: { requests: styleRequests },
        });
      }

      return `Successfully replaced document content with ${args.content.length} formatted sections.`;

    } catch (error: any) {
      log.error(`Error replacing document content: ${error.message || error}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to replace document content: ${error.message || 'Unknown error'}`);
    }
  }
});

// === GOOGLE SLIDES TOOLS ===

// --- Read Tools ---

server.addTool({
  name: 'getPresentation',
  description: 'Gets metadata and content of a Google Slides presentation including title, slide count, dimensions, and optionally full content.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the Google Slides presentation (from URL: docs.google.com/presentation/d/PRESENTATION_ID/edit).'),
    includeContent: z.boolean().optional().default(false).describe('If true, includes detailed content of all slides.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Getting presentation: ${args.presentationId}`);

    try {
      const presentation = await SlidesHelpers.getPresentation(slides, args.presentationId);

      const result: any = {
        presentationId: presentation.presentationId,
        title: presentation.title || 'Untitled Presentation',
        slideCount: presentation.slides?.length || 0,
        locale: presentation.locale,
        pageSize: presentation.pageSize ? {
          width: presentation.pageSize.width?.magnitude ? SlidesHelpers.pointsFromEmu(presentation.pageSize.width.magnitude) : null,
          height: presentation.pageSize.height?.magnitude ? SlidesHelpers.pointsFromEmu(presentation.pageSize.height.magnitude) : null,
          unit: 'points',
        } : null,
        slides: presentation.slides?.map((slide, index) => ({
          index,
          pageObjectId: slide.objectId,
          layoutObjectId: slide.slideProperties?.layoutObjectId,
        })) || [],
      };

      if (args.includeContent && presentation.slides) {
        result.slidesContent = presentation.slides.map((slide) => ({
          pageObjectId: slide.objectId,
          elements: slide.pageElements?.map((el) => ({
            objectId: el.objectId,
            type: el.shape ? 'shape' : el.image ? 'image' : el.table ? 'table' : el.line ? 'line' : el.video ? 'video' : 'other',
            shapeType: el.shape?.shapeType,
            hasText: !!el.shape?.text?.textElements?.length,
          })) || [],
        }));
      }

      return JSON.stringify(result, null, 2);
    } catch (error: any) {
      log.error(`Error getting presentation: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get presentation: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'listSlides',
  description: 'Lists all slides in a Google Slides presentation with their IDs and basic info.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the Google Slides presentation.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Listing slides in presentation: ${args.presentationId}`);

    try {
      const presentation = await SlidesHelpers.getPresentation(slides, args.presentationId);

      const slideList = presentation.slides?.map((slide, index) => {
        // Try to find a title placeholder text
        let title = null;
        if (slide.pageElements) {
          for (const el of slide.pageElements) {
            if (el.shape?.placeholder?.type === 'TITLE' || el.shape?.placeholder?.type === 'CENTERED_TITLE') {
              const textElements = el.shape?.text?.textElements;
              if (textElements) {
                title = textElements
                  .filter((te) => te.textRun?.content)
                  .map((te) => te.textRun?.content?.trim())
                  .join('')
                  .trim() || null;
              }
              break;
            }
          }
        }

        return {
          index,
          pageObjectId: slide.objectId,
          title: title,
          layoutObjectId: slide.slideProperties?.layoutObjectId,
          elementCount: slide.pageElements?.length || 0,
        };
      }) || [];

      return JSON.stringify({
        presentationId: args.presentationId,
        title: presentation.title || 'Untitled Presentation',
        slideCount: slideList.length,
        slides: slideList,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error listing slides: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to list slides: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'getSlide',
  description: 'Gets detailed content of one or more slides including all elements. Can fetch a single slide by ID, or multiple slides by index range.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the Google Slides presentation.'),
    pageObjectId: z.string().optional().describe('The object ID of a single slide to retrieve. Use this OR startIndex/endIndex, not both.'),
    startIndex: z.number().int().min(0).optional().describe('0-based start index for fetching a range of slides. Use with endIndex.'),
    endIndex: z.number().int().min(0).optional().describe('0-based end index (exclusive) for fetching a range of slides. Use with startIndex.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();

    // Extract text from shape elements
    const extractText = (shape: any): string | null => {
      if (!shape?.text?.textElements) return null;
      return shape.text.textElements
        .filter((te: any) => te.textRun?.content)
        .map((te: any) => te.textRun?.content)
        .join('')
        .trim() || null;
    };

    // Process a single slide into response format
    const processSlide = (slide: any, index?: number) => {
      const elements = slide.pageElements?.map((el: any) => {
        const element: any = {
          objectId: el.objectId,
          type: el.shape ? 'shape' : el.image ? 'image' : el.table ? 'table' : el.line ? 'line' : el.video ? 'video' : 'other',
        };

        if (el.shape) {
          element.shapeType = el.shape.shapeType;
          element.placeholderType = el.shape.placeholder?.type;
          element.text = extractText(el.shape);
        }

        if (el.image) {
          element.sourceUrl = el.image.sourceUrl;
          element.contentUrl = el.image.contentUrl;
        }

        if (el.table) {
          element.rows = el.table.rows;
          element.columns = el.table.columns;
          element.tableContent = el.table.tableRows?.map((row: any) =>
            row.tableCells?.map((cell: any) => extractText(cell) || '') ?? []
          ) ?? [];
        }

        // Include size/position if available
        if (el.size) {
          element.size = {
            width: el.size.width?.magnitude ? SlidesHelpers.pointsFromEmu(el.size.width.magnitude) : null,
            height: el.size.height?.magnitude ? SlidesHelpers.pointsFromEmu(el.size.height.magnitude) : null,
            unit: 'points',
          };
        }

        if (el.transform) {
          element.position = {
            x: el.transform.translateX ? SlidesHelpers.pointsFromEmu(el.transform.translateX) : 0,
            y: el.transform.translateY ? SlidesHelpers.pointsFromEmu(el.transform.translateY) : 0,
            unit: 'points',
          };
        }

        return element;
      }) || [];

      return {
        index: index,
        pageObjectId: slide.objectId,
        layoutObjectId: slide.slideProperties?.layoutObjectId,
        masterObjectId: slide.slideProperties?.masterObjectId,
        elementCount: elements.length,
        elements: elements,
      };
    };

    try {
      // Mode 1: Fetch single slide by pageObjectId
      if (args.pageObjectId) {
        log.info(`Getting slide ${args.pageObjectId} from presentation: ${args.presentationId}`);
        const slide = await SlidesHelpers.getSlide(slides, args.presentationId, args.pageObjectId);
        const result = processSlide(slide);
        delete result.index; // Remove index for single slide response (backward compatibility)
        return JSON.stringify(result, null, 2);
      }

      // Mode 2: Fetch slides by index range
      if (args.startIndex !== undefined && args.endIndex !== undefined) {
        if (args.startIndex >= args.endIndex) {
          throw new UserError('startIndex must be less than endIndex.');
        }

        log.info(`Getting slides ${args.startIndex}-${args.endIndex - 1} from presentation: ${args.presentationId}`);
        const presentation = await SlidesHelpers.getPresentation(slides, args.presentationId);
        const allSlides = presentation.slides || [];
        const totalSlides = allSlides.length;

        if (args.startIndex >= totalSlides) {
          throw new UserError(`startIndex (${args.startIndex}) is out of range. Presentation has ${totalSlides} slides.`);
        }

        const effectiveEndIndex = Math.min(args.endIndex, totalSlides);
        const selectedSlides = allSlides.slice(args.startIndex, effectiveEndIndex);

        const results = selectedSlides.map((slide, i) => processSlide(slide, args.startIndex! + i));

        return JSON.stringify({
          presentationId: args.presentationId,
          requestedRange: { startIndex: args.startIndex, endIndex: args.endIndex },
          actualRange: { startIndex: args.startIndex, endIndex: effectiveEndIndex },
          slideCount: results.length,
          totalSlides: totalSlides,
          slides: results,
        }, null, 2);
      }

      // No valid parameters provided
      throw new UserError('Either pageObjectId OR both startIndex and endIndex must be provided.');

    } catch (error: any) {
      log.error(`Error getting slide(s): ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get slide(s): ${error.message}`);
    }
  }
});

server.addTool({
  name: 'mapSlide',
  description: 'Analyzes a slide to map dimensions, element positions, and available space. Use this BEFORE adding elements to understand the layout.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the Google Slides presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide to analyze.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Mapping slide ${args.pageObjectId} in presentation: ${args.presentationId}`);

    try {
      // Get presentation for page dimensions
      const presentation = await SlidesHelpers.getPresentation(slides, args.presentationId);
      const pageWidth = presentation.pageSize?.width?.magnitude
        ? SlidesHelpers.pointsFromEmu(presentation.pageSize.width.magnitude)
        : 720;
      const pageHeight = presentation.pageSize?.height?.magnitude
        ? SlidesHelpers.pointsFromEmu(presentation.pageSize.height.magnitude)
        : 540;

      // Get slide content
      const slide = await SlidesHelpers.getSlide(slides, args.presentationId, args.pageObjectId);

      // Map element bounding boxes
      const elements: any[] = [];
      let minX = pageWidth, minY = pageHeight, maxX = 0, maxY = 0;

      for (const el of slide.pageElements || []) {
        const width = el.size?.width?.magnitude ? SlidesHelpers.pointsFromEmu(el.size.width.magnitude) : 0;
        const height = el.size?.height?.magnitude ? SlidesHelpers.pointsFromEmu(el.size.height.magnitude) : 0;
        const x = el.transform?.translateX ? SlidesHelpers.pointsFromEmu(el.transform.translateX) : 0;
        const y = el.transform?.translateY ? SlidesHelpers.pointsFromEmu(el.transform.translateY) : 0;

        // Extract element type and text
        let type = 'unknown';
        let text: string | null = null;
        if (el.shape) {
          type = el.shape.shapeType || 'shape';
          if (el.shape.text?.textElements) {
            text = el.shape.text.textElements
              .filter((te: any) => te.textRun?.content)
              .map((te: any) => te.textRun?.content)
              .join('')
              .trim() || null;
          }
        } else if (el.image) type = 'image';
        else if (el.table) {
          type = 'table';
          const cellTexts: string[] = [];
          for (const row of el.table.tableRows || []) {
            for (const cell of row.tableCells || []) {
              const cellText = cell.text?.textElements
                ?.filter((te: any) => te.textRun?.content)
                .map((te: any) => te.textRun?.content)
                .join('')
                .trim();
              if (cellText) cellTexts.push(cellText);
            }
          }
          text = cellTexts.join(' | ') || null;
        }
        else if (el.line) type = 'line';
        else if (el.video) type = 'video';

        const element = {
          objectId: el.objectId,
          type,
          text: text ? (text.length > 50 ? text.substring(0, 50) + '...' : text) : null,
          bounds: {
            x: Math.round(x),
            y: Math.round(y),
            width: Math.round(width),
            height: Math.round(height),
            right: Math.round(x + width),
            bottom: Math.round(y + height),
          }
        };
        elements.push(element);

        // Track content bounds
        if (x < minX) minX = x;
        if (y < minY) minY = y;
        if (x + width > maxX) maxX = x + width;
        if (y + height > maxY) maxY = y + height;
      }

      // Calculate available margins
      const margins = {
        top: Math.round(minY),
        bottom: Math.round(pageHeight - maxY),
        left: Math.round(minX),
        right: Math.round(pageWidth - maxX),
      };

      // Suggest safe insertion zones
      const safeZones = {
        topStrip: { x: 0, y: 0, width: pageWidth, height: Math.max(margins.top - 10, 40), note: 'Top margin area' },
        bottomStrip: { x: 0, y: maxY + 10, width: pageWidth, height: Math.max(margins.bottom - 10, 40), note: 'Bottom margin area' },
        topRight: { x: pageWidth - 100, y: 10, width: 90, height: 90, note: 'Top-right corner' },
        bottomRight: { x: pageWidth - 100, y: pageHeight - 100, width: 90, height: 90, note: 'Bottom-right corner' },
        bottomLeft: { x: 10, y: pageHeight - 100, width: 90, height: 90, note: 'Bottom-left corner' },
      };

      return JSON.stringify({
        page: {
          width: Math.round(pageWidth),
          height: Math.round(pageHeight),
          unit: 'points',
        },
        contentBounds: {
          minX: Math.round(minX),
          minY: Math.round(minY),
          maxX: Math.round(maxX),
          maxY: Math.round(maxY),
        },
        margins,
        elementCount: elements.length,
        elements,
        safeZones,
        tip: 'Use safeZones for adding new elements without overlap. All values in points.',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error mapping slide: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to map slide: ${error.message}`);
    }
  }
});

// --- Create Tools ---

server.addTool({
  name: 'createPresentation',
  description: 'Creates a new Google Slides presentation. Returns the new presentation ID and URL.',
  parameters: z.object({
    title: z.string().describe('The title for the new presentation.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Creating new presentation: ${args.title}`);

    try {
      const response = await slides.presentations.create({
        requestBody: {
          title: args.title,
        },
      });

      const presentation = response.data;
      const presentationId = presentation.presentationId;
      const url = `https://docs.google.com/presentation/d/${presentationId}/edit`;

      return JSON.stringify({
        presentationId: presentationId,
        title: presentation.title,
        url: url,
        slideCount: presentation.slides?.length || 0,
        firstSlideId: presentation.slides?.[0]?.objectId || null,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error creating presentation: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to create presentation: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'addSlide',
  description: 'Adds a new slide to a presentation. Returns the new slide ID.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    insertionIndex: z.number().int().min(0).optional().describe('Optional 0-based index where to insert the slide. Omit to add at the end.'),
    predefinedLayout: z.enum([
      'BLANK', 'CAPTION_ONLY', 'TITLE', 'TITLE_AND_BODY', 'TITLE_AND_TWO_COLUMNS',
      'TITLE_ONLY', 'SECTION_HEADER', 'SECTION_TITLE_AND_DESCRIPTION', 'ONE_COLUMN_TEXT',
      'MAIN_POINT', 'BIG_NUMBER'
    ]).optional().describe('Optional predefined layout type. Defaults to BLANK if not specified.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Adding slide to presentation: ${args.presentationId}`);

    try {
      const newSlideId = SlidesHelpers.generateObjectId('slide');
      const request = SlidesHelpers.buildCreateSlideRequest(
        args.insertionIndex,
        args.predefinedLayout,
        newSlideId
      );

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        pageObjectId: newSlideId,
        insertionIndex: args.insertionIndex ?? 'end',
        layout: args.predefinedLayout || 'BLANK',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error adding slide: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to add slide: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'duplicateSlide',
  description: 'Duplicates an existing slide in the presentation.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide to duplicate.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Duplicating slide ${args.pageObjectId} in presentation: ${args.presentationId}`);

    try {
      const newSlideId = SlidesHelpers.generateObjectId('slide');
      const request = SlidesHelpers.buildDuplicateObjectRequest(args.pageObjectId, {
        [args.pageObjectId]: newSlideId,
      });

      const response = await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        originalSlideId: args.pageObjectId,
        newSlideId: newSlideId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error duplicating slide: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to duplicate slide: ${error.message}`);
    }
  }
});

// --- Element Tools ---

server.addTool({
  name: 'addTextBox',
  description: 'Adds a text box to a slide at the specified position with optional initial text and styling.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide to add the text box to.'),
    x: z.number().min(0).describe('X position (left edge) in points from top-left corner.'),
    y: z.number().min(0).describe('Y position (top edge) in points from top-left corner.'),
    width: z.number().positive().describe('Width of the text box in points.'),
    height: z.number().positive().describe('Height of the text box in points.'),
    text: z.string().optional().describe('Optional initial text content for the text box.'),
    fontSize: z.number().positive().optional().describe('Font size in points (e.g., 12, 14, 18). If not specified, uses Google Slides default.'),
    fontFamily: z.string().optional().describe('Font family name (e.g., "Arial", "Times New Roman", "Roboto").'),
    bold: z.boolean().optional().describe('Whether the text should be bold.'),
    italic: z.boolean().optional().describe('Whether the text should be italic.'),
    textColor: z.string().regex(/^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/).optional().describe('Text color in hex format (e.g., "#000000" for black, "#FF0000" for red).'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Adding text box to slide ${args.pageObjectId}${args.fontSize ? ` with ${args.fontSize}pt font` : ''}`);

    try {
      const textBoxId = SlidesHelpers.generateObjectId('textbox');
      const requests: slides_v1.Schema$Request[] = [];

      // Create the text box shape
      requests.push(SlidesHelpers.buildCreateShapeRequest(
        args.pageObjectId,
        'TEXT_BOX',
        args.x,
        args.y,
        args.width,
        args.height,
        textBoxId
      ));

      // Insert text if provided
      if (args.text) {
        requests.push(SlidesHelpers.buildInsertTextRequest(textBoxId, args.text, 0));

        // Apply text styling if any style options are provided
        const hasStyleOptions = args.fontSize !== undefined ||
                                args.fontFamily !== undefined ||
                                args.bold !== undefined ||
                                args.italic !== undefined ||
                                args.textColor !== undefined;

        if (hasStyleOptions) {
          const styleRequest = SlidesHelpers.buildUpdateTextStyleRequest(
            textBoxId,
            {
              fontSize: args.fontSize,
              fontFamily: args.fontFamily,
              bold: args.bold,
              italic: args.italic,
              foregroundColor: args.textColor,
            },
            'ALL'
          );
          if (styleRequest) {
            requests.push(styleRequest);
          }
        }
      }

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, requests);

      return JSON.stringify({
        success: true,
        objectId: textBoxId,
        position: { x: args.x, y: args.y, unit: 'points' },
        size: { width: args.width, height: args.height, unit: 'points' },
        hasText: !!args.text,
        style: {
          fontSize: args.fontSize,
          fontFamily: args.fontFamily,
          bold: args.bold,
          italic: args.italic,
          textColor: args.textColor,
        },
      }, null, 2);
    } catch (error: any) {
      log.error(`Error adding text box: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to add text box: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'addShape',
  description: 'Adds a shape (rectangle, ellipse, etc.) to a slide.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide.'),
    shapeType: z.enum([
      'RECTANGLE', 'ROUND_RECTANGLE', 'ELLIPSE', 'TRIANGLE', 'RIGHT_TRIANGLE',
      'PARALLELOGRAM', 'TRAPEZOID', 'PENTAGON', 'HEXAGON', 'HEPTAGON', 'OCTAGON',
      'STAR_4', 'STAR_5', 'STAR_6', 'STAR_8', 'STAR_10', 'STAR_12',
      'ARROW_EAST', 'ARROW_NORTH', 'ARROW_NORTH_EAST', 'CLOUD', 'HEART', 'PLUS'
    ]).describe('The type of shape to create.'),
    x: z.number().min(0).describe('X position in points.'),
    y: z.number().min(0).describe('Y position in points.'),
    width: z.number().positive().describe('Width in points.'),
    height: z.number().positive().describe('Height in points.'),
    fillColor: z.string().regex(/^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/).optional().describe('Optional fill color in hex format (e.g., "#FF0000").'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Adding ${args.shapeType} shape to slide ${args.pageObjectId}`);

    try {
      const shapeId = SlidesHelpers.generateObjectId('shape');
      const requests: slides_v1.Schema$Request[] = [];

      // Create the shape
      requests.push(SlidesHelpers.buildCreateShapeRequest(
        args.pageObjectId,
        args.shapeType,
        args.x,
        args.y,
        args.width,
        args.height,
        shapeId
      ));

      // Apply fill color if provided
      if (args.fillColor) {
        const fillRequest = SlidesHelpers.buildUpdateShapePropertiesRequest(shapeId, args.fillColor);
        if (fillRequest) {
          requests.push(fillRequest);
        }
      }

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, requests);

      return JSON.stringify({
        success: true,
        objectId: shapeId,
        shapeType: args.shapeType,
        position: { x: args.x, y: args.y, unit: 'points' },
        size: { width: args.width, height: args.height, unit: 'points' },
        fillColor: args.fillColor || null,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error adding shape: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to add shape: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'addImage',
  description: 'Adds an image to a slide from a publicly accessible URL.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide.'),
    imageUrl: z.string().url().describe('Publicly accessible URL of the image to insert.'),
    x: z.number().min(0).describe('X position in points.'),
    y: z.number().min(0).describe('Y position in points.'),
    width: z.number().positive().describe('Width in points.'),
    height: z.number().positive().describe('Height in points.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Adding image to slide ${args.pageObjectId}`);

    try {
      const imageId = SlidesHelpers.generateObjectId('image');
      const request = SlidesHelpers.buildCreateImageRequest(
        args.pageObjectId,
        args.imageUrl,
        args.x,
        args.y,
        args.width,
        args.height,
        imageId
      );

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        objectId: imageId,
        sourceUrl: args.imageUrl,
        position: { x: args.x, y: args.y, unit: 'points' },
        size: { width: args.width, height: args.height, unit: 'points' },
      }, null, 2);
    } catch (error: any) {
      log.error(`Error adding image: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to add image: ${error.message}. Ensure the URL is publicly accessible.`);
    }
  }
});

server.addTool({
  name: 'addTable',
  description: 'Adds a table to a slide with the specified number of rows and columns.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide.'),
    rows: z.number().int().min(1).describe('Number of rows (minimum 1).'),
    columns: z.number().int().min(1).describe('Number of columns (minimum 1).'),
    x: z.number().min(0).describe('X position in points.'),
    y: z.number().min(0).describe('Y position in points.'),
    width: z.number().positive().describe('Width in points.'),
    height: z.number().positive().describe('Height in points.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Adding ${args.rows}x${args.columns} table to slide ${args.pageObjectId}`);

    try {
      const tableId = SlidesHelpers.generateObjectId('table');
      const request = SlidesHelpers.buildCreateTableRequest(
        args.pageObjectId,
        args.rows,
        args.columns,
        args.x,
        args.y,
        args.width,
        args.height,
        tableId
      );

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        objectId: tableId,
        dimensions: { rows: args.rows, columns: args.columns },
        position: { x: args.x, y: args.y, unit: 'points' },
        size: { width: args.width, height: args.height, unit: 'points' },
      }, null, 2);
    } catch (error: any) {
      log.error(`Error adding table: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to add table: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'editSlideTableCell',
  description: 'Edits the text content and optional styling of a specific cell in a Slides table. Replaces all existing cell text with the new text.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    tableObjectId: z.string().describe('The object ID of the table element (from getSlide or addTable).'),
    rowIndex: z.number().int().min(0).describe('0-based row index of the cell to edit.'),
    columnIndex: z.number().int().min(0).describe('0-based column index of the cell to edit.'),
    text: z.string().describe('The new text content for the cell.'),
    fontSize: z.number().positive().optional().describe('Font size in points (e.g., 12, 14, 18).'),
    fontFamily: z.string().optional().describe('Font family name (e.g., "Arial", "Times New Roman", "Roboto").'),
    bold: z.boolean().optional().describe('Whether the text should be bold.'),
    italic: z.boolean().optional().describe('Whether the text should be italic.'),
    textColor: z.string().regex(/^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/).optional().describe('Text color in hex format (e.g., "#000000" for black, "#FF0000" for red).'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Editing table cell [${args.rowIndex}, ${args.columnIndex}] in table ${args.tableObjectId}`);

    try {
      const cellLocation = { rowIndex: args.rowIndex, columnIndex: args.columnIndex };
      const requests: slides_v1.Schema$Request[] = [];

      // Delete existing cell text
      requests.push({
        deleteText: {
          objectId: args.tableObjectId,
          cellLocation,
          textRange: { type: 'ALL' },
        },
      });

      // Insert new text if non-empty
      if (args.text.length > 0) {
        requests.push({
          insertText: {
            objectId: args.tableObjectId,
            cellLocation,
            text: args.text,
            insertionIndex: 0,
          },
        });
      }

      // Apply text styling if any style options are provided
      const hasStyleOptions = args.fontSize !== undefined ||
                              args.fontFamily !== undefined ||
                              args.bold !== undefined ||
                              args.italic !== undefined ||
                              args.textColor !== undefined;

      if (hasStyleOptions && args.text.length > 0) {
        const style: slides_v1.Schema$TextStyle = {};
        const fields: string[] = [];

        if (args.fontSize !== undefined) {
          style.fontSize = { magnitude: args.fontSize, unit: 'PT' };
          fields.push('fontSize');
        }
        if (args.fontFamily) {
          style.fontFamily = args.fontFamily;
          fields.push('fontFamily');
        }
        if (args.bold !== undefined) {
          style.bold = args.bold;
          fields.push('bold');
        }
        if (args.italic !== undefined) {
          style.italic = args.italic;
          fields.push('italic');
        }
        if (args.textColor) {
          const rgbColor = SlidesHelpers.hexToRgbColor(args.textColor);
          if (rgbColor) {
            style.foregroundColor = { opaqueColor: { rgbColor } };
            fields.push('foregroundColor');
          }
        }

        if (fields.length > 0) {
          requests.push({
            updateTextStyle: {
              objectId: args.tableObjectId,
              cellLocation,
              style,
              textRange: { type: 'ALL' },
              fields: fields.join(','),
            },
          });
        }
      }

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, requests);

      return JSON.stringify({
        success: true,
        tableObjectId: args.tableObjectId,
        cell: { rowIndex: args.rowIndex, columnIndex: args.columnIndex },
        textLength: args.text.length,
        styled: hasStyleOptions,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error editing table cell: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to edit table cell: ${error.message}`);
    }
  }
});

// --- Modify Tools ---

server.addTool({
  name: 'deleteSlide',
  description: 'Deletes a slide from the presentation. Cannot delete the only slide.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide to delete.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Deleting slide ${args.pageObjectId} from presentation: ${args.presentationId}`);

    try {
      // First check if this is the only slide
      const presentation = await SlidesHelpers.getPresentation(slides, args.presentationId);
      if (!presentation.slides || presentation.slides.length <= 1) {
        throw new UserError('Cannot delete the only slide in a presentation. Add another slide first.');
      }

      const request = SlidesHelpers.buildDeleteObjectRequest(args.pageObjectId);
      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        deletedSlideId: args.pageObjectId,
        remainingSlides: presentation.slides.length - 1,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting slide: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete slide: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'deleteElement',
  description: 'Deletes an element (shape, image, table) from a slide.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    objectId: z.string().describe('The object ID of the element to delete.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Deleting element ${args.objectId} from presentation: ${args.presentationId}`);

    try {
      const request = SlidesHelpers.buildDeleteObjectRequest(args.objectId);
      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        deletedObjectId: args.objectId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting element: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete element: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'updateSpeakerNotes',
  description: 'Updates the speaker notes for a slide.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide.'),
    notes: z.string().describe('The speaker notes text content.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Updating speaker notes for slide ${args.pageObjectId}`);

    try {
      // Get the slide to find the speaker notes shape ID
      const slide = await SlidesHelpers.getSlide(slides, args.presentationId, args.pageObjectId);
      const notesShapeId = SlidesHelpers.getSpeakerNotesShapeId(slide);

      if (!notesShapeId) {
        throw new UserError('Could not find speaker notes shape for this slide.');
      }

      const requests: slides_v1.Schema$Request[] = [];

      // Delete existing text first (if any)
      // We need to get the notes content length to delete properly
      const notesPage = slide.slideProperties?.notesPage;
      let hasExistingText = false;
      if (notesPage?.pageElements) {
        for (const el of notesPage.pageElements) {
          if (el.objectId === notesShapeId && el.shape?.text?.textElements) {
            const textLength = el.shape.text.textElements
              .filter((te: any) => te.textRun?.content)
              .map((te: any) => te.textRun?.content?.length || 0)
              .reduce((a: number, b: number) => a + b, 0);
            if (textLength > 0) {
              hasExistingText = true;
              requests.push(SlidesHelpers.buildDeleteTextRequest(notesShapeId, 0, textLength));
            }
            break;
          }
        }
      }

      // Insert new notes text
      requests.push(SlidesHelpers.buildInsertTextRequest(notesShapeId, args.notes, 0));

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, requests);

      return JSON.stringify({
        success: true,
        slideId: args.pageObjectId,
        notesShapeId: notesShapeId,
        notesLength: args.notes.length,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error updating speaker notes: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to update speaker notes: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'moveSlide',
  description: 'Moves a slide to a new position in the presentation.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    pageObjectId: z.string().describe('The object ID of the slide to move.'),
    newIndex: z.number().int().min(0).describe('The new 0-based index position for the slide.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Moving slide ${args.pageObjectId} to index ${args.newIndex}`);

    try {
      const request = SlidesHelpers.buildUpdateSlidesPositionRequest(
        [args.pageObjectId],
        args.newIndex
      );

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, [request]);

      return JSON.stringify({
        success: true,
        slideId: args.pageObjectId,
        newIndex: args.newIndex,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error moving slide: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to move slide: ${error.message}`);
    }
  }
});

server.addTool({
  name: 'insertTextInElement',
  description: 'Inserts or replaces text in a shape or text box element.',
  parameters: z.object({
    presentationId: z.string().describe('The ID of the presentation.'),
    objectId: z.string().describe('The object ID of the shape/text box to insert text into.'),
    text: z.string().describe('The text to insert.'),
    insertionIndex: z.number().int().min(0).optional().describe('Optional index where to insert text (0 for beginning). Omit to replace all text.'),
    replaceAll: z.boolean().optional().default(false).describe('If true, deletes all existing text before inserting new text.'),
  }),
  execute: async (args, { log }) => {
    const slides = await getSlidesClient();
    log.info(`Inserting text into element ${args.objectId}`);

    try {
      const requests: slides_v1.Schema$Request[] = [];

      // If replaceAll, we need to delete existing text first
      // This requires knowing the current text length, which we'd need to fetch
      // For simplicity, we'll delete a large range if replaceAll is true
      if (args.replaceAll) {
        // Delete all text (using a large end index - API handles out-of-bounds)
        requests.push({
          deleteText: {
            objectId: args.objectId,
            textRange: {
              type: 'ALL',
            },
          },
        });
      }

      // Insert the new text
      requests.push(SlidesHelpers.buildInsertTextRequest(
        args.objectId,
        args.text,
        args.replaceAll ? 0 : (args.insertionIndex ?? 0)
      ));

      await SlidesHelpers.executeBatchUpdate(slides, args.presentationId, requests);

      return JSON.stringify({
        success: true,
        objectId: args.objectId,
        textInserted: args.text.length,
        replacedAll: args.replaceAll,
        insertionIndex: args.replaceAll ? 0 : (args.insertionIndex ?? 0),
      }, null, 2);
    } catch (error: any) {
      log.error(`Error inserting text: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to insert text: ${error.message}`);
    }
  }
});

// ========================================
// === GMAIL TOOLS ===
// ========================================

// --- send_email ---
server.addTool({
  name: 'send_email',
  description: 'Sends a new email with optional attachments. Supports plain text, HTML, and multipart/alternative formats.',
  parameters: SendEmailParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Sending email to: ${args.to.join(', ')}`);

    try {
      const rawEmail = await GmailHelpers.createEmailWithAttachments({
        to: args.to,
        cc: args.cc,
        bcc: args.bcc,
        subject: args.subject,
        body: args.body,
        htmlBody: args.htmlBody,
        mimeType: args.mimeType,
        attachments: args.attachments,
        inReplyTo: args.inReplyTo,
      });

      const message = await GmailHelpers.sendEmail(gmail, rawEmail, args.threadId);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        threadId: message.threadId,
        to: args.to,
        subject: args.subject,
        attachmentsCount: args.attachments?.length || 0,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error sending email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to send email: ${error.message}`);
    }
  }
});

// --- draft_email ---
server.addTool({
  name: 'draft_email',
  description: 'Creates a new email draft with optional attachments.',
  parameters: DraftEmailParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Creating draft for: ${args.to.join(', ')}`);

    try {
      const rawEmail = await GmailHelpers.createEmailWithAttachments({
        to: args.to,
        cc: args.cc,
        bcc: args.bcc,
        subject: args.subject,
        body: args.body,
        htmlBody: args.htmlBody,
        mimeType: args.mimeType,
        attachments: args.attachments,
        inReplyTo: args.inReplyTo,
      });

      const draft = await GmailHelpers.createDraft(gmail, rawEmail, args.threadId);

      return JSON.stringify({
        success: true,
        draftId: draft.id,
        messageId: draft.message?.id,
        to: args.to,
        subject: args.subject,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error creating draft: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to create draft: ${error.message}`);
    }
  }
});

// --- read_email ---
server.addTool({
  name: 'read_email',
  description: 'Retrieves the content of a specific email message including headers, body, and attachment information.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Reading email: ${args.messageId}`);

    try {
      const message = await GmailHelpers.getMessage(gmail, args.messageId);
      const formatted = GmailHelpers.formatMessage(message);

      return JSON.stringify({
        id: formatted.id,
        threadId: formatted.threadId,
        from: formatted.from,
        to: formatted.to,
        cc: formatted.cc,
        subject: formatted.subject,
        date: formatted.date,
        snippet: formatted.snippet,
        body: formatted.body,
        hasHtmlBody: !!formatted.htmlBody,
        attachments: formatted.attachments,
        labels: formatted.labelIds,
        isUnread: formatted.isUnread,
        isStarred: formatted.isStarred,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error reading email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to read email: ${error.message}`);
    }
  }
});

// --- search_emails ---
server.addTool({
  name: 'search_emails',
  description: 'Searches for emails using Gmail search syntax (e.g., "from:user@example.com", "is:unread", "subject:meeting").',
  parameters: GmailSearchParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Searching emails with query: ${args.query}`);

    try {
      const results = await GmailHelpers.searchMessages(gmail, {
        query: args.query,
        maxResults: args.maxResults,
        pageToken: args.pageToken,
        includeSpamTrash: args.includeSpamTrash,
      });

      const formattedMessages = results.messages.map(msg => {
        const headers = GmailHelpers.parseEmailHeaders(msg.payload?.headers || []);
        return {
          id: msg.id,
          threadId: msg.threadId,
          from: headers['from'],
          subject: headers['subject'] || '(No Subject)',
          date: headers['date'],
          snippet: msg.snippet,
          labels: msg.labelIds,
        };
      });

      return JSON.stringify({
        query: args.query,
        resultCount: formattedMessages.length,
        resultSizeEstimate: results.resultSizeEstimate,
        nextPageToken: results.nextPageToken,
        messages: formattedMessages,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error searching emails: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to search emails: ${error.message}`);
    }
  }
});

// --- modify_email ---
server.addTool({
  name: 'modify_email',
  description: 'Modifies email labels (add or remove). Use to move emails between folders, mark as read/unread, star/unstar.',
  parameters: ModifyLabelsParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Modifying labels for email: ${args.messageId}`);

    try {
      const message = await GmailHelpers.modifyMessageLabels(
        gmail,
        args.messageId,
        args.addLabelIds,
        args.removeLabelIds
      );

      return JSON.stringify({
        success: true,
        messageId: message.id,
        currentLabels: message.labelIds,
        addedLabels: args.addLabelIds || [],
        removedLabels: args.removeLabelIds || [],
      }, null, 2);
    } catch (error: any) {
      log.error(`Error modifying email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to modify email: ${error.message}`);
    }
  }
});

// --- delete_email ---
server.addTool({
  name: 'delete_email',
  description: 'Permanently deletes an email. This action cannot be undone. Use trash_email for recoverable deletion.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Permanently deleting email: ${args.messageId}`);

    try {
      await GmailHelpers.deleteMessage(gmail, args.messageId);

      return JSON.stringify({
        success: true,
        messageId: args.messageId,
        action: 'permanently_deleted',
        warning: 'This action cannot be undone.',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete email: ${error.message}`);
    }
  }
});

// --- download_attachment ---
server.addTool({
  name: 'download_attachment',
  description: 'Downloads an email attachment to a specified location.',
  parameters: DownloadAttachmentParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Downloading attachment ${args.attachmentId} from message ${args.messageId}`);

    try {
      const result = await GmailHelpers.downloadAttachment(
        gmail,
        args.messageId,
        args.attachmentId,
        args.savePath,
        args.filename
      );

      return JSON.stringify({
        success: true,
        savedTo: result.savedTo,
        sizeBytes: result.size,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error downloading attachment: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to download attachment: ${error.message}`);
    }
  }
});

// --- list_email_labels ---
server.addTool({
  name: 'list_email_labels',
  description: 'Retrieves all available Gmail labels (both system and user-created).',
  parameters: z.object({}),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info('Listing all Gmail labels');

    try {
      const labels = await LabelManager.listLabels(gmail);

      return JSON.stringify({
        totalLabels: labels.length,
        systemLabels: labels.filter(l => l.type === 'system').length,
        userLabels: labels.filter(l => l.type === 'user').length,
        labels: labels,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error listing labels: ${error.message}`);
      throw new UserError(`Failed to list labels: ${error.message}`);
    }
  }
});

// --- create_label ---
server.addTool({
  name: 'create_label',
  description: 'Creates a new Gmail label.',
  parameters: CreateLabelParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Creating label: ${args.name}`);

    try {
      const label = await LabelManager.createLabel(gmail, {
        name: args.name,
        labelListVisibility: args.labelListVisibility,
        messageListVisibility: args.messageListVisibility,
      });

      return JSON.stringify({
        success: true,
        label: label,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error creating label: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to create label: ${error.message}`);
    }
  }
});

// --- update_label ---
server.addTool({
  name: 'update_label',
  description: 'Updates an existing Gmail label (rename or change visibility).',
  parameters: z.object({
    id: z.string().describe('ID of the label to update.'),
    name: z.string().optional().describe('New name for the label.'),
    labelListVisibility: z.enum(['labelShow', 'labelShowIfUnread', 'labelHide']).optional()
      .describe('Visibility of the label in the label list.'),
    messageListVisibility: z.enum(['show', 'hide']).optional()
      .describe('Whether to show or hide the label in the message list.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Updating label: ${args.id}`);

    try {
      const label = await LabelManager.updateLabel(gmail, args.id, {
        name: args.name,
        labelListVisibility: args.labelListVisibility,
        messageListVisibility: args.messageListVisibility,
      });

      return JSON.stringify({
        success: true,
        label: label,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error updating label: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to update label: ${error.message}`);
    }
  }
});

// --- delete_label ---
server.addTool({
  name: 'delete_label',
  description: 'Deletes a Gmail label. System labels cannot be deleted.',
  parameters: z.object({
    id: z.string().describe('ID of the label to delete.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Deleting label: ${args.id}`);

    try {
      await LabelManager.deleteLabel(gmail, args.id);

      return JSON.stringify({
        success: true,
        deletedLabelId: args.id,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting label: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete label: ${error.message}`);
    }
  }
});

// --- get_or_create_label ---
server.addTool({
  name: 'get_or_create_label',
  description: 'Gets an existing label by name or creates it if it does not exist. Idempotent operation.',
  parameters: CreateLabelParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Getting or creating label: ${args.name}`);

    try {
      const label = await LabelManager.getOrCreateLabel(gmail, {
        name: args.name,
        labelListVisibility: args.labelListVisibility,
        messageListVisibility: args.messageListVisibility,
      });

      return JSON.stringify({
        success: true,
        label: label,
        created: !!(await LabelManager.findLabelByName(gmail, args.name)),
      }, null, 2);
    } catch (error: any) {
      log.error(`Error getting/creating label: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get or create label: ${error.message}`);
    }
  }
});

// --- batch_modify_emails ---
server.addTool({
  name: 'batch_modify_emails',
  description: 'Modifies labels for multiple emails in batches. More efficient than individual modifications.',
  parameters: BatchModifyLabelsParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Batch modifying ${args.messageIds.length} emails`);

    try {
      const result = await GmailHelpers.batchModifyMessages(
        gmail,
        args.messageIds,
        args.addLabelIds,
        args.removeLabelIds,
        args.batchSize
      );

      return JSON.stringify({
        success: result.errors.length === 0,
        totalMessages: args.messageIds.length,
        processedCount: result.processed,
        errors: result.errors,
        addedLabels: args.addLabelIds || [],
        removedLabels: args.removeLabelIds || [],
      }, null, 2);
    } catch (error: any) {
      log.error(`Error batch modifying emails: ${error.message}`);
      throw new UserError(`Failed to batch modify emails: ${error.message}`);
    }
  }
});

// --- batch_delete_emails ---
server.addTool({
  name: 'batch_delete_emails',
  description: 'Permanently deletes multiple emails in batches. This action cannot be undone.',
  parameters: BatchDeleteParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Batch deleting ${args.messageIds.length} emails`);

    try {
      const result = await GmailHelpers.batchDeleteMessages(
        gmail,
        args.messageIds,
        args.batchSize
      );

      return JSON.stringify({
        success: result.errors.length === 0,
        totalMessages: args.messageIds.length,
        deletedCount: result.deleted,
        errors: result.errors,
        warning: 'Permanently deleted messages cannot be recovered.',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error batch deleting emails: ${error.message}`);
      throw new UserError(`Failed to batch delete emails: ${error.message}`);
    }
  }
});

// --- create_filter ---
server.addTool({
  name: 'create_filter',
  description: 'Creates a new Gmail filter with custom criteria and actions.',
  parameters: CreateFilterParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info('Creating new Gmail filter');

    try {
      const filter = await FilterManager.createFilter(gmail, args.criteria, args.action);

      return JSON.stringify({
        success: true,
        filter: filter,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error creating filter: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to create filter: ${error.message}`);
    }
  }
});

// --- list_filters ---
server.addTool({
  name: 'list_filters',
  description: 'Retrieves all Gmail filters.',
  parameters: z.object({}),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info('Listing all Gmail filters');

    try {
      const filters = await FilterManager.listFilters(gmail);

      return JSON.stringify({
        totalFilters: filters.length,
        filters: filters,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error listing filters: ${error.message}`);
      throw new UserError(`Failed to list filters: ${error.message}`);
    }
  }
});

// --- get_filter ---
server.addTool({
  name: 'get_filter',
  description: 'Gets details of a specific Gmail filter.',
  parameters: FilterIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Getting filter: ${args.filterId}`);

    try {
      const filter = await FilterManager.getFilter(gmail, args.filterId);

      return JSON.stringify({
        filter: filter,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error getting filter: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get filter: ${error.message}`);
    }
  }
});

// --- delete_filter ---
server.addTool({
  name: 'delete_filter',
  description: 'Deletes a Gmail filter.',
  parameters: FilterIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Deleting filter: ${args.filterId}`);

    try {
      await FilterManager.deleteFilter(gmail, args.filterId);

      return JSON.stringify({
        success: true,
        deletedFilterId: args.filterId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting filter: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete filter: ${error.message}`);
    }
  }
});

// --- create_filter_from_template ---
server.addTool({
  name: 'create_filter_from_template',
  description: 'Creates a filter using a pre-defined template for common scenarios (fromSender, withSubject, withAttachments, largeEmails, containingText, mailingList).',
  parameters: FilterTemplateParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Creating filter from template: ${args.template}`);

    try {
      const filter = await FilterManager.createFilterFromTemplate(
        gmail,
        args.template,
        args.parameters
      );

      return JSON.stringify({
        success: true,
        template: args.template,
        filter: filter,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error creating filter from template: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to create filter: ${error.message}`);
    }
  }
});

// --- get_thread ---
server.addTool({
  name: 'get_thread',
  description: 'Gets a full email conversation thread with all messages.',
  parameters: ThreadIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Getting thread: ${args.threadId}`);

    try {
      const thread = await GmailHelpers.getThread(gmail, args.threadId);

      const formattedMessages = (thread.messages || []).map(msg => {
        const formatted = GmailHelpers.formatMessage(msg);
        return {
          id: formatted.id,
          from: formatted.from,
          to: formatted.to,
          subject: formatted.subject,
          date: formatted.date,
          snippet: formatted.snippet,
          body: formatted.body,
          attachments: formatted.attachments.length,
        };
      });

      return JSON.stringify({
        threadId: thread.id,
        messageCount: formattedMessages.length,
        messages: formattedMessages,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error getting thread: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get thread: ${error.message}`);
    }
  }
});

// --- list_threads ---
server.addTool({
  name: 'list_threads',
  description: 'Lists email threads with optional search query.',
  parameters: z.object({
    query: z.string().optional().describe('Gmail search query to filter threads.'),
    maxResults: z.number().int().min(1).max(500).optional().default(10)
      .describe('Maximum number of threads to return.'),
    pageToken: z.string().optional().describe('Page token for pagination.'),
    includeSpamTrash: z.boolean().optional().default(false)
      .describe('Include threads from SPAM and TRASH.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Listing threads${args.query ? ` with query: ${args.query}` : ''}`);

    try {
      const result = await GmailHelpers.listThreads(gmail, {
        query: args.query,
        maxResults: args.maxResults,
        pageToken: args.pageToken,
        includeSpamTrash: args.includeSpamTrash,
      });

      return JSON.stringify({
        threadCount: result.threads.length,
        resultSizeEstimate: result.resultSizeEstimate,
        nextPageToken: result.nextPageToken,
        threads: result.threads.map(t => ({
          id: t.id,
          snippet: t.snippet,
          historyId: t.historyId,
        })),
      }, null, 2);
    } catch (error: any) {
      log.error(`Error listing threads: ${error.message}`);
      throw new UserError(`Failed to list threads: ${error.message}`);
    }
  }
});

// --- reply_to_email ---
server.addTool({
  name: 'reply_to_email',
  description: 'Sends a reply to an existing email, maintaining the conversation thread.',
  parameters: z.object({
    messageId: z.string().describe('ID of the message to reply to.'),
    body: z.string().describe('Reply message body.'),
    htmlBody: z.string().optional().describe('HTML version of the reply body.'),
    mimeType: z.enum(['text/plain', 'text/html', 'multipart/alternative']).optional().default('text/plain')
      .describe('Content type of the reply.'),
    cc: z.array(z.string().email()).optional().describe('Additional CC recipients.'),
    bcc: z.array(z.string().email()).optional().describe('BCC recipients.'),
    attachments: z.array(z.string()).optional().describe('File paths to attach.'),
    replyAll: z.boolean().optional().default(false).describe('Reply to all recipients.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Replying to message: ${args.messageId}`);

    try {
      // Get original message to extract headers
      const originalMessage = await GmailHelpers.getMessage(gmail, args.messageId, 'metadata');
      const headers = GmailHelpers.parseEmailHeaders(originalMessage.payload?.headers || []);

      const to = args.replyAll
        ? [headers['from'], ...(headers['to']?.split(',').map(s => s.trim()) || [])].filter(Boolean)
        : [headers['from']];

      const subject = headers['subject']?.startsWith('Re:')
        ? headers['subject']
        : `Re: ${headers['subject'] || ''}`;

      const messageId = headers['message-id'];

      const rawEmail = await GmailHelpers.createEmailWithAttachments({
        to,
        cc: args.cc,
        bcc: args.bcc,
        subject,
        body: args.body,
        htmlBody: args.htmlBody,
        mimeType: args.mimeType,
        attachments: args.attachments,
        inReplyTo: messageId,
      });

      const message = await GmailHelpers.sendEmail(gmail, rawEmail, originalMessage.threadId || undefined);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        threadId: message.threadId,
        inReplyTo: args.messageId,
        to: to,
        subject: subject,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error replying to email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to reply to email: ${error.message}`);
    }
  }
});

// --- forward_email ---
server.addTool({
  name: 'forward_email',
  description: 'Forwards an email to new recipients.',
  parameters: z.object({
    messageId: z.string().describe('ID of the message to forward.'),
    to: z.array(z.string().email()).describe('Recipients to forward to.'),
    additionalMessage: z.string().optional().describe('Additional message to include with the forward.'),
    cc: z.array(z.string().email()).optional().describe('CC recipients.'),
    bcc: z.array(z.string().email()).optional().describe('BCC recipients.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Forwarding message ${args.messageId} to: ${args.to.join(', ')}`);

    try {
      // Get original message
      const originalMessage = await GmailHelpers.getMessage(gmail, args.messageId);
      const formatted = GmailHelpers.formatMessage(originalMessage);

      const subject = formatted.subject.startsWith('Fwd:')
        ? formatted.subject
        : `Fwd: ${formatted.subject}`;

      let body = args.additionalMessage ? `${args.additionalMessage}\n\n` : '';
      body += `---------- Forwarded message ---------\n`;
      body += `From: ${formatted.from}\n`;
      body += `Date: ${formatted.date}\n`;
      body += `Subject: ${formatted.subject}\n`;
      body += `To: ${formatted.to}\n\n`;
      body += formatted.body;

      const rawEmail = GmailHelpers.createSimpleEmail({
        to: args.to,
        cc: args.cc,
        bcc: args.bcc,
        subject,
        body,
        mimeType: 'text/plain',
      });

      const message = await GmailHelpers.sendEmail(gmail, rawEmail);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        threadId: message.threadId,
        forwardedFrom: args.messageId,
        to: args.to,
        subject: subject,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error forwarding email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to forward email: ${error.message}`);
    }
  }
});

// --- trash_email ---
server.addTool({
  name: 'trash_email',
  description: 'Moves an email to trash. Can be restored using untrash_email.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Moving to trash: ${args.messageId}`);

    try {
      const message = await GmailHelpers.trashMessage(gmail, args.messageId);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        action: 'moved_to_trash',
        canRestore: true,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error trashing email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to trash email: ${error.message}`);
    }
  }
});

// --- untrash_email ---
server.addTool({
  name: 'untrash_email',
  description: 'Restores an email from trash.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Restoring from trash: ${args.messageId}`);

    try {
      const message = await GmailHelpers.untrashMessage(gmail, args.messageId);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        action: 'restored_from_trash',
        currentLabels: message.labelIds,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error untrashing email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to untrash email: ${error.message}`);
    }
  }
});

// --- archive_email ---
server.addTool({
  name: 'archive_email',
  description: 'Archives an email by removing it from the inbox (keeps in "All Mail").',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Archiving email: ${args.messageId}`);

    try {
      const message = await GmailHelpers.modifyMessageLabels(
        gmail,
        args.messageId,
        undefined,
        ['INBOX']
      );

      return JSON.stringify({
        success: true,
        messageId: message.id,
        action: 'archived',
        currentLabels: message.labelIds,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error archiving email: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to archive email: ${error.message}`);
    }
  }
});

// --- mark_as_read ---
server.addTool({
  name: 'mark_as_read',
  description: 'Marks an email as read.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Marking as read: ${args.messageId}`);

    try {
      const message = await GmailHelpers.modifyMessageLabels(
        gmail,
        args.messageId,
        undefined,
        ['UNREAD']
      );

      return JSON.stringify({
        success: true,
        messageId: message.id,
        action: 'marked_as_read',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error marking as read: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to mark as read: ${error.message}`);
    }
  }
});

// --- mark_as_unread ---
server.addTool({
  name: 'mark_as_unread',
  description: 'Marks an email as unread.',
  parameters: MessageIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Marking as unread: ${args.messageId}`);

    try {
      const message = await GmailHelpers.modifyMessageLabels(
        gmail,
        args.messageId,
        ['UNREAD'],
        undefined
      );

      return JSON.stringify({
        success: true,
        messageId: message.id,
        action: 'marked_as_unread',
      }, null, 2);
    } catch (error: any) {
      log.error(`Error marking as unread: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to mark as unread: ${error.message}`);
    }
  }
});

// --- get_user_profile ---
server.addTool({
  name: 'get_user_profile',
  description: 'Gets the authenticated Gmail user profile (email address and message counts).',
  parameters: z.object({}),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info('Getting user profile');

    try {
      const profile = await GmailHelpers.getUserProfile(gmail);

      return JSON.stringify({
        emailAddress: profile.emailAddress,
        messagesTotal: profile.messagesTotal,
        threadsTotal: profile.threadsTotal,
        historyId: profile.historyId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error getting profile: ${error.message}`);
      throw new UserError(`Failed to get user profile: ${error.message}`);
    }
  }
});

// --- list_drafts ---
server.addTool({
  name: 'list_drafts',
  description: 'Lists all email drafts.',
  parameters: z.object({
    maxResults: z.number().int().min(1).max(500).optional().default(10)
      .describe('Maximum number of drafts to return.'),
    pageToken: z.string().optional().describe('Page token for pagination.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info('Listing drafts');

    try {
      const result = await GmailHelpers.listDrafts(gmail, {
        maxResults: args.maxResults,
        pageToken: args.pageToken,
      });

      return JSON.stringify({
        draftCount: result.drafts.length,
        nextPageToken: result.nextPageToken,
        drafts: result.drafts.map(d => ({
          id: d.id,
          messageId: d.message?.id,
        })),
      }, null, 2);
    } catch (error: any) {
      log.error(`Error listing drafts: ${error.message}`);
      throw new UserError(`Failed to list drafts: ${error.message}`);
    }
  }
});

// --- get_draft ---
server.addTool({
  name: 'get_draft',
  description: 'Gets the content of a specific draft.',
  parameters: DraftIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Getting draft: ${args.draftId}`);

    try {
      const draft = await GmailHelpers.getDraft(gmail, args.draftId);
      const formatted = draft.message ? GmailHelpers.formatMessage(draft.message) : null;

      return JSON.stringify({
        draftId: draft.id,
        message: formatted ? {
          to: formatted.to,
          cc: formatted.cc,
          subject: formatted.subject,
          body: formatted.body,
          attachments: formatted.attachments,
        } : null,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error getting draft: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to get draft: ${error.message}`);
    }
  }
});

// --- update_draft ---
server.addTool({
  name: 'update_draft',
  description: 'Updates an existing draft with new content.',
  parameters: z.object({
    draftId: z.string().describe('ID of the draft to update.'),
    to: z.array(z.string().email()).describe('List of recipient email addresses.'),
    subject: z.string().describe('Email subject.'),
    body: z.string().describe('Email body content.'),
    htmlBody: z.string().optional().describe('HTML version of the email body.'),
    mimeType: z.enum(['text/plain', 'text/html', 'multipart/alternative']).optional().default('text/plain')
      .describe('Content type of the email.'),
    cc: z.array(z.string().email()).optional().describe('CC recipients.'),
    bcc: z.array(z.string().email()).optional().describe('BCC recipients.'),
  }),
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Updating draft: ${args.draftId}`);

    try {
      const rawEmail = GmailHelpers.createSimpleEmail({
        to: args.to,
        cc: args.cc,
        bcc: args.bcc,
        subject: args.subject,
        body: args.body,
        htmlBody: args.htmlBody,
        mimeType: args.mimeType,
      });

      const draft = await GmailHelpers.updateDraft(gmail, args.draftId, rawEmail);

      return JSON.stringify({
        success: true,
        draftId: draft.id,
        messageId: draft.message?.id,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error updating draft: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to update draft: ${error.message}`);
    }
  }
});

// --- delete_draft ---
server.addTool({
  name: 'delete_draft',
  description: 'Deletes a draft. This action cannot be undone.',
  parameters: DraftIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Deleting draft: ${args.draftId}`);

    try {
      await GmailHelpers.deleteDraft(gmail, args.draftId);

      return JSON.stringify({
        success: true,
        deletedDraftId: args.draftId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error deleting draft: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to delete draft: ${error.message}`);
    }
  }
});

// --- send_draft ---
server.addTool({
  name: 'send_draft',
  description: 'Sends an existing draft as an email.',
  parameters: DraftIdParameter,
  execute: async (args, { log }) => {
    const gmail = await getGmailClient();
    log.info(`Sending draft: ${args.draftId}`);

    try {
      const message = await GmailHelpers.sendDraft(gmail, args.draftId);

      return JSON.stringify({
        success: true,
        messageId: message.id,
        threadId: message.threadId,
        sentDraftId: args.draftId,
      }, null, 2);
    } catch (error: any) {
      log.error(`Error sending draft: ${error.message}`);
      if (error instanceof UserError) throw error;
      throw new UserError(`Failed to send draft: ${error.message}`);
    }
  }
});

// ========================================
// === END GMAIL TOOLS ===
// ========================================

// --- Server Startup ---
async function startServer() {
try {
await initializeGoogleClient(); // Authorize BEFORE starting listeners
console.error("Starting Ultimate Google Docs, Sheets & Slides MCP server...");

      // Using stdio as before
      const configToUse = {
          transportType: "stdio" as const,
      };

      // Start the server with proper error handling
      server.start(configToUse);
      console.error(`MCP Server running using ${configToUse.transportType}. Awaiting client connection...`);

      // Log that error handling has been enabled
      console.error('Process-level error handling configured to prevent crashes from timeout errors.');

} catch(startError: any) {
console.error("FATAL: Server failed to start:", startError.message || startError);
process.exit(1);
}
}

startServer(); // Removed .catch here, let errors propagate if startup fails critically
