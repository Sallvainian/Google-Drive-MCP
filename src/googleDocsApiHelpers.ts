// src/googleDocsApiHelpers.ts
import { google, docs_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { UserError } from 'fastmcp';
import { TextStyleArgs, ParagraphStyleArgs, hexToRgbColor, NotImplementedError } from './types.js';

type Docs = docs_v1.Docs; // Alias for convenience

// --- Constants ---
const MAX_BATCH_UPDATE_REQUESTS = 50; // Google API limits batch size

// --- Core Helper to Execute Batch Updates ---
export async function executeBatchUpdate(docs: Docs, documentId: string, requests: docs_v1.Schema$Request[]): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
if (!requests || requests.length === 0) {
// console.warn("executeBatchUpdate called with no requests.");
return {}; // Nothing to do
}

    // TODO: Consider splitting large request arrays into multiple batches if needed
    if (requests.length > MAX_BATCH_UPDATE_REQUESTS) {
         console.warn(`Attempting batch update with ${requests.length} requests, exceeding typical limits. May fail.`);
    }

    try {
        const response = await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: { requests },
        });
        return response.data;
    } catch (error: any) {
        console.error(`Google API batchUpdate Error for doc ${documentId}:`, error.response?.data || error.message);
        // Translate common API errors to UserErrors
        if (error.code === 400 && error.message.includes('Invalid requests')) {
             // Try to extract more specific info if available
             const details = error.response?.data?.error?.details;
             let detailMsg = '';
             if (details && Array.isArray(details)) {
                 detailMsg = details.map(d => d.description || JSON.stringify(d)).join('; ');
             }
            throw new UserError(`Invalid request sent to Google Docs API. Details: ${detailMsg || error.message}`);
        }
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}). Check the ID.`);
        if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}). Ensure the authenticated user has edit access.`);
        // Generic internal error for others
        throw new Error(`Google API Error (${error.code}): ${error.message}`);
    }

}

// --- Text Finding Helper ---
// This improved version is more robust in handling various text structure scenarios
export async function findTextRange(docs: Docs, documentId: string, textToFind: string, instance: number = 1): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    // Request more detailed information about the document structure
    const res = await docs.documents.get({
        documentId,
        // Request more fields to handle various container types (not just paragraphs)
        fields: 'body(content(paragraph(elements(startIndex,endIndex,textRun(content))),table,sectionBreak,tableOfContents,startIndex,endIndex))',
    });

    if (!res.data.body?.content) {
        console.warn(`No content found in document ${documentId}`);
        return null;
    }

    // More robust text collection and index tracking
    let fullText = '';
    const segments: { text: string, start: number, end: number }[] = [];

    // Process all content elements, including structural ones
    const collectTextFromContent = (content: any[]) => {
        content.forEach(element => {
            // Handle paragraph elements
            if (element.paragraph?.elements) {
                element.paragraph.elements.forEach((pe: any) => {
                    if (pe.textRun?.content && pe.startIndex !== undefined && pe.endIndex !== undefined) {
                        const content = pe.textRun.content;
                        fullText += content;
                        segments.push({
                            text: content,
                            start: pe.startIndex,
                            end: pe.endIndex
                        });
                    }
                });
            }

            // Handle table elements - this is simplified and might need expansion
            if (element.table && element.table.tableRows) {
                element.table.tableRows.forEach((row: any) => {
                    if (row.tableCells) {
                        row.tableCells.forEach((cell: any) => {
                            if (cell.content) {
                                collectTextFromContent(cell.content);
                            }
                        });
                    }
                });
            }

            // Add handling for other structural elements as needed
        });
    };

    collectTextFromContent(res.data.body.content);

    // Sort segments by starting position to ensure correct ordering
    segments.sort((a, b) => a.start - b.start);

    console.log(`Document ${documentId} contains ${segments.length} text segments and ${fullText.length} characters in total.`);

    // Find the specified instance of the text
    let startIndex = -1;
    let endIndex = -1;
    let foundCount = 0;
    let searchStartIndex = 0;

    while (foundCount < instance) {
        const currentIndex = fullText.indexOf(textToFind, searchStartIndex);
        if (currentIndex === -1) {
            console.log(`Search text "${textToFind}" not found for instance ${foundCount + 1} (requested: ${instance})`);
            break;
        }

        foundCount++;
        console.log(`Found instance ${foundCount} of "${textToFind}" at position ${currentIndex} in full text`);

        if (foundCount === instance) {
            const targetStartInFullText = currentIndex;
            const targetEndInFullText = currentIndex + textToFind.length;
            let currentPosInFullText = 0;

            console.log(`Target text range in full text: ${targetStartInFullText}-${targetEndInFullText}`);

            for (const seg of segments) {
                const segStartInFullText = currentPosInFullText;
                const segTextLength = seg.text.length;
                const segEndInFullText = segStartInFullText + segTextLength;

                // Map from reconstructed text position to actual document indices
                if (startIndex === -1 && targetStartInFullText >= segStartInFullText && targetStartInFullText < segEndInFullText) {
                    startIndex = seg.start + (targetStartInFullText - segStartInFullText);
                    console.log(`Mapped start to segment ${seg.start}-${seg.end}, position ${startIndex}`);
                }

                if (targetEndInFullText > segStartInFullText && targetEndInFullText <= segEndInFullText) {
                    endIndex = seg.start + (targetEndInFullText - segStartInFullText);
                    console.log(`Mapped end to segment ${seg.start}-${seg.end}, position ${endIndex}`);
                    break;
                }

                currentPosInFullText = segEndInFullText;
            }

            if (startIndex === -1 || endIndex === -1) {
                console.warn(`Failed to map text "${textToFind}" instance ${instance} to actual document indices`);
                // Reset and try next occurrence
                startIndex = -1;
                endIndex = -1;
                searchStartIndex = currentIndex + 1;
                foundCount--;
                continue;
            }

            console.log(`Successfully mapped "${textToFind}" to document range ${startIndex}-${endIndex}`);
            return { startIndex, endIndex };
        }

        // Prepare for next search iteration
        searchStartIndex = currentIndex + 1;
    }

    console.warn(`Could not find instance ${instance} of text "${textToFind}" in document ${documentId}`);
    return null; // Instance not found or mapping failed for all attempts
} catch (error: any) {
    console.error(`Error finding text "${textToFind}" in doc ${documentId}: ${error.message || 'Unknown error'}`);
    if (error.code === 404) throw new UserError(`Document not found while searching text (ID: ${documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied while searching text in doc ${documentId}.`);
    throw new Error(`Failed to retrieve doc for text searching: ${error.message || 'Unknown error'}`);
}
}

// --- Paragraph Boundary Helper ---
// Enhanced version to handle document structural elements more robustly
export async function getParagraphRange(docs: Docs, documentId: string, indexWithin: number): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    console.log(`Finding paragraph containing index ${indexWithin} in document ${documentId}`);

    // Request more detailed document structure to handle nested elements
    const res = await docs.documents.get({
        documentId,
        // Request more comprehensive structure information
        fields: 'body(content(startIndex,endIndex,paragraph,table,sectionBreak,tableOfContents))',
    });

    if (!res.data.body?.content) {
        console.warn(`No content found in document ${documentId}`);
        return null;
    }

    // Find paragraph containing the index
    // We'll look at all structural elements recursively
    const findParagraphInContent = (content: any[]): { startIndex: number; endIndex: number } | null => {
        for (const element of content) {
            // Check if we have element boundaries defined
            if (element.startIndex !== undefined && element.endIndex !== undefined) {
                // Check if index is within this element's range first
                if (indexWithin >= element.startIndex && indexWithin < element.endIndex) {
                    // If it's a paragraph, we've found our target
                    if (element.paragraph) {
                        console.log(`Found paragraph containing index ${indexWithin}, range: ${element.startIndex}-${element.endIndex}`);
                        return {
                            startIndex: element.startIndex,
                            endIndex: element.endIndex
                        };
                    }

                    // If it's a table, we need to check cells recursively
                    if (element.table && element.table.tableRows) {
                        console.log(`Index ${indexWithin} is within a table, searching cells...`);
                        for (const row of element.table.tableRows) {
                            if (row.tableCells) {
                                for (const cell of row.tableCells) {
                                    if (cell.content) {
                                        const result = findParagraphInContent(cell.content);
                                        if (result) return result;
                                    }
                                }
                            }
                        }
                    }

                    // For other structural elements, we didn't find a paragraph
                    // but we know the index is within this element
                    console.warn(`Index ${indexWithin} is within element (${element.startIndex}-${element.endIndex}) but not in a paragraph`);
                }
            }
        }

        return null;
    };

    const paragraphRange = findParagraphInContent(res.data.body.content);

    if (!paragraphRange) {
        console.warn(`Could not find paragraph containing index ${indexWithin}`);
    } else {
        console.log(`Returning paragraph range: ${paragraphRange.startIndex}-${paragraphRange.endIndex}`);
    }

    return paragraphRange;

} catch (error: any) {
    console.error(`Error getting paragraph range for index ${indexWithin} in doc ${documentId}: ${error.message || 'Unknown error'}`);
    if (error.code === 404) throw new UserError(`Document not found while finding paragraph (ID: ${documentId}).`);
    if (error.code === 403) throw new UserError(`Permission denied while accessing doc ${documentId}.`);
    throw new Error(`Failed to find paragraph: ${error.message || 'Unknown error'}`);
}
}

// --- Style Request Builders ---

export function buildUpdateTextStyleRequest(
startIndex: number,
endIndex: number,
style: TextStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    const textStyle: docs_v1.Schema$TextStyle = {};
const fieldsToUpdate: string[] = [];

    if (style.bold !== undefined) { textStyle.bold = style.bold; fieldsToUpdate.push('bold'); }
    if (style.italic !== undefined) { textStyle.italic = style.italic; fieldsToUpdate.push('italic'); }
    if (style.underline !== undefined) { textStyle.underline = style.underline; fieldsToUpdate.push('underline'); }
    if (style.strikethrough !== undefined) { textStyle.strikethrough = style.strikethrough; fieldsToUpdate.push('strikethrough'); }
    if (style.fontSize !== undefined) { textStyle.fontSize = { magnitude: style.fontSize, unit: 'PT' }; fieldsToUpdate.push('fontSize'); }
    if (style.fontFamily !== undefined) { textStyle.weightedFontFamily = { fontFamily: style.fontFamily }; fieldsToUpdate.push('weightedFontFamily'); }
    if (style.foregroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.foregroundColor);
        if (!rgbColor) throw new UserError(`Invalid foreground hex color format: ${style.foregroundColor}`);
        textStyle.foregroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('foregroundColor');
    }
     if (style.backgroundColor !== undefined) {
        const rgbColor = hexToRgbColor(style.backgroundColor);
        if (!rgbColor) throw new UserError(`Invalid background hex color format: ${style.backgroundColor}`);
        textStyle.backgroundColor = { color: { rgbColor: rgbColor } }; fieldsToUpdate.push('backgroundColor');
    }
    if (style.linkUrl !== undefined) {
        textStyle.link = { url: style.linkUrl }; fieldsToUpdate.push('link');
    }
    // TODO: Handle clearing formatting

    if (fieldsToUpdate.length === 0) return null; // No styles to apply

    const request: docs_v1.Schema$Request = {
        updateTextStyle: {
            range: { startIndex, endIndex },
            textStyle: textStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
    return { request, fields: fieldsToUpdate };

}

export function buildUpdateParagraphStyleRequest(
startIndex: number,
endIndex: number,
style: ParagraphStyleArgs
): { request: docs_v1.Schema$Request, fields: string[] } | null {
    // Create style object and track which fields to update
    const paragraphStyle: docs_v1.Schema$ParagraphStyle = {};
    const fieldsToUpdate: string[] = [];

    console.log(`Building paragraph style request for range ${startIndex}-${endIndex} with options:`, style);

    // Process alignment option (LEFT, CENTER, RIGHT, JUSTIFIED)
    if (style.alignment !== undefined) {
        paragraphStyle.alignment = style.alignment;
        fieldsToUpdate.push('alignment');
        console.log(`Setting alignment to ${style.alignment}`);
    }

    // Process indentation options
    if (style.indentStart !== undefined) {
        paragraphStyle.indentStart = { magnitude: style.indentStart, unit: 'PT' };
        fieldsToUpdate.push('indentStart');
        console.log(`Setting left indent to ${style.indentStart}pt`);
    }

    if (style.indentEnd !== undefined) {
        paragraphStyle.indentEnd = { magnitude: style.indentEnd, unit: 'PT' };
        fieldsToUpdate.push('indentEnd');
        console.log(`Setting right indent to ${style.indentEnd}pt`);
    }

    // Process spacing options
    if (style.spaceAbove !== undefined) {
        paragraphStyle.spaceAbove = { magnitude: style.spaceAbove, unit: 'PT' };
        fieldsToUpdate.push('spaceAbove');
        console.log(`Setting space above to ${style.spaceAbove}pt`);
    }

    if (style.spaceBelow !== undefined) {
        paragraphStyle.spaceBelow = { magnitude: style.spaceBelow, unit: 'PT' };
        fieldsToUpdate.push('spaceBelow');
        console.log(`Setting space below to ${style.spaceBelow}pt`);
    }

    // Process named style types (headings, etc.)
    if (style.namedStyleType !== undefined) {
        paragraphStyle.namedStyleType = style.namedStyleType;
        fieldsToUpdate.push('namedStyleType');
        console.log(`Setting named style to ${style.namedStyleType}`);
    }

    // Process page break control
    if (style.keepWithNext !== undefined) {
        paragraphStyle.keepWithNext = style.keepWithNext;
        fieldsToUpdate.push('keepWithNext');
        console.log(`Setting keepWithNext to ${style.keepWithNext}`);
    }

    // Verify we have styles to apply
    if (fieldsToUpdate.length === 0) {
        console.warn("No paragraph styling options were provided");
        return null; // No styles to apply
    }

    // Build the request object
    const request: docs_v1.Schema$Request = {
        updateParagraphStyle: {
            range: { startIndex, endIndex },
            paragraphStyle: paragraphStyle,
            fields: fieldsToUpdate.join(','),
        }
    };

    console.log(`Created paragraph style request with fields: ${fieldsToUpdate.join(', ')}`);
    return { request, fields: fieldsToUpdate };
}

// --- Table Cell Range Helper ---

/**
 * Gets the content range of a specific table cell
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param tableStartIndex - The start index of the table element
 * @param rowIndex - 0-based row index
 * @param columnIndex - 0-based column index
 * @returns Object with cell and content boundaries, or null if not found
 */
export async function getTableCellRange(
    docs: Docs,
    documentId: string,
    tableStartIndex: number,
    rowIndex: number,
    columnIndex: number
): Promise<{
    cellStartIndex: number;
    cellEndIndex: number;
    contentStartIndex: number;
    contentEndIndex: number;
    paragraphEndIndex: number;  // Includes trailing newline for paragraph styling
} | null> {
    try {
        // Fetch document structure with table details
        const res = await docs.documents.get({
            documentId,
            fields: 'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(paragraph(elements(startIndex,endIndex))))))))',
        });

        if (!res.data.body?.content) {
            console.warn(`No content found in document ${documentId}`);
            return null;
        }

        // Find the table element at the specified start index
        let table: any = null;
        for (const element of res.data.body.content) {
            if (element.startIndex === tableStartIndex && element.table) {
                table = element.table;
                break;
            }
        }

        if (!table) {
            throw new UserError(`No table found at index ${tableStartIndex}. Use readGoogleDoc with format: json to find table indices.`);
        }

        // Validate row bounds
        const tableRows = table.tableRows || [];
        if (rowIndex < 0 || rowIndex >= tableRows.length) {
            throw new UserError(`Row index ${rowIndex} out of bounds. Table has ${tableRows.length} rows (0-${tableRows.length - 1}).`);
        }

        const row = tableRows[rowIndex];
        const tableCells = row.tableCells || [];

        // Validate column bounds
        if (columnIndex < 0 || columnIndex >= tableCells.length) {
            throw new UserError(`Column index ${columnIndex} out of bounds. Row has ${tableCells.length} columns (0-${tableCells.length - 1}).`);
        }

        const cell = tableCells[columnIndex];
        const cellStartIndex = cell.startIndex;
        const cellEndIndex = cell.endIndex;

        if (cellStartIndex === undefined || cellEndIndex === undefined) {
            console.warn(`Cell at (${rowIndex}, ${columnIndex}) missing index data`);
            return null;
        }

        // Calculate content range from paragraph elements in the cell
        // Each cell contains at least one paragraph with at least one newline
        let contentStartIndex = cellStartIndex;
        let contentEndIndex = cellStartIndex;

        if (cell.content && cell.content.length > 0) {
            // Get the first paragraph's first element start
            const firstPara = cell.content[0];
            if (firstPara.paragraph?.elements && firstPara.paragraph.elements.length > 0) {
                contentStartIndex = firstPara.paragraph.elements[0].startIndex || cellStartIndex;
            }

            // Get the last paragraph's last element end
            const lastPara = cell.content[cell.content.length - 1];
            if (lastPara.paragraph?.elements && lastPara.paragraph.elements.length > 0) {
                const elements = lastPara.paragraph.elements;
                const lastElement = elements[elements.length - 1];
                contentEndIndex = lastElement.endIndex || cellStartIndex;
            }
        }

        // Store paragraphEndIndex BEFORE subtracting (includes trailing newline for paragraph styling)
        // Google Docs paragraph styles CAN be applied to empty paragraphs if the range includes the newline
        const paragraphEndIndex = contentEndIndex;

        // The content end typically includes a trailing newline - preserve it by stopping one char before
        // If contentEndIndex > contentStartIndex, the actual editable content is contentStart to contentEnd - 1
        // But we need to handle the case where the cell only contains a newline
        if (contentEndIndex > contentStartIndex) {
            contentEndIndex = contentEndIndex - 1; // Exclude trailing newline
        }

        console.log(`Found cell (${rowIndex}, ${columnIndex}): cell=${cellStartIndex}-${cellEndIndex}, content=${contentStartIndex}-${contentEndIndex}, paragraphEnd=${paragraphEndIndex}`);

        return {
            cellStartIndex,
            cellEndIndex,
            contentStartIndex,
            contentEndIndex,
            paragraphEndIndex
        };

    } catch (error: any) {
        if (error instanceof UserError) throw error;
        console.error(`Error getting table cell range: ${error.message || 'Unknown error'}`);
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}).`);
        if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}).`);
        throw new Error(`Failed to get table cell range: ${error.message || 'Unknown error'}`);
    }
}

// --- Specific Feature Helpers ---

export async function createTable(docs: Docs, documentId: string, rows: number, columns: number, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (rows < 1 || columns < 1) {
        throw new UserError("Table must have at least 1 row and 1 column.");
    }
    const request: docs_v1.Schema$Request = {
insertTable: {
location: { index },
rows: rows,
columns: columns,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

export async function insertText(docs: Docs, documentId: string, text: string, index: number): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    if (!text) return {}; // Nothing to insert
    const request: docs_v1.Schema$Request = {
insertText: {
location: { index },
text: text,
}
};
return executeBatchUpdate(docs, documentId, [request]);
}

// --- Complex / Stubbed Helpers ---

export async function findParagraphsMatchingStyle(
docs: Docs,
documentId: string,
styleCriteria: any // Define a proper type for criteria (e.g., { fontFamily: 'Arial', bold: true })
): Promise<{ startIndex: number; endIndex: number }[]> {
// TODO: Implement logic
// 1. Get document content with paragraph elements and their styles.
// 2. Iterate through paragraphs.
// 3. For each paragraph, check if its computed style matches the criteria.
// 4. Return ranges of matching paragraphs.
console.warn("findParagraphsMatchingStyle is not implemented.");
throw new NotImplementedError("Finding paragraphs by style criteria is not yet implemented.");
// return [];
}

export async function detectAndFormatLists(
docs: Docs,
documentId: string,
startIndex?: number,
endIndex?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
// TODO: Implement complex logic
// 1. Get document content (paragraphs, text runs) in the specified range (or whole doc).
// 2. Iterate through paragraphs.
// 3. Identify sequences of paragraphs starting with list-like markers (e.g., "-", "*", "1.", "a)").
// 4. Determine nesting levels based on indentation or marker patterns.
// 5. Generate CreateParagraphBulletsRequests for the identified sequences.
// 6. Potentially delete the original marker text.
// 7. Execute the batch update.
console.warn("detectAndFormatLists is not implemented.");
throw new NotImplementedError("Automatic list detection and formatting is not yet implemented.");
// return {};
}

export async function addCommentHelper(docs: Docs, documentId: string, text: string, startIndex: number, endIndex: number): Promise<void> {
// NOTE: Adding comments typically requires the Google Drive API v3 and different scopes!
// 'https://www.googleapis.com/auth/drive' or more specific comment scopes.
// This helper is a placeholder assuming Drive API client (`drive`) is available and authorized.
/*
const drive = google.drive({version: 'v3', auth: authClient}); // Assuming authClient is available
await drive.comments.create({
fileId: documentId,
requestBody: {
content: text,
anchor: JSON.stringify({ // Anchor format might need verification
'type': 'workbook#textAnchor', // Or appropriate type for Docs
'refs': [{
'docRevisionId': 'head', // Or specific revision
'range': {
'start': startIndex,
'end': endIndex,
}
}]
})
},
fields: 'id'
});
*/
console.warn("addCommentHelper requires Google Drive API and is not implemented.");
throw new NotImplementedError("Adding comments requires Drive API setup and is not yet implemented.");
}

// --- Image Insertion Helpers ---

/**
 * Inserts an inline image into a document from a publicly accessible URL
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param imageUrl - Publicly accessible URL to the image
 * @param index - Position in the document where image should be inserted (1-based)
 * @param width - Optional width in points
 * @param height - Optional height in points
 * @returns Promise with batch update response
 */
export async function insertInlineImage(
    docs: Docs,
    documentId: string,
    imageUrl: string,
    index: number,
    width?: number,
    height?: number
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    // Validate URL format
    try {
        new URL(imageUrl);
    } catch (e) {
        throw new UserError(`Invalid image URL format: ${imageUrl}`);
    }

    // Build the insertInlineImage request
    const request: docs_v1.Schema$Request = {
        insertInlineImage: {
            location: { index },
            uri: imageUrl,
            ...(width && height && {
                objectSize: {
                    height: { magnitude: height, unit: 'PT' },
                    width: { magnitude: width, unit: 'PT' }
                }
            })
        }
    };

    return executeBatchUpdate(docs, documentId, [request]);
}

/**
 * Uploads a local image file to Google Drive and returns its public URL
 * @param drive - Google Drive API client
 * @param localFilePath - Path to the local image file
 * @param parentFolderId - Optional parent folder ID (defaults to root)
 * @returns Promise with the public webContentLink URL
 */
export async function uploadImageToDrive(
    drive: any, // drive_v3.Drive type
    localFilePath: string,
    parentFolderId?: string
): Promise<string> {
    const fs = await import('fs');
    const path = await import('path');

    // Verify file exists
    if (!fs.existsSync(localFilePath)) {
        throw new UserError(`Image file not found: ${localFilePath}`);
    }

    // Get file name and mime type
    const fileName = path.basename(localFilePath);
    const mimeTypeMap: { [key: string]: string } = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.webp': 'image/webp',
        '.svg': 'image/svg+xml'
    };

    const ext = path.extname(localFilePath).toLowerCase();
    const mimeType = mimeTypeMap[ext] || 'application/octet-stream';

    // Upload file to Drive
    const fileMetadata: any = {
        name: fileName,
        mimeType: mimeType
    };

    if (parentFolderId) {
        fileMetadata.parents = [parentFolderId];
    }

    const media = {
        mimeType: mimeType,
        body: fs.createReadStream(localFilePath)
    };

    const uploadResponse = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id,webViewLink,webContentLink'
    });

    const fileId = uploadResponse.data.id;
    if (!fileId) {
        throw new Error('Failed to upload image to Drive - no file ID returned');
    }

    // Make the file publicly readable
    await drive.permissions.create({
        fileId: fileId,
        requestBody: {
            role: 'reader',
            type: 'anyone'
        }
    });

    // Get the webContentLink
    const fileInfo = await drive.files.get({
        fileId: fileId,
        fields: 'webContentLink'
    });

    const webContentLink = fileInfo.data.webContentLink;
    if (!webContentLink) {
        throw new Error('Failed to get public URL for uploaded image');
    }

    return webContentLink;
}

// --- Tab Management Helpers ---

/**
 * Interface for a tab with hierarchy level information
 */
export interface TabWithLevel extends docs_v1.Schema$Tab {
    level: number;
}

/**
 * Recursively collect all tabs from a document in a flat list with hierarchy info
 * @param doc - The Google Doc document object
 * @returns Array of tabs with nesting level information
 */
export function getAllTabs(doc: docs_v1.Schema$Document): TabWithLevel[] {
    const allTabs: TabWithLevel[] = [];
    if (!doc.tabs || doc.tabs.length === 0) {
        return allTabs;
    }

    for (const tab of doc.tabs) {
        addCurrentAndChildTabs(tab, allTabs, 0);
    }
    return allTabs;
}

/**
 * Recursive helper to add tabs with their nesting level
 * @param tab - The tab to add
 * @param allTabs - The accumulator array
 * @param level - Current nesting level (0 for top-level)
 */
function addCurrentAndChildTabs(tab: docs_v1.Schema$Tab, allTabs: TabWithLevel[], level: number): void {
    allTabs.push({ ...tab, level });
    if (tab.childTabs && tab.childTabs.length > 0) {
        for (const childTab of tab.childTabs) {
            addCurrentAndChildTabs(childTab, allTabs, level + 1);
        }
    }
}

/**
 * Get the text length from a DocumentTab
 * @param documentTab - The DocumentTab object
 * @returns Total character count
 */
export function getTabTextLength(documentTab: docs_v1.Schema$DocumentTab | undefined): number {
    let totalLength = 0;

    if (!documentTab?.body?.content) {
        return 0;
    }

    documentTab.body.content.forEach((element: any) => {
        // Handle paragraphs
        if (element.paragraph?.elements) {
            element.paragraph.elements.forEach((pe: any) => {
                if (pe.textRun?.content) {
                    totalLength += pe.textRun.content.length;
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
                                totalLength += pe.textRun.content.length;
                            }
                        });
                    });
                });
            });
        }
    });

    return totalLength;
}

/**
 * Find a specific tab by ID in a document (searches recursively through child tabs)
 * @param doc - The Google Doc document object
 * @param tabId - The tab ID to search for
 * @returns The tab object if found, null otherwise
 */
export function findTabById(doc: docs_v1.Schema$Document, tabId: string): docs_v1.Schema$Tab | null {
    if (!doc.tabs || doc.tabs.length === 0) {
        return null;
    }

    // Helper function to search through tabs recursively
    const searchTabs = (tabs: docs_v1.Schema$Tab[]): docs_v1.Schema$Tab | null => {
        for (const tab of tabs) {
            if (tab.tabProperties?.tabId === tabId) {
                return tab;
            }
            // Recursively search child tabs
            if (tab.childTabs && tab.childTabs.length > 0) {
                const found = searchTabs(tab.childTabs);
                if (found) return found;
            }
        }
        return null;
    };

    return searchTabs(doc.tabs);
}

// --- Section Range Helper ---

const HEADING_STYLES: Record<string, number> = {
  'TITLE': 0,
  'SUBTITLE': 0,
  'HEADING_1': 1,
  'HEADING_2': 2,
  'HEADING_3': 3,
  'HEADING_4': 4,
  'HEADING_5': 5,
  'HEADING_6': 6,
};

function getHeadingLevel(namedStyleType: string | null | undefined): number | null {
  if (!namedStyleType || !(namedStyleType in HEADING_STYLES)) return null;
  return HEADING_STYLES[namedStyleType];
}

function getParagraphText(paragraph: docs_v1.Schema$Paragraph): string {
  let text = '';
  if (paragraph.elements) {
    for (const element of paragraph.elements) {
      if (element.textRun?.content) {
        text += element.textRun.content;
      }
    }
  }
  return text.replace(/\n$/, '');
}

export async function findSectionRange(
  docs: Docs,
  documentId: string,
  headingText: string,
): Promise<{ headingStart: number; headingEnd: number; sectionEnd: number } | null> {
  try {
    const res = await docs.documents.get({
      documentId,
      fields: 'body(content(paragraph(elements(textRun(content)),paragraphStyle(namedStyleType)),startIndex,endIndex))',
    });

    if (!res.data.body?.content) {
      console.warn(`No content found in document ${documentId}`);
      return null;
    }

    const content = res.data.body.content;
    let headingStart: number | null = null;
    let headingEnd: number | null = null;
    let headingLevel: number | null = null;

    // Step 1: Find the heading paragraph matching the text
    for (let i = 0; i < content.length; i++) {
      const element = content[i];
      if (!element.paragraph) continue;

      const style = element.paragraph.paragraphStyle?.namedStyleType;
      const level = getHeadingLevel(style);
      if (level === null) continue;

      const text = getParagraphText(element.paragraph);
      if (text.trim() === headingText.trim()) {
        headingStart = element.startIndex ?? null;
        headingEnd = element.endIndex ?? null;
        headingLevel = level;

        // Step 2: Scan forward to find section end
        let sectionEnd = headingEnd!;
        for (let j = i + 1; j < content.length; j++) {
          const nextElement = content[j];
          if (nextElement.paragraph) {
            const nextStyle = nextElement.paragraph.paragraphStyle?.namedStyleType;
            const nextLevel = getHeadingLevel(nextStyle);
            // Stop at same-or-higher level heading (lower number = higher level)
            if (nextLevel !== null && nextLevel <= headingLevel) {
              break;
            }
          }
          sectionEnd = nextElement.endIndex ?? sectionEnd;
        }

        if (headingStart === null || headingEnd === null) {
          console.warn(`Heading "${headingText}" found but missing index data`);
          return null;
        }

        console.log(`Found section "${headingText}" (level ${headingLevel}): heading=${headingStart}-${headingEnd}, sectionEnd=${sectionEnd}`);
        return { headingStart, headingEnd, sectionEnd };
      }
    }

    console.warn(`Heading "${headingText}" not found in document ${documentId}`);
    return null;
  } catch (error: unknown) {
    const err = error as { code?: number; message?: string };
    console.error(`Error finding section range for "${headingText}" in doc ${documentId}: ${err.message || 'Unknown error'}`);
    if (err.code === 404) throw new UserError(`Document not found (ID: ${documentId}).`);
    if (err.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}).`);
    throw new Error(`Failed to find section range: ${err.message || 'Unknown error'}`);
  }
}
