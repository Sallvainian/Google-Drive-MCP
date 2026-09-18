// src/googleDocsApiHelpers.ts
import { google, docs_v1 } from 'googleapis';
import { OAuth2Client } from 'google-auth-library';
import { UserError } from 'fastmcp';
import { TextStyleArgs, ParagraphStyleArgs, hexToRgbColor, NotImplementedError } from './types.js';

type Docs = docs_v1.Docs; // Alias for convenience

// --- Constants ---
const MAX_BATCH_UPDATE_REQUESTS = 50; // Google API limits batch size
export const FIND_TEXT_RANGE_FIELDS = 'body(content(paragraph(elements(startIndex,endIndex,textRun(content))),table,sectionBreak,tableOfContents,startIndex,endIndex))';
export const GET_PARAGRAPH_RANGE_FIELDS = 'body(content(startIndex,endIndex,paragraph,table,sectionBreak,tableOfContents))';
export const GET_TABLE_CELL_RANGE_FIELDS = 'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(paragraph(elements(startIndex,endIndex))))))))';
// Body-only (no tabs wrap). findElements sets SUGGESTIONS_VIEW_MODE on this get.
export const FIND_ELEMENT_FIELDS = 'body(content(startIndex,endIndex,table(rows,columns,tableRows(tableCells(content(paragraph(elements(startIndex,endIndex,textRun(content))))))),paragraph(elements(startIndex,endIndex,textRun(content)))))';

export const SUGGESTIONS_VIEW_MODE = 'PREVIEW_WITHOUT_SUGGESTIONS' as const;
export const TAB_ID_PROPERTIES = 'tabProperties(tabId)';
export const TAB_LIST_PROPERTIES = 'tabProperties(tabId,title,index,parentTabId)';

/**
 * Builds a tabs field mask with 3-deep childTabs so findTabById can see nested tabs.
 * @param documentTabFields - documentTab(...) subfields, e.g. 'documentTab(body(content(endIndex)))'
 */
export function buildTabsFieldMask(documentTabFields: string): string {
  const child3 = `childTabs(${TAB_ID_PROPERTIES},${documentTabFields})`;
  const child2 = `childTabs(${TAB_ID_PROPERTIES},${documentTabFields},${child3})`;
  const child1 = `childTabs(${TAB_ID_PROPERTIES},${documentTabFields},${child2})`;
  return `tabs(${TAB_ID_PROPERTIES},${documentTabFields},${child1})`;
}

export const TAB_BODY_RANGE_DOCUMENT_TAB_FIELDS = 'documentTab(body(content(startIndex,endIndex)))';
export const TAB_BODY_END_DOCUMENT_TAB_FIELDS = 'documentTab(body(content(endIndex)))';
export const TAB_VERIFY_DOCUMENT_TAB_FIELDS = TAB_BODY_END_DOCUMENT_TAB_FIELDS;
export const TAB_READ_DOCUMENT_TAB_FIELDS = 'documentTab(body,documentStyle,namedStyles,lists)';

export const TAB_BODY_RANGE_FIELDS = buildTabsFieldMask(TAB_BODY_RANGE_DOCUMENT_TAB_FIELDS);
export const TAB_BODY_END_INDEX_FIELDS = buildTabsFieldMask(TAB_BODY_END_DOCUMENT_TAB_FIELDS);
export const TAB_VERIFY_FIELDS = buildTabsFieldMask(TAB_VERIFY_DOCUMENT_TAB_FIELDS);
export const TAB_READ_CONTENT_FIELDS = `title,documentId,${buildTabsFieldMask(TAB_READ_DOCUMENT_TAB_FIELDS)}`;

const TAB_LIST_CHILD_FIELDS = `childTabs(${TAB_LIST_PROPERTIES},childTabs(${TAB_LIST_PROPERTIES},childTabs(${TAB_LIST_PROPERTIES})))`;
const TAB_LIST_WITH_CONTENT_CHILD_FIELDS = `childTabs(${TAB_LIST_PROPERTIES},${TAB_BODY_END_DOCUMENT_TAB_FIELDS},childTabs(${TAB_LIST_PROPERTIES},${TAB_BODY_END_DOCUMENT_TAB_FIELDS},childTabs(${TAB_LIST_PROPERTIES},${TAB_BODY_END_DOCUMENT_TAB_FIELDS})))`;

export const TAB_LIST_FIELDS = `title,tabs(${TAB_LIST_PROPERTIES},${TAB_LIST_CHILD_FIELDS})`;
export const TAB_LIST_WITH_CONTENT_FIELDS = `title,tabs(${TAB_LIST_PROPERTIES},${TAB_BODY_END_DOCUMENT_TAB_FIELDS},${TAB_LIST_WITH_CONTENT_CHILD_FIELDS})`;

export const TABLE_INDEX_BODY_FIELDS =
  'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex)))))';
export const TABLE_CONTENT_BASIC_BODY_FIELDS =
  'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(paragraph(elements(textRun(content)))))))))';
export const TABLE_CONTENT_INDEXED_BODY_FIELDS =
  'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(startIndex,endIndex,paragraph(elements(startIndex,endIndex,textRun(content)))))))))';
export const CLONE_TABLE_SOURCE_BODY_FIELDS =
  'body(content(startIndex,endIndex,table(rows,columns,tableStyle(tableColumnProperties(width,widthType)),tableRows(startIndex,endIndex,tableRowStyle(minRowHeight,preventOverflow,tableHeader),tableCells(startIndex,endIndex,tableCellStyle(backgroundColor,borderTop(color,width,dashStyle),borderBottom(color,width,dashStyle),borderLeft(color,width,dashStyle),borderRight(color,width,dashStyle),contentAlignment,paddingTop,paddingBottom,paddingLeft,paddingRight,rowSpan,columnSpan),content(paragraph(elements(startIndex,endIndex,textRun(content,textStyle(bold))))))))))';
export const CLONE_TABLE_TARGET_BODY_FIELDS =
  'body(content(startIndex,endIndex,table(rows,columns,tableRows(tableCells(startIndex,endIndex,content(paragraph(elements(startIndex,endIndex,textRun(content)))))))))';

/**
 * Returns bodyFields unchanged for a legacy (no tabId) documents.get, or wraps them
 * in a 3-deep tabs mask when tabId is set. Do not mix a root body(...) with tabs(...).
 */
export function buildDocumentGetFields(bodyFields: string, tabId?: string): string {
  return tabId ? buildTabsFieldMask(`documentTab(${bodyFields})`) : bodyFields;
}

// --- Core Helper to Execute Batch Updates ---
export async function executeBatchUpdate(docs: Docs, documentId: string, requests: docs_v1.Schema$Request[], writeControl?: docs_v1.Schema$WriteControl): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
  if (!requests || requests.length === 0) {
    return {}; // Nothing to do
  }

  // Split into chunks of MAX_BATCH_UPDATE_REQUESTS and execute sequentially
  // (order matters — document indices shift between batches)
  if (requests.length > MAX_BATCH_UPDATE_REQUESTS) {
    console.log(`Splitting ${requests.length} requests into batches of ${MAX_BATCH_UPDATE_REQUESTS}`);
    let combinedResponse: docs_v1.Schema$BatchUpdateDocumentResponse = {};
    let currentWriteControl = writeControl;
    let deleteCommitted = false;

    for (let i = 0; i < requests.length; i += MAX_BATCH_UPDATE_REQUESTS) {
      const chunk = requests.slice(i, i + MAX_BATCH_UPDATE_REQUESTS);
      console.log(`Executing batch ${Math.floor(i / MAX_BATCH_UPDATE_REQUESTS) + 1} (${chunk.length} requests)`);
      try {
        const response = await executeSingleBatch(docs, documentId, chunk, currentWriteControl);
        if (chunk.some((request) => request.deleteContentRange)) {
          deleteCommitted = true;
        }
        // Merge replies
        if (response.replies) {
          combinedResponse.replies = [...(combinedResponse.replies || []), ...response.replies];
        }
        combinedResponse.documentId = response.documentId;
        combinedResponse.writeControl = response.writeControl;
        if (response.writeControl?.requiredRevisionId) {
          currentWriteControl = { requiredRevisionId: response.writeControl.requiredRevisionId };
        }
      } catch (error: unknown) {
        const original = error instanceof Error ? error.message : String(error);
        if (deleteCommitted) {
          throw new UserError(`${original} Original content was already deleted and was not restored.`);
        }
        if (i > 0) {
          throw new UserError(`${original} Partial write: earlier batches were committed and leftover content was not undone.`);
        }
        throw error;
      }
    }

    return combinedResponse;
  }

  return executeSingleBatch(docs, documentId, requests, writeControl);
}

// Internal: execute a single batch (guaranteed <= MAX_BATCH_UPDATE_REQUESTS)
async function executeSingleBatch(
  docs: Docs,
  documentId: string,
  requests: docs_v1.Schema$Request[],
  writeControl?: docs_v1.Schema$WriteControl
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
    try {
        const response = await docs.documents.batchUpdate({
            documentId: documentId,
            requestBody: writeControl ? { requests, writeControl } : { requests },
        });
        return response.data;
    } catch (error: any) {
        console.error(`Google API batchUpdate Error for doc ${documentId}:`, error.response?.data || error.message);
        // Translate common API errors to UserErrors
        if (error.code === 400 && error.message.includes('Invalid requests')) {
             const details = error.response?.data?.error?.details;
             let detailMsg = '';
             if (details && Array.isArray(details)) {
                 detailMsg = details.map((d: any) => d.description || JSON.stringify(d)).join('; ');
             }
            throw new UserError(`Invalid request sent to Google Docs API. Details: ${detailMsg || error.message}`);
        }
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}). Check the ID.`);
        if (error.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}). Ensure the authenticated user has edit access.`);
        throw new UserError(`Google API Error (${error.code}): ${error.message}`);
    }
}

// --- Text Finding Helper ---
// This improved version is more robust in handling various text structure scenarios
export async function findTextRange(docs: Docs, documentId: string, textToFind: string, instance: number = 1, tabId?: string): Promise<{ startIndex: number; endIndex: number } | null> {
try {
    let bodyContent: docs_v1.Schema$StructuralElement[] | undefined;
    if (tabId) {
        const tab = await getDocumentTab(docs, documentId, tabId, `documentTab(${FIND_TEXT_RANGE_FIELDS})`);
        bodyContent = tab.documentTab?.body?.content;
    } else {
        const res = await docs.documents.get({
            documentId,
            fields: FIND_TEXT_RANGE_FIELDS,
        });
        bodyContent = res.data.body?.content;
    }

    if (!bodyContent) {
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

    collectTextFromContent(bodyContent);

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
    throw new UserError(`Failed to retrieve doc for text searching: ${error.message || 'Unknown error'}`);
}
}

// --- Element Finder ---
// textQuery: every non-overlapping occurrence's document index range (first tab / body only).
// elementType paragraph|table: top-level structural listing with a short preview.
export interface FoundElement {
    type: 'text' | 'paragraph' | 'table';
    startIndex: number;
    endIndex: number;
    instance?: number; // 1-based occurrence number, for text matches
    text?: string;     // matched text or a preview of the element
}

export async function findElements(
    docs: Docs,
    documentId: string,
    options: { textQuery?: string; elementType?: 'paragraph' | 'table' | 'list' | 'image' }
): Promise<FoundElement[]> {
    const { textQuery, elementType } = options;
    if (!textQuery && !elementType) {
        throw new UserError('findElement requires at least one of "textQuery" or "elementType".');
    }

    let res;
    try {
        res = await docs.documents.get({
            documentId,
            fields: FIND_ELEMENT_FIELDS,
            suggestionsViewMode: SUGGESTIONS_VIEW_MODE,
        });
    } catch (error: any) {
        if (error.code === 404) throw new UserError(`Document not found (ID: ${documentId}).`);
        if (error.code === 403) throw new UserError(`Permission denied for document ${documentId}.`);
        throw new Error(`Failed to retrieve document for findElement: ${error.message || 'Unknown error'}`);
    }

    const content = res.data.body?.content;
    if (!content) return [];

    const results: FoundElement[] = [];

    // --- Structural listing (paragraph / table) ---
    if (elementType === 'paragraph' || elementType === 'table') {
        for (const element of content) {
            if (elementType === 'paragraph' && element.paragraph?.elements) {
                const text = element.paragraph.elements
                    .map((pe: any) => pe.textRun?.content || '')
                    .join('');
                if (element.startIndex == null || element.endIndex == null) continue;
                results.push({
                    type: 'paragraph',
                    startIndex: element.startIndex,
                    endIndex: element.endIndex,
                    text: text.replace(/\n$/, '').slice(0, 120),
                });
            } else if (elementType === 'table' && element.table) {
                if (element.startIndex == null || element.endIndex == null) continue;
                results.push({
                    type: 'table',
                    startIndex: element.startIndex,
                    endIndex: element.endIndex,
                    text: `table ${element.table.rows ?? '?'}x${element.table.columns ?? '?'}`,
                });
            }
        }
        if (!textQuery) return results;
    } else if (elementType === 'list' || elementType === 'image') {
        // Reject even when textQuery is set so matches are not mislabeled as list/image.
        throw new UserError(`elementType "${elementType}" is not supported. Omit elementType and pass textQuery to locate content by text.`);
    }

    // Each paragraph (top-level or in a table cell) is one searchable unit. Concatenate
    // only contiguous text runs (startIndex === lastMappedIndex + 1) and map each
    // character back to its document index. A unit ends at a missing textRun.content
    // or an index gap, so a match cannot span an inline object. Paragraphs/cells are
    // hard boundaries. Nested tables are not searched (mask only populates top-level table).
    if (textQuery) {
        interface ParaUnit { text: string; map: number[]; firstIndex: number; }
        const units: ParaUnit[] = [];
        const collect = (items: any[]) => {
            items.forEach(element => {
                if (element.paragraph?.elements) {
                    let text = '';
                    let map: number[] = [];
                    const flush = () => {
                        if (map.length > 0) units.push({ text, map, firstIndex: map[0] });
                        text = '';
                        map = [];
                    };
                    element.paragraph.elements.forEach((pe: any) => {
                        const runContent = pe.textRun?.content;
                        if (runContent && pe.startIndex != null) {
                            if (map.length > 0 && pe.startIndex !== map[map.length - 1] + 1) {
                                flush();
                            }
                            for (let i = 0; i < runContent.length; i++) {
                                text += runContent[i];
                                map.push(pe.startIndex + i);
                            }
                        } else {
                            flush();
                        }
                    });
                    flush();
                }
                if (element.table?.tableRows) {
                    element.table.tableRows.forEach((row: any) => {
                        row.tableCells?.forEach((cell: any) => {
                            if (cell.content) collect(cell.content);
                        });
                    });
                }
            });
        };
        collect(content);
        units.sort((a, b) => a.firstIndex - b.firstIndex);

        let instance = 0;
        for (const unit of units) {
            let from = 0;
            while (true) {
                const at = unit.text.indexOf(textQuery, from);
                if (at === -1) break;
                instance++;
                results.push({
                    type: 'text',
                    instance,
                    startIndex: unit.map[at],
                    endIndex: unit.map[at + textQuery.length - 1] + 1,
                    text: textQuery,
                });
                from = at + textQuery.length;
            }
        }
    }

    return results;
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
        fields: GET_PARAGRAPH_RANGE_FIELDS,
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
    throw new UserError(`Failed to find paragraph: ${error.message || 'Unknown error'}`);
}
}

// --- Style Request Builders ---

export function buildUpdateTextStyleRequest(
startIndex: number,
endIndex: number,
style: TextStyleArgs,
tabId?: string
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

    const range: docs_v1.Schema$Range = { startIndex, endIndex };
    if (tabId) {
        range.tabId = tabId;
    }

    const request: docs_v1.Schema$Request = {
        updateTextStyle: {
            range,
            textStyle: textStyle,
            fields: fieldsToUpdate.join(','),
        }
    };
    return { request, fields: fieldsToUpdate };

}

export function buildUpdateParagraphStyleRequest(
startIndex: number,
endIndex: number,
style: ParagraphStyleArgs,
tabId?: string
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

    const range: docs_v1.Schema$Range = { startIndex, endIndex };
    if (tabId) {
        range.tabId = tabId;
    }

    // Build the request object
    const request: docs_v1.Schema$Request = {
        updateParagraphStyle: {
            range,
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
            fields: GET_TABLE_CELL_RANGE_FIELDS,
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
        throw new UserError(`Failed to get table cell range: ${error.message || 'Unknown error'}`);
    }
}

/**
 * Finds the startIndex of the first table whose cells contain the given search text.
 * @param docs - Google Docs API client
 * @param documentId - The document ID
 * @param searchText - Text to search for in any table cell
 * @returns The startIndex of the matching table, or null if not found
 */
export async function findTableStartIndexByText(
    docs: Docs,
    documentId: string,
    searchText: string
): Promise<number | null> {
    const res = await docs.documents.get({
        documentId,
        fields: 'body(content(startIndex,endIndex,table(tableRows(tableCells(content(paragraph(elements(textRun(content)))))))))',
    });

    if (!res.data.body?.content) return null;

    const lower = searchText.toLowerCase();
    for (const element of res.data.body.content) {
        if (!element.table || element.startIndex === undefined) continue;
        for (const row of element.table.tableRows || []) {
            for (const cell of row.tableCells || []) {
                let cellText = '';
                for (const para of cell.content || []) {
                    for (const el of para.paragraph?.elements || []) {
                        cellText += el.textRun?.content || '';
                    }
                }
                if (cellText.toLowerCase().includes(lower)) {
                    return element.startIndex;
                }
            }
        }
    }
    return null;
}

/**
 * Finds the cell BELOW a header cell matching the given text.
 * Returns the table's startIndex plus the row/column of the content cell (one row below the header).
 */
export async function findCellBelowHeader(
    docs: Docs,
    documentId: string,
    headerText: string
): Promise<{ tableStartIndex: number; rowIndex: number; columnIndex: number } | null> {
    const res = await docs.documents.get({
        documentId,
        fields: 'body(content(startIndex,endIndex,table(tableRows(tableCells(content(paragraph(elements(textRun(content)))))))))',
    });

    if (!res.data.body?.content) return null;

    const lower = headerText.toLowerCase();
    for (const element of res.data.body.content) {
        if (!element.table || element.startIndex === undefined) continue;
        const rows = element.table.tableRows || [];
        for (let r = 0; r < rows.length; r++) {
            const cells = rows[r].tableCells || [];
            for (let c = 0; c < cells.length; c++) {
                let cellText = '';
                for (const para of cells[c].content || []) {
                    for (const el of para.paragraph?.elements || []) {
                        cellText += el.textRun?.content || '';
                    }
                }
                if (cellText.toLowerCase().includes(lower)) {
                    // Return the cell one row below (same column)
                    const targetRow = r + 1;
                    if (targetRow < rows.length) {
                        return {
                            tableStartIndex: element.startIndex as number,
                            rowIndex: targetRow,
                            columnIndex: c,
                        };
                    }
                }
            }
        }
    }
    return null;
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
 * Uploads a local image file to Google Drive and grants anyone/reader so Docs can fetch it
 * @param drive - Google Drive API client
 * @param localFilePath - Path to the local image file
 * @param parentFolderId - Optional parent folder ID (defaults to root)
 * @returns Promise with fileId, webContentLink, and the anyone/reader permissionId
 */
export async function uploadImageToDrive(
    drive: any, // drive_v3.Drive type
    localFilePath: string,
    parentFolderId?: string
): Promise<{ fileId: string; webContentLink: string; permissionId: string }> {
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

    try {
        const uploadResponse = await drive.files.create({
            requestBody: fileMetadata,
            media: media,
            fields: 'id,webViewLink,webContentLink',
            supportsAllDrives: true
        });

        const fileId = uploadResponse.data.id;
        if (!fileId) {
            throw new UserError('Failed to upload image to Drive - no file ID returned');
        }

        // Make the file publicly readable
        const permissionResponse = await drive.permissions.create({
            fileId: fileId,
            requestBody: {
                role: 'reader',
                type: 'anyone'
            },
            fields: 'id',
            supportsAllDrives: true
        });

        const permissionId = permissionResponse.data.id;
        if (!permissionId) {
            throw new UserError('Failed to upload image to Drive - no permission ID returned');
        }

        try {
            const fileInfo = await drive.files.get({
                fileId: fileId,
                fields: 'webContentLink',
                supportsAllDrives: true
            });

            const webContentLink = fileInfo.data.webContentLink;
            if (!webContentLink) {
                throw new UserError('Failed to get public URL for uploaded image');
            }

            return { fileId, webContentLink, permissionId };
        } catch (getError) {
            await revokeAnyoneReaderGrant(drive, fileId, permissionId);
            throw getError;
        }
    } catch (error: any) {
        if (error instanceof UserError) throw error;
        if (error.code === 404) throw new UserError('Drive file not found while uploading image.');
        if (error.code === 403) throw new UserError('Permission denied while uploading image to Drive.');
        throw new UserError(`Failed to upload image to Drive: ${error.message || 'Unknown error'}`);
    }
}

/**
 * Revokes an anyone/reader grant on a Drive file.
 * @param drive - Google Drive API client
 * @param fileId - Drive file ID
 * @param permissionId - Permission ID returned by permissions.create
 */
export async function revokeAnyoneReaderGrant(
    drive: any,
    fileId: string,
    permissionId: string
): Promise<void> {
    try {
        await drive.permissions.delete({ fileId, permissionId, supportsAllDrives: true });
    } catch (error: any) {
        throw new UserError(`Failed to revoke the temporary anyone/reader grant on file ${fileId}; the anyone grant may still exist: ${error.message || 'Unknown error'}`);
    }
}

/**
 * Uploads a local image, inserts it into a document, then revokes the temporary anyone/reader grant.
 * On insert failure the grant is still revoked, then the insert error is rethrown.
 * @param docs - Google Docs API client
 * @param drive - Google Drive API client
 * @param localFilePath - Path to the local image file
 * @param documentId - The document ID
 * @param index - Position in the document where image should be inserted (1-based)
 * @param width - Optional width in points
 * @param height - Optional height in points
 * @param parentFolderId - Optional parent folder ID (defaults to root)
 * @returns Promise with fileId and webContentLink of the uploaded file
 */
export async function insertLocalImageFromPath(
    docs: Docs,
    drive: any,
    localFilePath: string,
    documentId: string,
    index: number,
    width?: number,
    height?: number,
    parentFolderId?: string
): Promise<{ fileId: string; webContentLink: string }> {
    const uploaded = await uploadImageToDrive(drive, localFilePath, parentFolderId);
    try {
        await insertInlineImage(
            docs,
            documentId,
            uploaded.webContentLink,
            index,
            width,
            height
        );
    } catch (insertError) {
        await revokeAnyoneReaderGrant(drive, uploaded.fileId, uploaded.permissionId);
        throw insertError;
    }
    await revokeAnyoneReaderGrant(drive, uploaded.fileId, uploaded.permissionId);
    return { fileId: uploaded.fileId, webContentLink: uploaded.webContentLink };
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

/**
 * Fetches a document with tab content and validates that tabId exists and has a
 * documentTab. Default mask is tabId + nested childTabs + documentTab presence.
 */
export async function getDocumentTab(
    docs: Docs,
    documentId: string,
    tabId: string,
    documentTabFields: string = TAB_VERIFY_DOCUMENT_TAB_FIELDS
): Promise<docs_v1.Schema$Tab> {
    const res = await docs.documents.get({
        documentId,
        includeTabsContent: true,
        suggestionsViewMode: SUGGESTIONS_VIEW_MODE,
        fields: buildTabsFieldMask(documentTabFields),
    });
    const tab = findTabById(res.data, tabId);
    if (!tab) {
        throw new UserError(`Tab with ID "${tabId}" not found in document.`);
    }
    if (!tab.documentTab) {
        throw new UserError(`Tab "${tabId}" does not have content (may not be a document tab).`);
    }
    return tab;
}

const ORDERED_GLYPH_TYPES = new Set([
    'DECIMAL',
    'ZERO_DECIMAL',
    'ALPHA',
    'UPPER_ALPHA',
    'ROMAN',
    'UPPER_ROMAN',
]);

/**
 * Assembles the convertDocsJsonToMarkdown input for a document tab, falling
 * back to the parent document lists map when the tab omits one.
 */
export function markdownContentSourceFromDocumentTab(
    documentTab: docs_v1.Schema$DocumentTab,
    documentLists?: docs_v1.Schema$Document['lists']
): { body?: docs_v1.Schema$Body | null; lists?: docs_v1.Schema$Document['lists'] } {
    return {
        body: documentTab.body,
        lists: documentTab.lists || documentLists,
    };
}

/**
 * Converts Google Docs JSON structure to Markdown format
 */
export function convertDocsJsonToMarkdown(docData: any): string {
    let markdown = '';

    if (!docData.body?.content) {
        return 'Document appears to be empty.';
    }

    const content = docData.body.content;
    for (let i = 0; i < content.length; i++) {
        const element = content[i];
        if (element.paragraph) {
            markdown += convertParagraphToMarkdown(element.paragraph, docData.lists);
            // A following non-list paragraph must not lazy-continue into this item.
            if (element.paragraph.bullet && !content[i + 1]?.paragraph?.bullet && !markdown.endsWith('\n\n')) {
                markdown += '\n';
            }
        } else if (element.table) {
            markdown += convertTableToMarkdown(element.table);
        } else if (element.sectionBreak) {
            markdown += '\n---\n\n'; // Section break as horizontal rule
        }
    }

    return markdown.trimEnd();
}

/**
 * Converts a paragraph element to markdown
 */
function convertParagraphToMarkdown(paragraph: any, lists?: any): string {
    let text = '';
    let isHeading = false;
    let headingLevel = 0;
    let isList = false;

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

    if (paragraph.bullet) {
        isList = true;
    }

    // Process text elements
    if (paragraph.elements) {
        paragraph.elements.forEach((element: any) => {
            if (element.textRun) {
                text += convertTextRunToMarkdown(element.textRun);
            }
        });
    }

    // Format based on style. List markers win over heading namedStyleType so a
    // DECIMAL item styled HEADING_1 stays `1. text`, not `# text`.
    if (isList) {
        const nestingLevel = paragraph.bullet.nestingLevel ?? 0;
        const indent = '  '.repeat(nestingLevel);
        const glyphType = lists?.[paragraph.bullet.listId]?.listProperties?.nestingLevels?.[nestingLevel]?.glyphType;
        const marker = ORDERED_GLYPH_TYPES.has(glyphType) ? '1.' : '-';
        return `${indent}${marker} ${text.trim()}\n`;
    } else if (isHeading && text.trim()) {
        const hashes = '#'.repeat(Math.min(headingLevel, 6));
        return `${hashes} ${text.trim()}\n\n`;
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
    const trailingNewlines = text.match(/\n+$/)?.[0] ?? '';
    const core = trailingNewlines ? text.slice(0, -trailingNewlines.length) : text;

    // A run that is only newlines is not wrapped
    if (!core) {
        return text;
    }

    text = core;

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

    return text + trailingNewlines;
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
            cellText = cellText.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');
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
): Promise<{ headingStart: number; headingEnd: number; sectionEnd: number; revisionId?: string | null } | null> {
  try {
    const res = await docs.documents.get({
      documentId,
      fields: 'revisionId,body(content(paragraph(elements(textRun(content)),paragraphStyle(namedStyleType)),startIndex,endIndex))',
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
        return { headingStart, headingEnd, sectionEnd, revisionId: res.data.revisionId };
      }
    }

    console.warn(`Heading "${headingText}" not found in document ${documentId}`);
    return null;
  } catch (error: unknown) {
    const err = error as { code?: number; message?: string };
    console.error(`Error finding section range for "${headingText}" in doc ${documentId}: ${err.message || 'Unknown error'}`);
    if (err.code === 404) throw new UserError(`Document not found (ID: ${documentId}).`);
    if (err.code === 403) throw new UserError(`Permission denied for document (ID: ${documentId}).`);
    throw new UserError(`Failed to find section range: ${err.message || 'Unknown error'}`);
  }
}

export type FormattedSection = {
  type: 'heading1' | 'heading2' | 'heading3' | 'heading4' | 'title' | 'subtitle' | 'normal' | 'bullet' | 'numbered';
  text: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
};

function assertFormattedContentNotEmpty(content: FormattedSection[]): void {
  if (!content || content.length === 0 || content.some((section) => !section.text)) {
    throw new UserError('content must not be empty.');
  }
}

export function buildFormattedContentRequests(
  sections: FormattedSection[],
  startingIndex: number
): { textRequests: docs_v1.Schema$Request[], styleRequests: docs_v1.Schema$Request[], finalIndex: number } {
  assertFormattedContentNotEmpty(sections);
  const textRequests: docs_v1.Schema$Request[] = [];
  const styleRequests: docs_v1.Schema$Request[] = [];
  let currentIndex = startingIndex;

  for (const section of sections) {
    const text = section.text + '\n';
    const textLength = text.length;
    const startIndex = currentIndex;
    const endIndex = currentIndex + textLength;

    textRequests.push({
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
          range: { startIndex, endIndex: endIndex - 1 },
          paragraphStyle: { namedStyleType },
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

    const textStyleFields: string[] = [];
    const textStyle: docs_v1.Schema$TextStyle = {};

    if (section.bold) { textStyle.bold = true; textStyleFields.push('bold'); }
    if (section.italic) { textStyle.italic = true; textStyleFields.push('italic'); }
    if (section.color) {
      const rgbColor = hexToRgbColor(section.color);
      if (rgbColor) {
        textStyle.foregroundColor = { color: { rgbColor } };
        textStyleFields.push('foregroundColor');
      }
    }

    if (textStyleFields.length > 0) {
      styleRequests.push({
        updateTextStyle: {
          range: { startIndex, endIndex: endIndex - 1 },
          textStyle,
          fields: textStyleFields.join(','),
        },
      });
    }

    currentIndex = endIndex;
  }

  return { textRequests, styleRequests, finalIndex: currentIndex };
}

function requireRevisionId(revisionId: string | null | undefined): string {
  if (!revisionId) {
    throw new UserError('Document revision ID was not returned; cannot pin this write.');
  }
  return revisionId;
}

export async function replaceFormattedDocumentContent(
  docs: Docs,
  documentId: string,
  content: FormattedSection[],
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
  assertFormattedContentNotEmpty(content);

  const docResponse = await docs.documents.get({
    documentId,
    fields: 'revisionId,body(content(endIndex))',
  });

  const revisionId = requireRevisionId(docResponse.data.revisionId);

  const lastElement = docResponse.data.body?.content?.[docResponse.data.body.content.length - 1];
  if (lastElement?.endIndex == null) {
    throw new UserError('Document endIndex was not returned; cannot replace content safely.');
  }
  const endIndex = lastElement.endIndex;

  const requests: docs_v1.Schema$Request[] = [];
  if (endIndex > 2) {
    requests.push({
      deleteContentRange: {
        range: { startIndex: 1, endIndex: endIndex - 1 },
      },
    });
  }

  const { textRequests, styleRequests } = buildFormattedContentRequests(content, 1);
  requests.push(...textRequests, ...styleRequests);

  return executeBatchUpdate(docs, documentId, requests, { requiredRevisionId: revisionId });
}

export async function updateFormattedDocumentSection(
  docs: Docs,
  documentId: string,
  headingText: string,
  content: FormattedSection[],
  replaceHeading: boolean = false,
): Promise<docs_v1.Schema$BatchUpdateDocumentResponse> {
  assertFormattedContentNotEmpty(content);

  const range = await findSectionRange(docs, documentId, headingText);
  if (!range) {
    throw new UserError(`Could not find a heading matching "${headingText}" in the document. Make sure the heading text is an exact match.`);
  }

  const revisionId = requireRevisionId(range.revisionId);
  const deleteStart = replaceHeading ? range.headingStart : range.headingEnd;
  const deleteEnd = range.sectionEnd;

  const requests: docs_v1.Schema$Request[] = [];
  if (deleteEnd > deleteStart) {
    requests.push({
      deleteContentRange: {
        range: { startIndex: deleteStart, endIndex: deleteEnd - 1 },
      },
    });
  }

  const { textRequests, styleRequests } = buildFormattedContentRequests(content, deleteStart);
  requests.push(...textRequests, ...styleRequests);

  return executeBatchUpdate(docs, documentId, requests, { requiredRevisionId: revisionId });
}

export async function createFormattedDocument(
  drive: any,
  docs: Docs,
  args: { title: string; content: FormattedSection[]; parentFolderId?: string }
): Promise<{ id?: string | null; name?: string | null; webViewLink?: string | null }> {
  assertFormattedContentNotEmpty(args.content);
  const { textRequests, styleRequests } = buildFormattedContentRequests(args.content, 1);

  const documentMetadata: any = {
    name: args.title,
    mimeType: 'application/vnd.google-apps.document',
  };
  if (args.parentFolderId) {
    documentMetadata.parents = [args.parentFolderId];
  }

  const createResponse = await drive.files.create({
    requestBody: documentMetadata,
    fields: 'id,name,webViewLink',
    supportsAllDrives: true,
  });

  const document = createResponse.data;
  const documentId = document.id;
  if (!documentId) {
    throw new UserError('Failed to create formatted document - no file ID returned.');
  }

  try {
    await executeBatchUpdate(docs, documentId, [...textRequests, ...styleRequests]);
  } catch (error: unknown) {
    try {
      await drive.files.delete({
        fileId: documentId,
        supportsAllDrives: true,
      });
    } catch {
      const original = error instanceof Error ? error.message : String(error);
      throw new UserError(`${original} Leftover empty document was not deleted (ID: ${documentId}).`);
    }
    throw error;
  }

  return document;
}

export interface ExtractedTableCell {
  rowIndex: number;
  columnIndex: number;
  startIndex: number | null;
  endIndex: number | null;
  contentStartIndex: number | null;
  contentEndIndex: number | null;
  text: string;
}

export interface ExtractedTable {
  tableId: string;
  ordinal: number;
  startIndex: number | null;
  endIndex: number | null;
  rowCount: number;
  columnCount: number;
  cells: ExtractedTableCell[];
}

export interface ExtractedTableColumnStyle {
  columnIndex: number;
  widthPt?: number;
  widthType?: string | null;
}

export interface ExtractedTableRowStyle {
  rowIndex: number;
  minRowHeightPt?: number;
  preventOverflow?: boolean;
  tableHeader?: boolean;
}

export interface ExtractedTableCellStyle {
  rowIndex: number;
  columnIndex: number;
  backgroundColor?: docs_v1.Schema$RgbColor;
  contentAlignment?: 'CONTENT_ALIGNMENT_UNSPECIFIED' | 'TOP' | 'MIDDLE' | 'BOTTOM' | null;
  paddingTopPt?: number;
  paddingBottomPt?: number;
  paddingLeftPt?: number;
  paddingRightPt?: number;
  borderTop?: docs_v1.Schema$TableCellBorder;
  borderBottom?: docs_v1.Schema$TableCellBorder;
  borderLeft?: docs_v1.Schema$TableCellBorder;
  borderRight?: docs_v1.Schema$TableCellBorder;
  hasBoldText?: boolean;
}

export interface ExtractedTableSnapshot {
  tableId: string;
  startIndex: number | null;
  endIndex: number | null;
  rowCount: number;
  columnCount: number;
  data: string[][];
  columnStyles: ExtractedTableColumnStyle[];
  rowStyles: ExtractedTableRowStyle[];
  cellStyles: ExtractedTableCellStyle[];
  pinnedHeaderRowsCount: number;
}

function getTableContentSource(
  doc: docs_v1.Schema$Document,
  tabId?: string
): docs_v1.Schema$StructuralElement[] {
  if (tabId) {
    const targetTab = findTabById(doc, tabId);
    if (!targetTab?.documentTab?.body?.content) {
      return [];
    }
    return targetTab.documentTab.body.content;
  }
  return doc.body?.content ?? [];
}

function extractParagraphText(paragraph?: docs_v1.Schema$Paragraph): string {
  return (
    paragraph?.elements
      ?.map((element) => element.textRun?.content ?? '')
      .join('')
      .replace(/\n+$/g, '') ?? ''
  );
}

function extractCellText(content: docs_v1.Schema$StructuralElement[] = []): string {
  const parts: string[] = [];
  for (const element of content) {
    if (element.paragraph) {
      const text = extractParagraphText(element.paragraph);
      if (text) parts.push(text);
    }
    if (element.table?.tableRows) {
      for (const row of element.table.tableRows) {
        for (const cell of row.tableCells ?? []) {
          const text = extractCellText(cell.content ?? []);
          if (text) parts.push(text);
        }
      }
    }
  }
  return parts.join('\n').trim();
}

function extractCellContentRange(content: docs_v1.Schema$StructuralElement[] = []): {
  contentStartIndex: number | null;
  contentEndIndex: number | null;
} {
  let minStart: number | null = null;
  let maxEnd: number | null = null;
  const visitContent = (elements: docs_v1.Schema$StructuralElement[]) => {
    for (const element of elements) {
      for (const paragraphElement of element.paragraph?.elements ?? []) {
        const startIndex = paragraphElement.startIndex;
        if (typeof startIndex === 'number') {
          minStart = minStart === null ? startIndex : Math.min(minStart, startIndex);
        }
        const endIndex = paragraphElement.endIndex;
        if (typeof endIndex === 'number') {
          maxEnd = maxEnd === null ? endIndex : Math.max(maxEnd, endIndex);
        }
      }
      if (element.table?.tableRows) {
        for (const row of element.table.tableRows) {
          for (const cell of row.tableCells ?? []) {
            visitContent(cell.content ?? []);
          }
        }
      }
    }
  };
  visitContent(content);
  return { contentStartIndex: minStart, contentEndIndex: maxEnd };
}

function dimensionToPt(dimension?: docs_v1.Schema$Dimension): number | undefined {
  if (!dimension?.magnitude || dimension.unit !== 'PT') return undefined;
  return dimension.magnitude;
}

function normalizeCellStyle(
  rowIndex: number,
  columnIndex: number,
  cell: docs_v1.Schema$TableCell
): ExtractedTableCellStyle | null {
  const style = cell.tableCellStyle;
  const firstParagraphHasBoldText = (cell.content ?? []).some((element) =>
    (element.paragraph?.elements ?? []).some(
      (paragraphElement) => paragraphElement.textRun?.textStyle?.bold
    )
  );
  if (!style && !firstParagraphHasBoldText) return null;
  const contentAlignment =
    style?.contentAlignment === 'TOP' ||
    style?.contentAlignment === 'MIDDLE' ||
    style?.contentAlignment === 'BOTTOM' ||
    style?.contentAlignment === 'CONTENT_ALIGNMENT_UNSPECIFIED'
      ? style.contentAlignment
      : null;
  return {
    rowIndex,
    columnIndex,
    backgroundColor: style?.backgroundColor?.color?.rgbColor ?? undefined,
    contentAlignment,
    paddingTopPt: dimensionToPt(style?.paddingTop),
    paddingBottomPt: dimensionToPt(style?.paddingBottom),
    paddingLeftPt: dimensionToPt(style?.paddingLeft),
    paddingRightPt: dimensionToPt(style?.paddingRight),
    borderTop: style?.borderTop ?? undefined,
    borderBottom: style?.borderBottom ?? undefined,
    borderLeft: style?.borderLeft ?? undefined,
    borderRight: style?.borderRight ?? undefined,
    hasBoldText: firstParagraphHasBoldText || undefined,
  };
}

export function extractDocumentTables(
  doc: docs_v1.Schema$Document,
  tabId?: string
): ExtractedTable[] {
  const content = getTableContentSource(doc, tabId);
  const tables: ExtractedTable[] = [];
  const tabKey = tabId ?? 'body';
  for (const element of content) {
    if (!element.table?.tableRows) continue;
    const ordinal = tables.length;
    const cells: ExtractedTableCell[] = [];
    let columnCount = 0;
    element.table.tableRows.forEach((row, rowIndex) => {
      const rowCells = row.tableCells ?? [];
      columnCount = Math.max(columnCount, rowCells.length);
      rowCells.forEach((cell, columnIndex) => {
        const { contentStartIndex, contentEndIndex } = extractCellContentRange(cell.content ?? []);
        cells.push({
          rowIndex,
          columnIndex,
          startIndex: cell.startIndex ?? null,
          endIndex: cell.endIndex ?? null,
          contentStartIndex,
          contentEndIndex,
          text: extractCellText(cell.content ?? []),
        });
      });
    });
    tables.push({
      tableId: `table:${tabKey}:${ordinal}`,
      ordinal,
      startIndex: element.startIndex ?? null,
      endIndex: element.endIndex ?? null,
      rowCount: element.table.tableRows.length,
      columnCount,
      cells,
    });
  }
  return tables;
}

export function getTableById(
  doc: docs_v1.Schema$Document,
  tableId: string,
  tabId?: string
): ExtractedTable | null {
  return extractDocumentTables(doc, tabId).find((table) => table.tableId === tableId) ?? null;
}

export function extractTableSnapshot(
  doc: docs_v1.Schema$Document,
  tableId: string,
  tabId?: string
): ExtractedTableSnapshot | null {
  const content = getTableContentSource(doc, tabId);
  const tabKey = tabId ?? 'body';
  let ordinal = 0;
  for (const element of content) {
    if (!element.table?.tableRows) continue;
    const currentTableId = `table:${tabKey}:${ordinal}`;
    ordinal++;
    if (currentTableId !== tableId) continue;
    const data: string[][] = [];
    const rowStyles: ExtractedTableRowStyle[] = [];
    const cellStyles: ExtractedTableCellStyle[] = [];
    let pinnedHeaderRowsCount = 0;
    element.table.tableRows.forEach((row, rowIndex) => {
      const rowData: string[] = [];
      const rowStyle = row.tableRowStyle;
      if (rowStyle) {
        rowStyles.push({
          rowIndex,
          minRowHeightPt: dimensionToPt(rowStyle.minRowHeight),
          preventOverflow: rowStyle.preventOverflow ?? undefined,
          tableHeader: rowStyle.tableHeader ?? undefined,
        });
      }
      if ((rowStyle?.tableHeader ?? false) && pinnedHeaderRowsCount === rowIndex) {
        pinnedHeaderRowsCount++;
      }
      (row.tableCells ?? []).forEach((cell, columnIndex) => {
        rowData.push(extractCellText(cell.content ?? []));
        const cellStyle = normalizeCellStyle(rowIndex, columnIndex, cell);
        if (cellStyle) cellStyles.push(cellStyle);
      });
      data.push(rowData);
    });
    const columnStyles: ExtractedTableColumnStyle[] =
      element.table.tableStyle?.tableColumnProperties?.map((column, columnIndex) => ({
        columnIndex,
        widthPt: dimensionToPt(column.width),
        widthType: column.widthType,
      })) ?? [];
    return {
      tableId: currentTableId,
      startIndex: element.startIndex ?? null,
      endIndex: element.endIndex ?? null,
      rowCount: element.table.rows ?? data.length,
      columnCount: element.table.columns ?? Math.max(...data.map((row) => row.length), 0),
      data,
      columnStyles,
      rowStyles,
      cellStyles,
      pinnedHeaderRowsCount,
    };
  }
  return null;
}

export function buildReplaceTableCellContentRequests(
  cell: ExtractedTable['cells'][number],
  nextValue: string,
  tabId?: string
): docs_v1.Schema$Request[] {
  const requests: docs_v1.Schema$Request[] = [];
  const insertionIndex = cell.contentStartIndex ?? cell.startIndex;
  if (insertionIndex == null) {
    throw new UserError(
      `Cell [row=${cell.rowIndex}, col=${cell.columnIndex}] does not have a writable insertion index.`
    );
  }
  if (
    cell.text &&
    cell.contentStartIndex !== null &&
    cell.contentEndIndex !== null &&
    cell.contentEndIndex - 1 > cell.contentStartIndex
  ) {
    const range: docs_v1.Schema$Range = {
      startIndex: cell.contentStartIndex,
      endIndex: cell.contentEndIndex - 1,
    };
    if (tabId) range.tabId = tabId;
    requests.push({ deleteContentRange: { range } });
  }
  if (nextValue) {
    const location: docs_v1.Schema$Location = { index: insertionIndex };
    if (tabId) location.tabId = tabId;
    requests.push({ insertText: { location, text: nextValue } });
  }
  return requests;
}

export function buildReplaceTableRowRequests(
  table: ExtractedTable,
  rowIndex: number,
  values: string[],
  tabId?: string
): docs_v1.Schema$Request[] {
  if (rowIndex < 0 || rowIndex >= table.rowCount) {
    throw new UserError(
      `Row index ${rowIndex} is out of bounds for table ${table.tableId} with ${table.rowCount} rows.`
    );
  }
  if (values.length > table.columnCount) {
    throw new UserError(
      `Received ${values.length} values for table ${table.tableId}, but the table only has ${table.columnCount} columns.`
    );
  }
  return table.cells
    .filter((cell) => cell.rowIndex === rowIndex)
    .sort((a, b) => b.columnIndex - a.columnIndex)
    .flatMap((cell) =>
      buildReplaceTableCellContentRequests(cell, values[cell.columnIndex] ?? '', tabId)
    );
}

export async function replaceTableRowData(
  docs: Docs,
  documentId: string,
  table: ExtractedTable,
  rowIndex: number,
  values: string[],
  tabId?: string
): Promise<void> {
  const requests = buildReplaceTableRowRequests(table, rowIndex, values, tabId);
  if (requests.length === 0) return;
  await executeBatchUpdate(docs, documentId, requests);
}

export function buildInsertTableWithDataRequests(
  data: string[][],
  index: number,
  hasHeaderRow: boolean,
  tabId?: string
): docs_v1.Schema$Request[] {
  const numRows = data.length;
  const numCols = data.reduce((max, row) => Math.max(max, row.length), 0);
  if (numRows === 0 || numCols === 0) {
    throw new UserError(
      'Table data must contain at least one non-empty row with at least one cell.'
    );
  }
  const normalizedData = data.map((row) => {
    const padded = [...row];
    while (padded.length < numCols) padded.push('');
    return padded;
  });
  const insertRequests: docs_v1.Schema$Request[] = [];
  const formatRequests: docs_v1.Schema$Request[] = [];
  const location: docs_v1.Schema$Location = { index };
  if (tabId) location.tabId = tabId;
  insertRequests.push({
    insertTable: {
      location,
      rows: numRows,
      columns: numCols,
    },
  });
  let cumulativeTextLength = 0;
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cellText = normalizedData[r][c];
      if (!cellText) continue;
      const baseCellIndex = index + 4 + r * (1 + 2 * numCols) + 2 * c;
      const adjustedIndex = baseCellIndex + cumulativeTextLength;
      const cellLocation: docs_v1.Schema$Location = { index: adjustedIndex };
      if (tabId) cellLocation.tabId = tabId;
      insertRequests.push({
        insertText: {
          location: cellLocation,
          text: cellText,
        },
      });
      if (hasHeaderRow && r === 0) {
        const styleReq = buildUpdateTextStyleRequest(
          adjustedIndex,
          adjustedIndex + cellText.length,
          { bold: true },
          tabId
        );
        if (styleReq) formatRequests.push(styleReq.request);
      }
      cumulativeTextLength += cellText.length;
    }
  }
  return [...insertRequests, ...formatRequests];
}

export function buildInsertSectionBreakRequest(params: {
  index: number;
  sectionType: 'NEXT_PAGE' | 'CONTINUOUS';
  tabId?: string;
}): docs_v1.Schema$Request {
  const location: docs_v1.Schema$Location = { index: params.index };
  if (params.tabId) {
    location.tabId = params.tabId;
  }
  return {
    insertSectionBreak: {
      location,
      sectionType: params.sectionType,
    },
  };
}

export interface UpdateSectionStyleBuilderInput {
  startIndex: number;
  endIndex: number;
  flipPageOrientation?: boolean;
  sectionType?: 'SECTION_TYPE_UNSPECIFIED' | 'CONTINUOUS' | 'NEXT_PAGE';
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
  pageNumberStart?: number;
  tabId?: string;
}

export function buildUpdateSectionStyleRequest(
  params: UpdateSectionStyleBuilderInput
): { request: docs_v1.Schema$Request; fields: string[] } | null {
  const sectionStyle: docs_v1.Schema$SectionStyle = {};
  const fields: string[] = [];
  if (params.flipPageOrientation !== undefined) {
    sectionStyle.flipPageOrientation = params.flipPageOrientation;
    fields.push('flipPageOrientation');
  }
  if (params.sectionType !== undefined) {
    sectionStyle.sectionType = params.sectionType;
    fields.push('sectionType');
  }
  if (params.marginTop !== undefined) {
    sectionStyle.marginTop = { magnitude: params.marginTop, unit: 'PT' };
    fields.push('marginTop');
  }
  if (params.marginBottom !== undefined) {
    sectionStyle.marginBottom = { magnitude: params.marginBottom, unit: 'PT' };
    fields.push('marginBottom');
  }
  if (params.marginLeft !== undefined) {
    sectionStyle.marginLeft = { magnitude: params.marginLeft, unit: 'PT' };
    fields.push('marginLeft');
  }
  if (params.marginRight !== undefined) {
    sectionStyle.marginRight = { magnitude: params.marginRight, unit: 'PT' };
    fields.push('marginRight');
  }
  if (params.pageNumberStart !== undefined) {
    sectionStyle.pageNumberStart = params.pageNumberStart;
    fields.push('pageNumberStart');
  }
  if (fields.length === 0) {
    return null;
  }
  const range: docs_v1.Schema$Range = {
    startIndex: params.startIndex,
    endIndex: params.endIndex,
  };
  if (params.tabId) {
    range.tabId = params.tabId;
  }
  return {
    request: {
      updateSectionStyle: {
        range,
        sectionStyle,
        fields: fields.join(','),
      },
    },
    fields,
  };
}

export interface BuildModifyTextOpts {
  startIndex: number;
  endIndex?: number;
  text?: string;
  style?: TextStyleArgs;
  tabId?: string;
}

export function buildModifyTextRequests(opts: BuildModifyTextOpts): docs_v1.Schema$Request[] {
  const { startIndex, endIndex, text, style, tabId } = opts;
  const requests: docs_v1.Schema$Request[] = [];
  if (text === undefined && !style) return requests;
  if (endIndex !== undefined && text !== undefined) {
    const range: docs_v1.Schema$Range = { startIndex, endIndex };
    if (tabId) range.tabId = tabId;
    requests.push({ deleteContentRange: { range } });
  }
  if (text !== undefined) {
    const location: docs_v1.Schema$Location = { index: startIndex };
    if (tabId) location.tabId = tabId;
    requests.push({ insertText: { location, text } });
  }
  if (style) {
    const formatStart = startIndex;
    const formatEnd =
      text !== undefined
        ? startIndex + text.length
        : endIndex !== undefined
          ? endIndex
          : startIndex;
    if (formatEnd > formatStart) {
      const requestInfo = buildUpdateTextStyleRequest(formatStart, formatEnd, style, tabId);
      if (requestInfo) {
        requests.push(requestInfo.request);
      }
    }
  }
  return requests;
}

export function buildTableStartLocation(
  tableStartIndex: number,
  tabId?: string
): docs_v1.Schema$Location {
  const location: docs_v1.Schema$Location = { index: tableStartIndex };
  if (tabId) {
    location.tabId = tabId;
  }
  return location;
}

type TableCellStyleArgs = {
  backgroundColor?: docs_v1.Schema$RgbColor;
  contentAlignment?: 'CONTENT_ALIGNMENT_UNSPECIFIED' | 'TOP' | 'MIDDLE' | 'BOTTOM';
  rowSpan?: number;
  columnSpan?: number;
  paddingTopPt?: number;
  paddingBottomPt?: number;
  paddingLeftPt?: number;
  paddingRightPt?: number;
  borderTop?: docs_v1.Schema$TableCellBorder;
  borderBottom?: docs_v1.Schema$TableCellBorder;
  borderLeft?: docs_v1.Schema$TableCellBorder;
  borderRight?: docs_v1.Schema$TableCellBorder;
};

function pointDimension(magnitude: number): docs_v1.Schema$Dimension {
  return { magnitude, unit: 'PT' };
}

export function buildTableCellStyleRequest(
  tableStartIndex: number,
  rowIndex: number,
  columnIndex: number,
  style: TableCellStyleArgs,
  tabId?: string
): { request: docs_v1.Schema$Request; fields: string[] } | null {
  const tableCellStyle: docs_v1.Schema$TableCellStyle = {};
  const fields: string[] = [];
  if (style.backgroundColor) {
    tableCellStyle.backgroundColor = { color: { rgbColor: style.backgroundColor } };
    fields.push('backgroundColor');
  }
  if (style.contentAlignment) {
    tableCellStyle.contentAlignment = style.contentAlignment;
    fields.push('contentAlignment');
  }
  if (style.paddingTopPt !== undefined) {
    tableCellStyle.paddingTop = pointDimension(style.paddingTopPt);
    fields.push('paddingTop');
  }
  if (style.paddingBottomPt !== undefined) {
    tableCellStyle.paddingBottom = pointDimension(style.paddingBottomPt);
    fields.push('paddingBottom');
  }
  if (style.paddingLeftPt !== undefined) {
    tableCellStyle.paddingLeft = pointDimension(style.paddingLeftPt);
    fields.push('paddingLeft');
  }
  if (style.paddingRightPt !== undefined) {
    tableCellStyle.paddingRight = pointDimension(style.paddingRightPt);
    fields.push('paddingRight');
  }
  if (style.borderTop) {
    tableCellStyle.borderTop = style.borderTop;
    fields.push('borderTop');
  }
  if (style.borderBottom) {
    tableCellStyle.borderBottom = style.borderBottom;
    fields.push('borderBottom');
  }
  if (style.borderLeft) {
    tableCellStyle.borderLeft = style.borderLeft;
    fields.push('borderLeft');
  }
  if (style.borderRight) {
    tableCellStyle.borderRight = style.borderRight;
    fields.push('borderRight');
  }
  if (fields.length === 0) return null;
  const rowSpan = style.rowSpan ?? 1;
  const columnSpan = style.columnSpan ?? 1;
  return {
    request: {
      updateTableCellStyle: {
        tableRange: {
          tableCellLocation: {
            tableStartLocation: buildTableStartLocation(tableStartIndex, tabId),
            rowIndex,
            columnIndex,
          },
          rowSpan,
          columnSpan,
        },
        tableCellStyle,
        fields: fields.join(','),
      },
    },
    fields,
  };
}

export function buildTableBorder(
  color: docs_v1.Schema$RgbColor,
  widthPt: number,
  dashStyle: 'SOLID' | 'DASHED' | 'DOTTED'
): docs_v1.Schema$TableCellBorder {
  return {
    color: { color: { rgbColor: color } },
    width: pointDimension(widthPt),
    dashStyle,
  };
}

export function buildTableColumnWidthRequest(
  tableStartIndex: number,
  columnIndices: number[],
  widthPt: number,
  tabId?: string
): docs_v1.Schema$Request {
  return {
    updateTableColumnProperties: {
      tableStartLocation: buildTableStartLocation(tableStartIndex, tabId),
      columnIndices,
      tableColumnProperties: {
        widthType: 'FIXED_WIDTH',
        width: pointDimension(widthPt),
      },
      fields: 'widthType,width',
    },
  };
}

export function buildTableRowStyleRequest(
  tableStartIndex: number,
  rowIndices: number[],
  minRowHeightPt: number | undefined,
  preventOverflow: boolean | undefined,
  tabId?: string
): docs_v1.Schema$Request | null {
  const tableRowStyle: docs_v1.Schema$TableRowStyle = {};
  const fields: string[] = [];
  if (minRowHeightPt !== undefined) {
    tableRowStyle.minRowHeight = pointDimension(minRowHeightPt);
    fields.push('minRowHeight');
  }
  if (preventOverflow !== undefined) {
    tableRowStyle.preventOverflow = preventOverflow;
    fields.push('preventOverflow');
  }
  if (fields.length === 0) return null;
  return {
    updateTableRowStyle: {
      tableStartLocation: buildTableStartLocation(tableStartIndex, tabId),
      rowIndices,
      tableRowStyle,
      fields: fields.join(','),
    },
  };
}

export function buildPinTableHeaderRowsRequest(
  tableStartIndex: number,
  pinnedHeaderRowsCount: number,
  tabId?: string
): docs_v1.Schema$Request {
  return {
    pinTableHeaderRows: {
      tableStartLocation: buildTableStartLocation(tableStartIndex, tabId),
      pinnedHeaderRowsCount,
    },
  };
}
