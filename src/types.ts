// src/types.ts
import { z } from 'zod';
import { docs_v1 } from 'googleapis';

// --- Helper function for hex color validation ---
export const hexColorRegex = /^#?([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$/;
export const validateHexColor = (color: string) => hexColorRegex.test(color);

// --- Helper function for Hex to RGB conversion ---
export function hexToRgbColor(hex: string): docs_v1.Schema$RgbColor | null {
if (!hex) return null;
let hexClean = hex.startsWith('#') ? hex.slice(1) : hex;

if (hexClean.length === 3) {
hexClean = hexClean[0] + hexClean[0] + hexClean[1] + hexClean[1] + hexClean[2] + hexClean[2];
}
if (hexClean.length !== 6) return null;
const bigint = parseInt(hexClean, 16);
if (isNaN(bigint)) return null;

const r = ((bigint >> 16) & 255) / 255;
const g = ((bigint >> 8) & 255) / 255;
const b = (bigint & 255) / 255;

return { red: r, green: g, blue: b };
}

// --- Zod Schema Fragments for Reusability ---

export const DocumentIdParameter = z.object({
documentId: z.string().describe('The ID of the Google Document (from the URL).'),
});

export const RangeParameters = z.object({
startIndex: z.number().int().min(1).describe('The starting index of the text range (inclusive, starts from 1).'),
endIndex: z.number().int().min(1).describe('The ending index of the text range (exclusive).'),
}).refine(data => data.endIndex > data.startIndex, {
message: "endIndex must be greater than startIndex",
path: ["endIndex"],
});

export const OptionalRangeParameters = z.object({
startIndex: z.number().int().min(1).optional().describe('Optional: The starting index of the text range (inclusive, starts from 1). If omitted, might apply to a found element or whole paragraph.'),
endIndex: z.number().int().min(1).optional().describe('Optional: The ending index of the text range (exclusive). If omitted, might apply to a found element or whole paragraph.'),
}).refine(data => !data.startIndex || !data.endIndex || data.endIndex > data.startIndex, {
message: "If both startIndex and endIndex are provided, endIndex must be greater than startIndex",
path: ["endIndex"],
});

export const TextFindParameter = z.object({
textToFind: z.string().min(1).describe('The exact text string to locate.'),
matchInstance: z.number().int().min(1).optional().default(1).describe('Which instance of the text to target (1st, 2nd, etc.). Defaults to 1.'),
});

// --- Style Parameter Schemas ---

export const TextStyleParameters = z.object({
bold: z.boolean().optional().describe('Apply bold formatting.'),
italic: z.boolean().optional().describe('Apply italic formatting.'),
underline: z.boolean().optional().describe('Apply underline formatting.'),
strikethrough: z.boolean().optional().describe('Apply strikethrough formatting.'),
fontSize: z.number().min(1).optional().describe('Set font size (in points, e.g., 12).'),
fontFamily: z.string().optional().describe('Set font family (e.g., "Arial", "Times New Roman").'),
foregroundColor: z.string()
.refine(validateHexColor, { message: "Invalid hex color format (e.g., #FF0000 or #F00)" })
.optional()
.describe('Set text color using hex format (e.g., "#FF0000").'),
backgroundColor: z.string()
.refine(validateHexColor, { message: "Invalid hex color format (e.g., #00FF00 or #0F0)" })
.optional()
.describe('Set text background color using hex format (e.g., "#FFFF00").'),
linkUrl: z.string().url().optional().describe('Make the text a hyperlink pointing to this URL.'),
// clearDirectFormatting: z.boolean().optional().describe('If true, attempts to clear all direct text formatting within the range before applying new styles.') // Harder to implement perfectly
}).describe("Parameters for character-level text formatting.");

// Subset of TextStyle used for passing to helpers
export type TextStyleArgs = z.infer<typeof TextStyleParameters>;

export const ParagraphStyleParameters = z.object({
alignment: z.enum(['START', 'END', 'CENTER', 'JUSTIFIED']).optional().describe('Paragraph alignment. START=left for LTR languages, END=right for LTR languages.'),
indentStart: z.number().min(0).optional().describe('Left indentation in points.'),
indentEnd: z.number().min(0).optional().describe('Right indentation in points.'),
spaceAbove: z.number().min(0).optional().describe('Space before the paragraph in points.'),
spaceBelow: z.number().min(0).optional().describe('Space after the paragraph in points.'),
namedStyleType: z.enum([
'NORMAL_TEXT', 'TITLE', 'SUBTITLE',
'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6'
]).optional().describe('Apply a built-in named paragraph style (e.g., HEADING_1).'),
keepWithNext: z.boolean().optional().describe('Keep this paragraph together with the next one on the same page.'),
// Borders are more complex, might need separate objects/tools
// clearDirectFormatting: z.boolean().optional().describe('If true, attempts to clear all direct paragraph formatting within the range before applying new styles.') // Harder to implement perfectly
}).describe("Parameters for paragraph-level formatting.");

// Subset of ParagraphStyle used for passing to helpers
export type ParagraphStyleArgs = z.infer<typeof ParagraphStyleParameters>;

// --- Combination Schemas for Tools ---

export const ApplyTextStyleToolParameters = DocumentIdParameter.extend({
// Target EITHER by range OR by finding text
target: z.union([
RangeParameters,
TextFindParameter
]).describe("Specify the target range either by start/end indices or by finding specific text."),
style: TextStyleParameters.refine(
styleArgs => Object.values(styleArgs).some(v => v !== undefined),
{ message: "At least one text style option must be provided." }
).describe("The text styling to apply.")
});
export type ApplyTextStyleToolArgs = z.infer<typeof ApplyTextStyleToolParameters>;

export const ApplyParagraphStyleToolParameters = DocumentIdParameter.extend({
// Target EITHER by range OR by finding text (tool logic needs to find paragraph boundaries)
target: z.union([
RangeParameters, // User provides paragraph start/end (less likely)
TextFindParameter, // Find text within paragraph to apply style
z.object({ // Target by specific index within the paragraph
indexWithinParagraph: z.number().int().min(1).describe("An index located anywhere within the target paragraph.")
})
]).describe("Specify the target paragraph either by start/end indices, by finding text within it, or by providing an index within it."),
style: ParagraphStyleParameters.refine(
styleArgs => Object.values(styleArgs).some(v => v !== undefined),
{ message: "At least one paragraph style option must be provided." }
).describe("The paragraph styling to apply.")
});
export type ApplyParagraphStyleToolArgs = z.infer<typeof ApplyParagraphStyleToolParameters>;

// --- Error Class ---
// Use FastMCP's UserError for client-facing issues
// Define a custom error for internal issues if needed
export class NotImplementedError extends Error {
constructor(message = "This feature is not yet implemented.") {
super(message);
this.name = "NotImplementedError";
}
}

// === GOOGLE SLIDES SCHEMA FRAGMENTS ===

export const PresentationIdParameter = z.object({
  presentationId: z.string().describe('The ID of the Google Slides presentation (from the URL: docs.google.com/presentation/d/PRESENTATION_ID/edit).'),
});

export const PageObjectIdParameter = z.object({
  pageObjectId: z.string().describe('The object ID of the slide/page.'),
});

export const SlidePositionParameter = z.object({
  insertionIndex: z.number().int().min(0).optional().describe('Optional: The 0-based index where to insert the slide. Omit to add at the end.'),
});

export const ElementSizeParameter = z.object({
  width: z.number().positive().describe('Width of the element in points.'),
  height: z.number().positive().describe('Height of the element in points.'),
});

export const ElementPositionParameter = z.object({
  x: z.number().min(0).describe('X position (left edge) in points from the top-left corner of the slide.'),
  y: z.number().min(0).describe('Y position (top edge) in points from the top-left corner of the slide.'),
});

export const ShapeTypeEnum = z.enum([
  'TEXT_BOX',
  'RECTANGLE',
  'ROUND_RECTANGLE',
  'ELLIPSE',
  'TRIANGLE',
  'RIGHT_TRIANGLE',
  'PARALLELOGRAM',
  'TRAPEZOID',
  'PENTAGON',
  'HEXAGON',
  'HEPTAGON',
  'OCTAGON',
  'STAR_4',
  'STAR_5',
  'STAR_6',
  'STAR_8',
  'STAR_10',
  'STAR_12',
  'ARROW_EAST',
  'ARROW_NORTH',
  'ARROW_NORTH_EAST',
  'CLOUD',
  'HEART',
  'PLUS',
]).describe('The type of shape to create.');

export const PredefinedLayoutEnum = z.enum([
  'BLANK',
  'CAPTION_ONLY',
  'TITLE',
  'TITLE_AND_BODY',
  'TITLE_AND_TWO_COLUMNS',
  'TITLE_ONLY',
  'SECTION_HEADER',
  'SECTION_TITLE_AND_DESCRIPTION',
  'ONE_COLUMN_TEXT',
  'MAIN_POINT',
  'BIG_NUMBER',
]).describe('Predefined slide layout type.');

// Type exports for Slides
export type PresentationIdArgs = z.infer<typeof PresentationIdParameter>;
export type PageObjectIdArgs = z.infer<typeof PageObjectIdParameter>;
export type ElementSizeArgs = z.infer<typeof ElementSizeParameter>;
export type ElementPositionArgs = z.infer<typeof ElementPositionParameter>;

// === GMAIL SCHEMA FRAGMENTS ===

// --- Basic Gmail Parameters ---
export const MessageIdParameter = z.object({
  messageId: z.string().describe('The ID of the Gmail message.'),
});

export const ThreadIdParameter = z.object({
  threadId: z.string().describe('The ID of the Gmail thread/conversation.'),
});

export const DraftIdParameter = z.object({
  draftId: z.string().describe('The ID of the Gmail draft.'),
});

export const LabelIdParameter = z.object({
  labelId: z.string().describe('The ID of the Gmail label.'),
});

export const FilterIdParameter = z.object({
  filterId: z.string().describe('The ID of the Gmail filter.'),
});

// --- Email Composition Parameters ---
export const EmailRecipientsParameter = z.object({
  to: z.array(z.string().email()).describe('List of recipient email addresses.'),
  cc: z.array(z.string().email()).optional().describe('List of CC recipient email addresses.'),
  bcc: z.array(z.string().email()).optional().describe('List of BCC recipient email addresses.'),
});

export const EmailContentParameter = z.object({
  subject: z.string().describe('The email subject line.'),
  body: z.string().describe('The email body content (plain text or HTML based on mimeType).'),
  htmlBody: z.string().optional().describe('HTML version of the email body (for multipart/alternative).'),
  mimeType: z.enum(['text/plain', 'text/html', 'multipart/alternative']).optional().default('text/plain')
    .describe('Content type of the email body.'),
});

export const EmailAttachmentsParameter = z.object({
  attachments: z.array(z.string()).optional().describe('List of file paths to attach to the email.'),
});

export const EmailReplyParameter = z.object({
  threadId: z.string().optional().describe('Thread ID to reply to (maintains conversation).'),
  inReplyTo: z.string().optional().describe('Message-ID header of the email being replied to.'),
});

// --- Search Parameters ---
export const GmailSearchParameter = z.object({
  query: z.string().describe('Gmail search query using Gmail search syntax (e.g., "from:user@example.com", "is:unread", "subject:meeting").'),
  maxResults: z.number().int().min(1).max(500).optional().default(10).describe('Maximum number of results to return (1-500).'),
  pageToken: z.string().optional().describe('Page token for pagination.'),
  includeSpamTrash: z.boolean().optional().default(false).describe('Include messages from SPAM and TRASH in results.'),
});

// --- Label Management Parameters ---
export const LabelVisibilityParameter = z.object({
  labelListVisibility: z.enum(['labelShow', 'labelShowIfUnread', 'labelHide']).optional()
    .describe('Visibility of the label in the label list.'),
  messageListVisibility: z.enum(['show', 'hide']).optional()
    .describe('Whether to show or hide the label in the message list.'),
});

export const CreateLabelParameter = z.object({
  name: z.string().min(1).describe('Name for the new label.'),
}).merge(LabelVisibilityParameter);

export const UpdateLabelParameter = LabelIdParameter.extend({
  name: z.string().min(1).optional().describe('New name for the label.'),
}).merge(LabelVisibilityParameter);

// --- Label Modification Parameters ---
export const ModifyLabelsParameter = MessageIdParameter.extend({
  addLabelIds: z.array(z.string()).optional().describe('List of label IDs to add to the message.'),
  removeLabelIds: z.array(z.string()).optional().describe('List of label IDs to remove from the message.'),
});

export const BatchModifyLabelsParameter = z.object({
  messageIds: z.array(z.string()).min(1).describe('List of message IDs to modify.'),
  addLabelIds: z.array(z.string()).optional().describe('List of label IDs to add to all messages.'),
  removeLabelIds: z.array(z.string()).optional().describe('List of label IDs to remove from all messages.'),
  batchSize: z.number().int().min(1).max(1000).optional().default(50).describe('Number of messages to process in each batch.'),
});

// --- Filter Parameters ---
export const FilterCriteriaParameter = z.object({
  from: z.string().optional().describe('Sender email address to match.'),
  to: z.string().optional().describe('Recipient email address to match.'),
  subject: z.string().optional().describe('Subject text to match.'),
  query: z.string().optional().describe('Gmail search query for advanced matching.'),
  negatedQuery: z.string().optional().describe('Query that messages must NOT match.'),
  hasAttachment: z.boolean().optional().describe('Match only messages with attachments.'),
  excludeChats: z.boolean().optional().describe('Exclude chat messages from matching.'),
  size: z.number().int().optional().describe('Message size in bytes for size-based filtering.'),
  sizeComparison: z.enum(['unspecified', 'smaller', 'larger']).optional().describe('Size comparison operator.'),
});

export const FilterActionParameter = z.object({
  addLabelIds: z.array(z.string()).optional().describe('Label IDs to add to matching messages.'),
  removeLabelIds: z.array(z.string()).optional().describe('Label IDs to remove from matching messages.'),
  forward: z.string().email().optional().describe('Email address to forward matching messages to.'),
});

export const CreateFilterParameter = z.object({
  criteria: FilterCriteriaParameter.describe('Criteria for matching emails.'),
  action: FilterActionParameter.describe('Actions to perform on matching emails.'),
});

// --- Filter Templates ---
export const FilterTemplateType = z.enum([
  'fromSender',
  'withSubject',
  'withAttachments',
  'largeEmails',
  'containingText',
  'mailingList',
]).describe('Pre-defined filter template type.');

export const FilterTemplateParameter = z.object({
  template: FilterTemplateType,
  parameters: z.object({
    senderEmail: z.string().email().optional().describe('Sender email for fromSender template.'),
    subjectText: z.string().optional().describe('Subject text for withSubject template.'),
    searchText: z.string().optional().describe('Search text for containingText template.'),
    listIdentifier: z.string().optional().describe('Mailing list identifier for mailingList template.'),
    sizeInBytes: z.number().int().optional().describe('Size threshold for largeEmails template.'),
    labelIds: z.array(z.string()).optional().describe('Label IDs to apply.'),
    archive: z.boolean().optional().describe('Whether to skip the inbox (archive).'),
    markAsRead: z.boolean().optional().describe('Whether to mark as read.'),
    markImportant: z.boolean().optional().describe('Whether to mark as important.'),
  }).describe('Template-specific parameters.'),
});

// --- Attachment Parameters ---
export const DownloadAttachmentParameter = MessageIdParameter.extend({
  attachmentId: z.string().describe('The ID of the attachment to download.'),
  savePath: z.string().optional().describe('Directory path to save the attachment.'),
  filename: z.string().optional().describe('Custom filename for the saved attachment.'),
});

// --- Batch Operations ---
export const BatchDeleteParameter = z.object({
  messageIds: z.array(z.string()).min(1).describe('List of message IDs to delete.'),
  batchSize: z.number().int().min(1).max(1000).optional().default(50).describe('Number of messages to process in each batch.'),
});

// --- Combined Email Parameters for Tools ---
export const SendEmailParameter = EmailRecipientsParameter
  .merge(EmailContentParameter)
  .merge(EmailAttachmentsParameter)
  .merge(EmailReplyParameter);

export const DraftEmailParameter = EmailRecipientsParameter
  .merge(EmailContentParameter)
  .merge(EmailAttachmentsParameter)
  .merge(EmailReplyParameter);

// --- Type Exports for Gmail ---
export type MessageIdArgs = z.infer<typeof MessageIdParameter>;
export type ThreadIdArgs = z.infer<typeof ThreadIdParameter>;
export type DraftIdArgs = z.infer<typeof DraftIdParameter>;
export type LabelIdArgs = z.infer<typeof LabelIdParameter>;
export type FilterIdArgs = z.infer<typeof FilterIdParameter>;
export type EmailRecipientsArgs = z.infer<typeof EmailRecipientsParameter>;
export type EmailContentArgs = z.infer<typeof EmailContentParameter>;
export type EmailAttachmentsArgs = z.infer<typeof EmailAttachmentsParameter>;
export type GmailSearchArgs = z.infer<typeof GmailSearchParameter>;
export type CreateLabelArgs = z.infer<typeof CreateLabelParameter>;
export type UpdateLabelArgs = z.infer<typeof UpdateLabelParameter>;
export type ModifyLabelsArgs = z.infer<typeof ModifyLabelsParameter>;
export type BatchModifyLabelsArgs = z.infer<typeof BatchModifyLabelsParameter>;
export type FilterCriteriaArgs = z.infer<typeof FilterCriteriaParameter>;
export type FilterActionArgs = z.infer<typeof FilterActionParameter>;
export type CreateFilterArgs = z.infer<typeof CreateFilterParameter>;
export type FilterTemplateArgs = z.infer<typeof FilterTemplateParameter>;
export type DownloadAttachmentArgs = z.infer<typeof DownloadAttachmentParameter>;
export type BatchDeleteArgs = z.infer<typeof BatchDeleteParameter>;
export type SendEmailArgs = z.infer<typeof SendEmailParameter>;
export type DraftEmailArgs = z.infer<typeof DraftEmailParameter>;