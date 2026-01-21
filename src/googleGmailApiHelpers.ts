// src/googleGmailApiHelpers.ts
import { gmail_v1 } from 'googleapis';
import { UserError } from 'fastmcp';
import * as fs from 'fs/promises';
import * as path from 'path';

type Gmail = gmail_v1.Gmail;

// --- Constants ---
const MAX_BATCH_SIZE = 100; // Gmail API batch limit

// --- Email Validation ---
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

export function validateEmail(email: string): boolean {
  return EMAIL_REGEX.test(email);
}

// --- RFC 2047 MIME Encoding for Headers ---
export function encodeEmailHeader(text: string): string {
  // Check if encoding is needed (non-ASCII characters)
  if (/^[\x00-\x7F]*$/.test(text)) {
    return text; // Pure ASCII, no encoding needed
  }
  // Use UTF-8 Base64 encoding for non-ASCII
  const encoded = Buffer.from(text, 'utf-8').toString('base64');
  return `=?UTF-8?B?${encoded}?=`;
}

// --- Base64 URL-safe Encoding ---
export function base64UrlEncode(str: string): string {
  return Buffer.from(str, 'utf-8')
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

export function base64UrlDecode(str: string): string {
  // Add padding if needed
  let padded = str.replace(/-/g, '+').replace(/_/g, '/');
  const padding = 4 - (padded.length % 4);
  if (padding !== 4) {
    padded += '='.repeat(padding);
  }
  return Buffer.from(padded, 'base64').toString('utf-8');
}

// --- MIME Type Detection ---
export function getMimeType(filename: string): string {
  const ext = path.extname(filename).toLowerCase();
  const mimeTypes: Record<string, string> = {
    '.txt': 'text/plain',
    '.html': 'text/html',
    '.htm': 'text/html',
    '.css': 'text/css',
    '.js': 'application/javascript',
    '.json': 'application/json',
    '.xml': 'application/xml',
    '.pdf': 'application/pdf',
    '.doc': 'application/msword',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.xls': 'application/vnd.ms-excel',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.ppt': 'application/vnd.ms-powerpoint',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.webp': 'image/webp',
    '.ico': 'image/x-icon',
    '.mp3': 'audio/mpeg',
    '.wav': 'audio/wav',
    '.mp4': 'video/mp4',
    '.webm': 'video/webm',
    '.zip': 'application/zip',
    '.gz': 'application/gzip',
    '.tar': 'application/x-tar',
    '.rar': 'application/vnd.rar',
    '.7z': 'application/x-7z-compressed',
    '.csv': 'text/csv',
    '.md': 'text/markdown',
    '.yaml': 'text/yaml',
    '.yml': 'text/yaml',
  };
  return mimeTypes[ext] || 'application/octet-stream';
}

// --- Simple Email Creation (no attachments) ---
export function createSimpleEmail(options: {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  htmlBody?: string;
  mimeType?: 'text/plain' | 'text/html' | 'multipart/alternative';
  inReplyTo?: string;
  threadId?: string;
}): string {
  const { to, cc, bcc, subject, body, htmlBody, mimeType = 'text/plain', inReplyTo } = options;

  const boundary = `boundary_${Date.now()}_${Math.random().toString(36).substring(7)}`;
  const headers: string[] = [];

  // Build headers
  headers.push(`To: ${to.join(', ')}`);
  if (cc && cc.length > 0) {
    headers.push(`Cc: ${cc.join(', ')}`);
  }
  if (bcc && bcc.length > 0) {
    headers.push(`Bcc: ${bcc.join(', ')}`);
  }
  headers.push(`Subject: ${encodeEmailHeader(subject)}`);
  headers.push('MIME-Version: 1.0');

  if (inReplyTo) {
    headers.push(`In-Reply-To: ${inReplyTo}`);
    headers.push(`References: ${inReplyTo}`);
  }

  // Handle different MIME types
  if (mimeType === 'multipart/alternative' && htmlBody) {
    headers.push(`Content-Type: multipart/alternative; boundary="${boundary}"`);
    const bodyParts = [
      '',
      `--${boundary}`,
      'Content-Type: text/plain; charset=UTF-8',
      '',
      body,
      '',
      `--${boundary}`,
      'Content-Type: text/html; charset=UTF-8',
      '',
      htmlBody,
      '',
      `--${boundary}--`,
    ];
    return headers.join('\r\n') + '\r\n\r\n' + bodyParts.join('\r\n');
  } else {
    const contentType = mimeType === 'text/html' ? 'text/html' : 'text/plain';
    headers.push(`Content-Type: ${contentType}; charset=UTF-8`);
    return headers.join('\r\n') + '\r\n\r\n' + body;
  }
}

// --- Email Creation with Attachments ---
export async function createEmailWithAttachments(options: {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  htmlBody?: string;
  mimeType?: 'text/plain' | 'text/html' | 'multipart/alternative';
  attachments?: string[];
  inReplyTo?: string;
}): Promise<string> {
  const { to, cc, bcc, subject, body, htmlBody, mimeType = 'text/plain', attachments, inReplyTo } = options;

  // If no attachments, use simple email creation
  if (!attachments || attachments.length === 0) {
    return createSimpleEmail(options);
  }

  const boundary = `boundary_${Date.now()}_${Math.random().toString(36).substring(7)}`;
  const altBoundary = `alt_${boundary}`;
  const headers: string[] = [];

  // Build headers
  headers.push(`To: ${to.join(', ')}`);
  if (cc && cc.length > 0) {
    headers.push(`Cc: ${cc.join(', ')}`);
  }
  if (bcc && bcc.length > 0) {
    headers.push(`Bcc: ${bcc.join(', ')}`);
  }
  headers.push(`Subject: ${encodeEmailHeader(subject)}`);
  headers.push('MIME-Version: 1.0');
  headers.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);

  if (inReplyTo) {
    headers.push(`In-Reply-To: ${inReplyTo}`);
    headers.push(`References: ${inReplyTo}`);
  }

  const parts: string[] = [''];

  // Add text/html body part
  if (mimeType === 'multipart/alternative' && htmlBody) {
    parts.push(`--${boundary}`);
    parts.push(`Content-Type: multipart/alternative; boundary="${altBoundary}"`);
    parts.push('');
    parts.push(`--${altBoundary}`);
    parts.push('Content-Type: text/plain; charset=UTF-8');
    parts.push('');
    parts.push(body);
    parts.push('');
    parts.push(`--${altBoundary}`);
    parts.push('Content-Type: text/html; charset=UTF-8');
    parts.push('');
    parts.push(htmlBody);
    parts.push(`--${altBoundary}--`);
  } else {
    const contentType = mimeType === 'text/html' ? 'text/html' : 'text/plain';
    parts.push(`--${boundary}`);
    parts.push(`Content-Type: ${contentType}; charset=UTF-8`);
    parts.push('');
    parts.push(body);
  }

  // Add attachments
  for (const filePath of attachments) {
    try {
      const fileContent = await fs.readFile(filePath);
      const fileName = path.basename(filePath);
      const mimeType = getMimeType(fileName);
      const base64Content = fileContent.toString('base64');

      parts.push('');
      parts.push(`--${boundary}`);
      parts.push(`Content-Type: ${mimeType}; name="${encodeEmailHeader(fileName)}"`);
      parts.push('Content-Transfer-Encoding: base64');
      parts.push(`Content-Disposition: attachment; filename="${encodeEmailHeader(fileName)}"`);
      parts.push('');
      // Split base64 into 76-character lines per RFC 2045
      const lines = base64Content.match(/.{1,76}/g) || [];
      parts.push(lines.join('\r\n'));
    } catch (error: any) {
      throw new UserError(`Failed to read attachment "${filePath}": ${error.message}`);
    }
  }

  parts.push('');
  parts.push(`--${boundary}--`);

  return headers.join('\r\n') + '\r\n' + parts.join('\r\n');
}

// --- Parse Email Headers ---
export function parseEmailHeaders(headers: gmail_v1.Schema$MessagePartHeader[]): Record<string, string> {
  const result: Record<string, string> = {};
  for (const header of headers) {
    if (header.name && header.value) {
      result[header.name.toLowerCase()] = header.value;
    }
  }
  return result;
}

// --- Extract Plain Text from Message ---
export function extractPlainText(payload: gmail_v1.Schema$MessagePart): string {
  if (!payload) return '';

  // Direct text/plain part
  if (payload.mimeType === 'text/plain' && payload.body?.data) {
    return base64UrlDecode(payload.body.data);
  }

  // Look in parts
  if (payload.parts) {
    for (const part of payload.parts) {
      if (part.mimeType === 'text/plain' && part.body?.data) {
        return base64UrlDecode(part.body.data);
      }
      // Recurse into multipart
      if (part.mimeType?.startsWith('multipart/')) {
        const text = extractPlainText(part);
        if (text) return text;
      }
    }
    // If no text/plain, try text/html as fallback
    for (const part of payload.parts) {
      if (part.mimeType === 'text/html' && part.body?.data) {
        return base64UrlDecode(part.body.data);
      }
    }
  }

  return '';
}

// --- Extract HTML from Message ---
export function extractHtmlContent(payload: gmail_v1.Schema$MessagePart): string {
  if (!payload) return '';

  // Direct text/html part
  if (payload.mimeType === 'text/html' && payload.body?.data) {
    return base64UrlDecode(payload.body.data);
  }

  // Look in parts
  if (payload.parts) {
    for (const part of payload.parts) {
      if (part.mimeType === 'text/html' && part.body?.data) {
        return base64UrlDecode(part.body.data);
      }
      // Recurse into multipart
      if (part.mimeType?.startsWith('multipart/')) {
        const html = extractHtmlContent(part);
        if (html) return html;
      }
    }
  }

  return '';
}

// --- Extract Attachment Info ---
export interface AttachmentInfo {
  attachmentId: string;
  filename: string;
  mimeType: string;
  size: number;
}

export function extractAttachments(payload: gmail_v1.Schema$MessagePart): AttachmentInfo[] {
  const attachments: AttachmentInfo[] = [];

  function processpart(part: gmail_v1.Schema$MessagePart) {
    if (part.filename && part.body?.attachmentId) {
      attachments.push({
        attachmentId: part.body.attachmentId,
        filename: part.filename,
        mimeType: part.mimeType || 'application/octet-stream',
        size: part.body.size || 0,
      });
    }
    if (part.parts) {
      for (const subPart of part.parts) {
        processpart(subPart);
      }
    }
  }

  if (payload) {
    processpart(payload);
  }

  return attachments;
}

// --- Format Message for Display ---
export interface FormattedMessage {
  id: string;
  threadId: string;
  labelIds: string[];
  snippet: string;
  from: string;
  to: string;
  cc?: string;
  subject: string;
  date: string;
  body: string;
  htmlBody?: string;
  attachments: AttachmentInfo[];
  isUnread: boolean;
  isStarred: boolean;
}

export function formatMessage(message: gmail_v1.Schema$Message): FormattedMessage {
  const headers = parseEmailHeaders(message.payload?.headers || []);
  const labelIds = message.labelIds || [];

  return {
    id: message.id || '',
    threadId: message.threadId || '',
    labelIds,
    snippet: message.snippet || '',
    from: headers['from'] || '',
    to: headers['to'] || '',
    cc: headers['cc'],
    subject: headers['subject'] || '(No Subject)',
    date: headers['date'] || '',
    body: extractPlainText(message.payload || {}),
    htmlBody: extractHtmlContent(message.payload || {}),
    attachments: extractAttachments(message.payload || {}),
    isUnread: labelIds.includes('UNREAD'),
    isStarred: labelIds.includes('STARRED'),
  };
}

// --- Send Email Helper ---
export async function sendEmail(gmail: Gmail, rawEmail: string, threadId?: string): Promise<gmail_v1.Schema$Message> {
  try {
    const encodedEmail = base64UrlEncode(rawEmail);
    const response = await gmail.users.messages.send({
      userId: 'me',
      requestBody: {
        raw: encodedEmail,
        threadId: threadId,
      },
    });
    return response.data;
  } catch (error: any) {
    const message = error.response?.data?.error?.message || error.message;
    if (error.code === 400) {
      throw new UserError(`Invalid email format: ${message}`);
    }
    if (error.code === 403) {
      throw new UserError(`Permission denied to send email: ${message}`);
    }
    throw new Error(`Gmail API Error: ${message}`);
  }
}

// --- Create Draft Helper ---
export async function createDraft(gmail: Gmail, rawEmail: string, threadId?: string): Promise<gmail_v1.Schema$Draft> {
  try {
    const encodedEmail = base64UrlEncode(rawEmail);
    const response = await gmail.users.drafts.create({
      userId: 'me',
      requestBody: {
        message: {
          raw: encodedEmail,
          threadId: threadId,
        },
      },
    });
    return response.data;
  } catch (error: any) {
    const message = error.response?.data?.error?.message || error.message;
    throw new Error(`Gmail API Error creating draft: ${message}`);
  }
}

// --- Get Message Helper ---
export async function getMessage(gmail: Gmail, messageId: string, format: 'full' | 'metadata' | 'minimal' | 'raw' = 'full'): Promise<gmail_v1.Schema$Message> {
  try {
    const response = await gmail.users.messages.get({
      userId: 'me',
      id: messageId,
      format,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Message not found (ID: ${messageId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Search Messages Helper ---
export async function searchMessages(gmail: Gmail, options: {
  query: string;
  maxResults?: number;
  pageToken?: string;
  includeSpamTrash?: boolean;
}): Promise<{
  messages: gmail_v1.Schema$Message[];
  nextPageToken?: string;
  resultSizeEstimate?: number;
}> {
  try {
    const response = await gmail.users.messages.list({
      userId: 'me',
      q: options.query,
      maxResults: options.maxResults || 10,
      pageToken: options.pageToken,
      includeSpamTrash: options.includeSpamTrash || false,
    });

    const messages: gmail_v1.Schema$Message[] = [];
    if (response.data.messages) {
      // Get full message data for each result
      for (const msg of response.data.messages) {
        if (msg.id) {
          const fullMessage = await getMessage(gmail, msg.id, 'metadata');
          messages.push(fullMessage);
        }
      }
    }

    return {
      messages,
      nextPageToken: response.data.nextPageToken || undefined,
      resultSizeEstimate: response.data.resultSizeEstimate || undefined,
    };
  } catch (error: any) {
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Get Thread Helper ---
export async function getThread(gmail: Gmail, threadId: string, format: 'full' | 'metadata' | 'minimal' = 'full'): Promise<gmail_v1.Schema$Thread> {
  try {
    const response = await gmail.users.threads.get({
      userId: 'me',
      id: threadId,
      format,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Thread not found (ID: ${threadId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- List Threads Helper ---
export async function listThreads(gmail: Gmail, options: {
  query?: string;
  maxResults?: number;
  pageToken?: string;
  includeSpamTrash?: boolean;
}): Promise<{
  threads: gmail_v1.Schema$Thread[];
  nextPageToken?: string;
  resultSizeEstimate?: number;
}> {
  try {
    const response = await gmail.users.threads.list({
      userId: 'me',
      q: options.query,
      maxResults: options.maxResults || 10,
      pageToken: options.pageToken,
      includeSpamTrash: options.includeSpamTrash || false,
    });

    return {
      threads: response.data.threads || [],
      nextPageToken: response.data.nextPageToken || undefined,
      resultSizeEstimate: response.data.resultSizeEstimate || undefined,
    };
  } catch (error: any) {
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Modify Message Labels Helper ---
export async function modifyMessageLabels(gmail: Gmail, messageId: string, addLabelIds?: string[], removeLabelIds?: string[]): Promise<gmail_v1.Schema$Message> {
  try {
    const response = await gmail.users.messages.modify({
      userId: 'me',
      id: messageId,
      requestBody: {
        addLabelIds: addLabelIds || [],
        removeLabelIds: removeLabelIds || [],
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Message not found (ID: ${messageId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Batch Modify Messages Helper ---
export async function batchModifyMessages(gmail: Gmail, messageIds: string[], addLabelIds?: string[], removeLabelIds?: string[], batchSize: number = 50): Promise<{
  processed: number;
  errors: string[];
}> {
  const errors: string[] = [];
  let processed = 0;

  // Process in batches
  for (let i = 0; i < messageIds.length; i += batchSize) {
    const batch = messageIds.slice(i, i + batchSize);
    try {
      await gmail.users.messages.batchModify({
        userId: 'me',
        requestBody: {
          ids: batch,
          addLabelIds: addLabelIds || [],
          removeLabelIds: removeLabelIds || [],
        },
      });
      processed += batch.length;
    } catch (error: any) {
      errors.push(`Batch ${Math.floor(i / batchSize) + 1} failed: ${error.message}`);
    }
  }

  return { processed, errors };
}

// --- Delete Message Helper ---
export async function deleteMessage(gmail: Gmail, messageId: string): Promise<void> {
  try {
    await gmail.users.messages.delete({
      userId: 'me',
      id: messageId,
    });
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Message not found (ID: ${messageId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Batch Delete Messages Helper ---
export async function batchDeleteMessages(gmail: Gmail, messageIds: string[], batchSize: number = 50): Promise<{
  deleted: number;
  errors: string[];
}> {
  const errors: string[] = [];
  let deleted = 0;

  // Process in batches
  for (let i = 0; i < messageIds.length; i += batchSize) {
    const batch = messageIds.slice(i, i + batchSize);
    try {
      await gmail.users.messages.batchDelete({
        userId: 'me',
        requestBody: {
          ids: batch,
        },
      });
      deleted += batch.length;
    } catch (error: any) {
      errors.push(`Batch ${Math.floor(i / batchSize) + 1} failed: ${error.message}`);
    }
  }

  return { deleted, errors };
}

// --- Trash Message Helper ---
export async function trashMessage(gmail: Gmail, messageId: string): Promise<gmail_v1.Schema$Message> {
  try {
    const response = await gmail.users.messages.trash({
      userId: 'me',
      id: messageId,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Message not found (ID: ${messageId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Untrash Message Helper ---
export async function untrashMessage(gmail: Gmail, messageId: string): Promise<gmail_v1.Schema$Message> {
  try {
    const response = await gmail.users.messages.untrash({
      userId: 'me',
      id: messageId,
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Message not found (ID: ${messageId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Download Attachment Helper ---
export async function downloadAttachment(gmail: Gmail, messageId: string, attachmentId: string, savePath?: string, filename?: string): Promise<{
  savedTo: string;
  size: number;
}> {
  try {
    // Get the attachment data
    const response = await gmail.users.messages.attachments.get({
      userId: 'me',
      messageId,
      id: attachmentId,
    });

    if (!response.data.data) {
      throw new UserError('Attachment data is empty.');
    }

    // Decode base64url data
    const data = Buffer.from(response.data.data, 'base64');

    // Determine save path
    const saveDir = savePath || process.cwd();
    const saveFilename = filename || `attachment_${attachmentId.substring(0, 8)}`;
    const fullPath = path.join(saveDir, saveFilename);

    // Ensure directory exists
    await fs.mkdir(saveDir, { recursive: true });

    // Write file
    await fs.writeFile(fullPath, data);

    return {
      savedTo: fullPath,
      size: data.length,
    };
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Attachment not found (Message: ${messageId}, Attachment: ${attachmentId}).`);
    }
    throw error instanceof UserError ? error : new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Get User Profile Helper ---
export async function getUserProfile(gmail: Gmail): Promise<gmail_v1.Schema$Profile> {
  try {
    const response = await gmail.users.getProfile({
      userId: 'me',
    });
    return response.data;
  } catch (error: any) {
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Draft Management Helpers ---
export async function listDrafts(gmail: Gmail, options: {
  maxResults?: number;
  pageToken?: string;
}): Promise<{
  drafts: gmail_v1.Schema$Draft[];
  nextPageToken?: string;
}> {
  try {
    const response = await gmail.users.drafts.list({
      userId: 'me',
      maxResults: options.maxResults || 10,
      pageToken: options.pageToken,
    });
    return {
      drafts: response.data.drafts || [],
      nextPageToken: response.data.nextPageToken || undefined,
    };
  } catch (error: any) {
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

export async function getDraft(gmail: Gmail, draftId: string): Promise<gmail_v1.Schema$Draft> {
  try {
    const response = await gmail.users.drafts.get({
      userId: 'me',
      id: draftId,
      format: 'full',
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Draft not found (ID: ${draftId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

export async function updateDraft(gmail: Gmail, draftId: string, rawEmail: string): Promise<gmail_v1.Schema$Draft> {
  try {
    const encodedEmail = base64UrlEncode(rawEmail);
    const response = await gmail.users.drafts.update({
      userId: 'me',
      id: draftId,
      requestBody: {
        message: {
          raw: encodedEmail,
        },
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Draft not found (ID: ${draftId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

export async function deleteDraft(gmail: Gmail, draftId: string): Promise<void> {
  try {
    await gmail.users.drafts.delete({
      userId: 'me',
      id: draftId,
    });
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Draft not found (ID: ${draftId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}

export async function sendDraft(gmail: Gmail, draftId: string): Promise<gmail_v1.Schema$Message> {
  try {
    const response = await gmail.users.drafts.send({
      userId: 'me',
      requestBody: {
        id: draftId,
      },
    });
    return response.data;
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Draft not found (ID: ${draftId}).`);
    }
    throw new Error(`Gmail API Error: ${error.message}`);
  }
}
