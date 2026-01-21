// src/gmailLabelManager.ts
import { gmail_v1 } from 'googleapis';
import { UserError } from 'fastmcp';

type Gmail = gmail_v1.Gmail;

// --- Label Types ---
export interface LabelInfo {
  id: string;
  name: string;
  type: 'system' | 'user';
  messageListVisibility?: 'show' | 'hide';
  labelListVisibility?: 'labelShow' | 'labelShowIfUnread' | 'labelHide';
  messagesTotal?: number;
  messagesUnread?: number;
  threadsTotal?: number;
  threadsUnread?: number;
}

// --- System Labels (for reference) ---
export const SYSTEM_LABELS = [
  'INBOX',
  'SPAM',
  'TRASH',
  'UNREAD',
  'STARRED',
  'IMPORTANT',
  'SENT',
  'DRAFT',
  'CATEGORY_PERSONAL',
  'CATEGORY_SOCIAL',
  'CATEGORY_PROMOTIONS',
  'CATEGORY_UPDATES',
  'CATEGORY_FORUMS',
] as const;

export type SystemLabel = typeof SYSTEM_LABELS[number];

// --- List All Labels ---
export async function listLabels(gmail: Gmail): Promise<LabelInfo[]> {
  try {
    const response = await gmail.users.labels.list({
      userId: 'me',
    });

    const labels: LabelInfo[] = [];
    for (const label of response.data.labels || []) {
      if (label.id && label.name) {
        labels.push({
          id: label.id,
          name: label.name,
          type: label.type === 'system' ? 'system' : 'user',
          messageListVisibility: label.messageListVisibility as 'show' | 'hide' | undefined,
          labelListVisibility: label.labelListVisibility as 'labelShow' | 'labelShowIfUnread' | 'labelHide' | undefined,
          messagesTotal: label.messagesTotal || undefined,
          messagesUnread: label.messagesUnread || undefined,
          threadsTotal: label.threadsTotal || undefined,
          threadsUnread: label.threadsUnread || undefined,
        });
      }
    }

    return labels;
  } catch (error: any) {
    throw new Error(`Gmail API Error listing labels: ${error.message}`);
  }
}

// --- Get Label by ID ---
export async function getLabel(gmail: Gmail, labelId: string): Promise<LabelInfo> {
  try {
    const response = await gmail.users.labels.get({
      userId: 'me',
      id: labelId,
    });

    const label = response.data;
    if (!label.id || !label.name) {
      throw new UserError(`Invalid label data received for ID: ${labelId}`);
    }

    return {
      id: label.id,
      name: label.name,
      type: label.type === 'system' ? 'system' : 'user',
      messageListVisibility: label.messageListVisibility as 'show' | 'hide' | undefined,
      labelListVisibility: label.labelListVisibility as 'labelShow' | 'labelShowIfUnread' | 'labelHide' | undefined,
      messagesTotal: label.messagesTotal || undefined,
      messagesUnread: label.messagesUnread || undefined,
      threadsTotal: label.threadsTotal || undefined,
      threadsUnread: label.threadsUnread || undefined,
    };
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Label not found (ID: ${labelId}).`);
    }
    throw error instanceof UserError ? error : new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Find Label by Name ---
export async function findLabelByName(gmail: Gmail, name: string): Promise<LabelInfo | null> {
  const labels = await listLabels(gmail);
  const normalizedName = name.toLowerCase();
  return labels.find(l => l.name.toLowerCase() === normalizedName) || null;
}

// --- Create Label ---
export async function createLabel(gmail: Gmail, options: {
  name: string;
  labelListVisibility?: 'labelShow' | 'labelShowIfUnread' | 'labelHide';
  messageListVisibility?: 'show' | 'hide';
}): Promise<LabelInfo> {
  try {
    // Check if label already exists
    const existing = await findLabelByName(gmail, options.name);
    if (existing) {
      throw new UserError(`A label named "${options.name}" already exists (ID: ${existing.id}).`);
    }

    const response = await gmail.users.labels.create({
      userId: 'me',
      requestBody: {
        name: options.name,
        labelListVisibility: options.labelListVisibility || 'labelShow',
        messageListVisibility: options.messageListVisibility || 'show',
      },
    });

    const label = response.data;
    if (!label.id || !label.name) {
      throw new Error('Invalid label data received from create operation.');
    }

    return {
      id: label.id,
      name: label.name,
      type: 'user',
      labelListVisibility: label.labelListVisibility as 'labelShow' | 'labelShowIfUnread' | 'labelHide' | undefined,
      messageListVisibility: label.messageListVisibility as 'show' | 'hide' | undefined,
    };
  } catch (error: any) {
    if (error instanceof UserError) throw error;
    if (error.code === 409) {
      throw new UserError(`A label named "${options.name}" already exists.`);
    }
    throw new Error(`Gmail API Error creating label: ${error.message}`);
  }
}

// --- Update Label ---
export async function updateLabel(gmail: Gmail, labelId: string, options: {
  name?: string;
  labelListVisibility?: 'labelShow' | 'labelShowIfUnread' | 'labelHide';
  messageListVisibility?: 'show' | 'hide';
}): Promise<LabelInfo> {
  try {
    // Verify label exists and is not a system label
    const existing = await getLabel(gmail, labelId);
    if (existing.type === 'system') {
      throw new UserError(`Cannot modify system label "${existing.name}".`);
    }

    // If renaming, check for conflicts
    if (options.name && options.name.toLowerCase() !== existing.name.toLowerCase()) {
      const conflict = await findLabelByName(gmail, options.name);
      if (conflict) {
        throw new UserError(`A label named "${options.name}" already exists (ID: ${conflict.id}).`);
      }
    }

    const response = await gmail.users.labels.update({
      userId: 'me',
      id: labelId,
      requestBody: {
        id: labelId,
        name: options.name || existing.name,
        labelListVisibility: options.labelListVisibility || existing.labelListVisibility,
        messageListVisibility: options.messageListVisibility || existing.messageListVisibility,
      },
    });

    const label = response.data;
    if (!label.id || !label.name) {
      throw new Error('Invalid label data received from update operation.');
    }

    return {
      id: label.id,
      name: label.name,
      type: 'user',
      labelListVisibility: label.labelListVisibility as 'labelShow' | 'labelShowIfUnread' | 'labelHide' | undefined,
      messageListVisibility: label.messageListVisibility as 'show' | 'hide' | undefined,
    };
  } catch (error: any) {
    if (error instanceof UserError) throw error;
    if (error.code === 404) {
      throw new UserError(`Label not found (ID: ${labelId}).`);
    }
    throw new Error(`Gmail API Error updating label: ${error.message}`);
  }
}

// --- Delete Label ---
export async function deleteLabel(gmail: Gmail, labelId: string): Promise<void> {
  try {
    // Verify label exists and is not a system label
    const existing = await getLabel(gmail, labelId);
    if (existing.type === 'system') {
      throw new UserError(`Cannot delete system label "${existing.name}".`);
    }

    await gmail.users.labels.delete({
      userId: 'me',
      id: labelId,
    });
  } catch (error: any) {
    if (error instanceof UserError) throw error;
    if (error.code === 404) {
      throw new UserError(`Label not found (ID: ${labelId}).`);
    }
    throw new Error(`Gmail API Error deleting label: ${error.message}`);
  }
}

// --- Get or Create Label (Idempotent) ---
export async function getOrCreateLabel(gmail: Gmail, options: {
  name: string;
  labelListVisibility?: 'labelShow' | 'labelShowIfUnread' | 'labelHide';
  messageListVisibility?: 'show' | 'hide';
}): Promise<LabelInfo> {
  // Try to find existing label
  const existing = await findLabelByName(gmail, options.name);
  if (existing) {
    return existing;
  }

  // Create new label
  return createLabel(gmail, options);
}

// --- Resolve Label IDs from Names ---
export async function resolveLabelIds(gmail: Gmail, labelsOrIds: string[]): Promise<string[]> {
  const allLabels = await listLabels(gmail);
  const labelMap = new Map<string, string>();

  // Build lookup maps
  for (const label of allLabels) {
    labelMap.set(label.id.toLowerCase(), label.id);
    labelMap.set(label.name.toLowerCase(), label.id);
  }

  const resolvedIds: string[] = [];
  const notFound: string[] = [];

  for (const input of labelsOrIds) {
    const normalized = input.toLowerCase();
    const id = labelMap.get(normalized);
    if (id) {
      resolvedIds.push(id);
    } else {
      notFound.push(input);
    }
  }

  if (notFound.length > 0) {
    throw new UserError(`Labels not found: ${notFound.join(', ')}`);
  }

  return resolvedIds;
}

// --- Format Labels for Display ---
export function formatLabelsForDisplay(labels: LabelInfo[]): string {
  const systemLabels = labels.filter(l => l.type === 'system');
  const userLabels = labels.filter(l => l.type === 'user');

  let output = '';

  if (systemLabels.length > 0) {
    output += 'System Labels:\n';
    for (const label of systemLabels) {
      const unread = label.messagesUnread ? ` (${label.messagesUnread} unread)` : '';
      output += `  - ${label.name}${unread}\n`;
    }
  }

  if (userLabels.length > 0) {
    output += '\nUser Labels:\n';
    for (const label of userLabels) {
      const unread = label.messagesUnread ? ` (${label.messagesUnread} unread)` : '';
      const visibility = label.labelListVisibility !== 'labelShow' ? ` [${label.labelListVisibility}]` : '';
      output += `  - ${label.name}${unread}${visibility} (ID: ${label.id})\n`;
    }
  }

  return output || 'No labels found.';
}
