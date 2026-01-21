// src/gmailFilterManager.ts
import { gmail_v1 } from 'googleapis';
import { UserError } from 'fastmcp';
import { FilterCriteriaArgs, FilterActionArgs } from './types.js';

type Gmail = gmail_v1.Gmail;

// --- Filter Types ---
export interface FilterInfo {
  id: string;
  criteria: FilterCriteriaArgs;
  action: FilterActionArgs;
}

// --- Filter Templates ---
export type FilterTemplateType =
  | 'fromSender'
  | 'withSubject'
  | 'withAttachments'
  | 'largeEmails'
  | 'containingText'
  | 'mailingList';

export interface FilterTemplateParams {
  senderEmail?: string;
  subjectText?: string;
  searchText?: string;
  listIdentifier?: string;
  sizeInBytes?: number;
  labelIds?: string[];
  archive?: boolean;
  markAsRead?: boolean;
  markImportant?: boolean;
}

// --- List All Filters ---
export async function listFilters(gmail: Gmail): Promise<FilterInfo[]> {
  try {
    const response = await gmail.users.settings.filters.list({
      userId: 'me',
    });

    const filters: FilterInfo[] = [];
    for (const filter of response.data.filter || []) {
      if (filter.id) {
        filters.push({
          id: filter.id,
          criteria: {
            from: filter.criteria?.from || undefined,
            to: filter.criteria?.to || undefined,
            subject: filter.criteria?.subject || undefined,
            query: filter.criteria?.query || undefined,
            negatedQuery: filter.criteria?.negatedQuery || undefined,
            hasAttachment: filter.criteria?.hasAttachment || undefined,
            excludeChats: filter.criteria?.excludeChats || undefined,
            size: filter.criteria?.size || undefined,
            sizeComparison: filter.criteria?.sizeComparison as 'unspecified' | 'smaller' | 'larger' | undefined,
          },
          action: {
            addLabelIds: filter.action?.addLabelIds || undefined,
            removeLabelIds: filter.action?.removeLabelIds || undefined,
            forward: filter.action?.forward || undefined,
          },
        });
      }
    }

    return filters;
  } catch (error: any) {
    throw new Error(`Gmail API Error listing filters: ${error.message}`);
  }
}

// --- Get Filter by ID ---
export async function getFilter(gmail: Gmail, filterId: string): Promise<FilterInfo> {
  try {
    const response = await gmail.users.settings.filters.get({
      userId: 'me',
      id: filterId,
    });

    const filter = response.data;
    if (!filter.id) {
      throw new UserError(`Invalid filter data received for ID: ${filterId}`);
    }

    return {
      id: filter.id,
      criteria: {
        from: filter.criteria?.from || undefined,
        to: filter.criteria?.to || undefined,
        subject: filter.criteria?.subject || undefined,
        query: filter.criteria?.query || undefined,
        negatedQuery: filter.criteria?.negatedQuery || undefined,
        hasAttachment: filter.criteria?.hasAttachment || undefined,
        excludeChats: filter.criteria?.excludeChats || undefined,
        size: filter.criteria?.size || undefined,
        sizeComparison: filter.criteria?.sizeComparison as 'unspecified' | 'smaller' | 'larger' | undefined,
      },
      action: {
        addLabelIds: filter.action?.addLabelIds || undefined,
        removeLabelIds: filter.action?.removeLabelIds || undefined,
        forward: filter.action?.forward || undefined,
      },
    };
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Filter not found (ID: ${filterId}).`);
    }
    throw error instanceof UserError ? error : new Error(`Gmail API Error: ${error.message}`);
  }
}

// --- Create Filter ---
export async function createFilter(gmail: Gmail, criteria: FilterCriteriaArgs, action: FilterActionArgs): Promise<FilterInfo> {
  try {
    // Validate that at least one criteria is provided
    const hasCriteria = Object.values(criteria).some(v => v !== undefined);
    if (!hasCriteria) {
      throw new UserError('At least one filter criteria must be specified.');
    }

    // Validate that at least one action is provided
    const hasAction = Object.values(action).some(v => v !== undefined && (Array.isArray(v) ? v.length > 0 : true));
    if (!hasAction) {
      throw new UserError('At least one filter action must be specified.');
    }

    const response = await gmail.users.settings.filters.create({
      userId: 'me',
      requestBody: {
        criteria: {
          from: criteria.from,
          to: criteria.to,
          subject: criteria.subject,
          query: criteria.query,
          negatedQuery: criteria.negatedQuery,
          hasAttachment: criteria.hasAttachment,
          excludeChats: criteria.excludeChats,
          size: criteria.size,
          sizeComparison: criteria.sizeComparison,
        },
        action: {
          addLabelIds: action.addLabelIds,
          removeLabelIds: action.removeLabelIds,
          forward: action.forward,
        },
      },
    });

    const filter = response.data;
    if (!filter.id) {
      throw new Error('Invalid filter data received from create operation.');
    }

    return {
      id: filter.id,
      criteria: {
        from: filter.criteria?.from || undefined,
        to: filter.criteria?.to || undefined,
        subject: filter.criteria?.subject || undefined,
        query: filter.criteria?.query || undefined,
        negatedQuery: filter.criteria?.negatedQuery || undefined,
        hasAttachment: filter.criteria?.hasAttachment || undefined,
        excludeChats: filter.criteria?.excludeChats || undefined,
        size: filter.criteria?.size || undefined,
        sizeComparison: filter.criteria?.sizeComparison as 'unspecified' | 'smaller' | 'larger' | undefined,
      },
      action: {
        addLabelIds: filter.action?.addLabelIds || undefined,
        removeLabelIds: filter.action?.removeLabelIds || undefined,
        forward: filter.action?.forward || undefined,
      },
    };
  } catch (error: any) {
    if (error instanceof UserError) throw error;
    if (error.code === 400) {
      throw new UserError(`Invalid filter configuration: ${error.message}`);
    }
    throw new Error(`Gmail API Error creating filter: ${error.message}`);
  }
}

// --- Delete Filter ---
export async function deleteFilter(gmail: Gmail, filterId: string): Promise<void> {
  try {
    await gmail.users.settings.filters.delete({
      userId: 'me',
      id: filterId,
    });
  } catch (error: any) {
    if (error.code === 404) {
      throw new UserError(`Filter not found (ID: ${filterId}).`);
    }
    throw new Error(`Gmail API Error deleting filter: ${error.message}`);
  }
}

// --- Create Filter from Template ---
export async function createFilterFromTemplate(
  gmail: Gmail,
  template: FilterTemplateType,
  params: FilterTemplateParams
): Promise<FilterInfo> {
  let criteria: FilterCriteriaArgs = {};
  let action: FilterActionArgs = {};

  // Build base action from common parameters
  if (params.labelIds && params.labelIds.length > 0) {
    action.addLabelIds = params.labelIds;
  }
  if (params.archive) {
    // Archive = remove from INBOX
    action.removeLabelIds = [...(action.removeLabelIds || []), 'INBOX'];
  }
  if (params.markAsRead) {
    // Mark as read = remove UNREAD label
    action.removeLabelIds = [...(action.removeLabelIds || []), 'UNREAD'];
  }
  if (params.markImportant) {
    // Mark as important = add IMPORTANT label
    action.addLabelIds = [...(action.addLabelIds || []), 'IMPORTANT'];
  }

  // Build criteria based on template type
  switch (template) {
    case 'fromSender':
      if (!params.senderEmail) {
        throw new UserError('fromSender template requires senderEmail parameter.');
      }
      criteria.from = params.senderEmail;
      break;

    case 'withSubject':
      if (!params.subjectText) {
        throw new UserError('withSubject template requires subjectText parameter.');
      }
      criteria.subject = params.subjectText;
      break;

    case 'withAttachments':
      criteria.hasAttachment = true;
      break;

    case 'largeEmails':
      if (!params.sizeInBytes) {
        throw new UserError('largeEmails template requires sizeInBytes parameter.');
      }
      criteria.size = params.sizeInBytes;
      criteria.sizeComparison = 'larger';
      break;

    case 'containingText':
      if (!params.searchText) {
        throw new UserError('containingText template requires searchText parameter.');
      }
      criteria.query = params.searchText;
      break;

    case 'mailingList':
      if (!params.listIdentifier) {
        throw new UserError('mailingList template requires listIdentifier parameter.');
      }
      // Use list: query operator for mailing lists
      criteria.query = `list:${params.listIdentifier}`;
      break;

    default:
      throw new UserError(`Unknown template type: ${template}`);
  }

  // Validate that at least one action is provided
  const hasAction = Object.values(action).some(v => v !== undefined && (Array.isArray(v) ? v.length > 0 : true));
  if (!hasAction) {
    throw new UserError('At least one action must be specified (labelIds, archive, markAsRead, or markImportant).');
  }

  return createFilter(gmail, criteria, action);
}

// --- Format Filter for Display ---
export function formatFilterForDisplay(filter: FilterInfo): string {
  let output = `Filter ID: ${filter.id}\n`;

  // Format criteria
  output += 'Criteria:\n';
  if (filter.criteria.from) output += `  From: ${filter.criteria.from}\n`;
  if (filter.criteria.to) output += `  To: ${filter.criteria.to}\n`;
  if (filter.criteria.subject) output += `  Subject: ${filter.criteria.subject}\n`;
  if (filter.criteria.query) output += `  Query: ${filter.criteria.query}\n`;
  if (filter.criteria.negatedQuery) output += `  Exclude: ${filter.criteria.negatedQuery}\n`;
  if (filter.criteria.hasAttachment) output += `  Has Attachment: Yes\n`;
  if (filter.criteria.excludeChats) output += `  Exclude Chats: Yes\n`;
  if (filter.criteria.size && filter.criteria.sizeComparison) {
    output += `  Size: ${filter.criteria.sizeComparison} ${filter.criteria.size} bytes\n`;
  }

  // Format actions
  output += 'Actions:\n';
  if (filter.action.addLabelIds && filter.action.addLabelIds.length > 0) {
    output += `  Add Labels: ${filter.action.addLabelIds.join(', ')}\n`;
  }
  if (filter.action.removeLabelIds && filter.action.removeLabelIds.length > 0) {
    output += `  Remove Labels: ${filter.action.removeLabelIds.join(', ')}\n`;
  }
  if (filter.action.forward) {
    output += `  Forward To: ${filter.action.forward}\n`;
  }

  return output;
}

// --- Format Multiple Filters for Display ---
export function formatFiltersForDisplay(filters: FilterInfo[]): string {
  if (filters.length === 0) {
    return 'No filters found.';
  }

  return filters.map((f, i) => `--- Filter ${i + 1} ---\n${formatFilterForDisplay(f)}`).join('\n');
}
