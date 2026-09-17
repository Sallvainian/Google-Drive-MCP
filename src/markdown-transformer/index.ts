// src/markdown-transformer/index.ts
//
// Markdown -> Google Docs write API.
//

import { docs_v1 } from 'googleapis';
import { convertMarkdownToRequests } from './markdownToDocs.js';
import type { ConversionOptions } from './markdownToDocs.js';
import { executeBatchUpdate } from '../googleDocsApiHelpers.js';

export { convertMarkdownToRequests } from './markdownToDocs.js';
export type { ConversionOptions } from './markdownToDocs.js';

interface InsertOptions {
  /** The 1-based document index where content should be inserted. Defaults to 1. */
  startIndex?: number;
  /** Target a specific tab by ID. */
  tabId?: string;
  /** Treat the first H1 (`# ...`) as a Google Docs TITLE instead of HEADING_1. */
  firstHeadingAsTitle?: boolean;
}

interface BatchUpdateMetadata {
  totalRequests: number;
  phases: {
    delete: { requests: number; apiCalls: number; elapsedMs: number };
    insert: { requests: number; apiCalls: number; elapsedMs: number };
    format: { requests: number; apiCalls: number; elapsedMs: number };
  };
  totalApiCalls: number;
  totalElapsedMs: number;
}

/** Debug metadata returned by insertMarkdown(). */
export interface InsertMarkdownResult {
  /** Total number of Google Docs API requests generated from the markdown. */
  totalRequests: number;
  /** Breakdown of requests by type (e.g. insertText, updateTextStyle, etc.). */
  requestsByType: Record<string, number>;
  /** Time spent parsing markdown and generating requests, in milliseconds. */
  parseElapsedMs: number;
  /** Metadata from the batch update execution (API call counts, per-phase timing). */
  batchUpdate: BatchUpdateMetadata;
  /** Total wall-clock time for the entire insertMarkdown operation, in milliseconds. */
  totalElapsedMs: number;
}

const emptyPhases = {
  delete: { requests: 0, apiCalls: 0, elapsedMs: 0 },
  insert: { requests: 0, apiCalls: 0, elapsedMs: 0 },
  format: { requests: 0, apiCalls: 0, elapsedMs: 0 },
};

/** Formats InsertMarkdownResult into a concise human-readable debug summary. */
export function formatInsertResult(result: InsertMarkdownResult): string {
  const lines: string[] = [];
  lines.push(`Markdown insert completed in ${result.totalElapsedMs}ms`);
  lines.push(`  Parse: ${result.parseElapsedMs}ms`);
  lines.push(
    `  Requests: ${result.totalRequests} total (${Object.entries(result.requestsByType)
      .map(([k, v]) => `${v} ${k}`)
      .join(', ')})`
  );
  lines.push(
    `  API calls: ${result.batchUpdate.totalApiCalls} batchUpdate calls in ${result.batchUpdate.totalElapsedMs}ms`
  );
  const { phases } = result.batchUpdate;
  if (phases.delete.requests > 0) {
    lines.push(
      `    Delete phase: ${phases.delete.requests} requests, ${phases.delete.apiCalls} calls, ${phases.delete.elapsedMs}ms`
    );
  }
  if (phases.insert.requests > 0) {
    lines.push(
      `    Insert phase: ${phases.insert.requests} requests, ${phases.insert.apiCalls} calls, ${phases.insert.elapsedMs}ms`
    );
  }
  if (phases.format.requests > 0) {
    lines.push(
      `    Format phase: ${phases.format.requests} requests, ${phases.format.apiCalls} calls, ${phases.format.elapsedMs}ms`
    );
  }
  return lines.join('\n');
}

/**
 * Converts markdown to Google Docs formatting and inserts it into a document.
 *
 * Handles the full pipeline: markdown parsing, request generation, and batch
 * execution against the Docs API. Callers never see raw API requests.
 *
 * @param docs - An authenticated Google Docs API client
 * @param documentId - The document ID
 * @param markdown - The markdown content to insert
 * @param options - Optional: startIndex (default 1), tabId
 * @returns Debug metadata about the operation (request counts, timing, API calls)
 */
export async function insertMarkdown(
  docs: docs_v1.Docs,
  documentId: string,
  markdown: string,
  options?: InsertOptions
): Promise<InsertMarkdownResult> {
  const overallStart = performance.now();
  const startIndex = options?.startIndex ?? 1;
  const tabId = options?.tabId;

  const parseStart = performance.now();
  const conversionOptions: ConversionOptions | undefined = options?.firstHeadingAsTitle
    ? { firstHeadingAsTitle: true }
    : undefined;
  const requests = convertMarkdownToRequests(markdown, startIndex, tabId, conversionOptions);
  const parseElapsedMs = Math.round(performance.now() - parseStart);

  const requestsByType: Record<string, number> = {};
  for (const r of requests) {
    const type = Object.keys(r)[0];
    requestsByType[type] = (requestsByType[type] || 0) + 1;
  }

  if (requests.length === 0) {
    return {
      totalRequests: 0,
      requestsByType,
      parseElapsedMs,
      batchUpdate: {
        totalRequests: 0,
        phases: { ...emptyPhases },
        totalApiCalls: 0,
        totalElapsedMs: 0,
      },
      totalElapsedMs: Math.round(performance.now() - overallStart),
    };
  }

  const batchStart = performance.now();
  await executeBatchUpdate(docs, documentId, requests);
  const totalElapsedMsBatch = Math.round(performance.now() - batchStart);
  const totalApiCalls = Math.ceil(requests.length / 50);

  return {
    totalRequests: requests.length,
    requestsByType,
    parseElapsedMs,
    batchUpdate: {
      totalRequests: requests.length,
      phases: {
        delete: { requests: 0, apiCalls: 0, elapsedMs: 0 },
        insert: { requests: requests.length, apiCalls: totalApiCalls, elapsedMs: totalElapsedMsBatch },
        format: { requests: 0, apiCalls: 0, elapsedMs: 0 },
      },
      totalApiCalls,
      totalElapsedMs: totalElapsedMsBatch,
    },
    totalElapsedMs: Math.round(performance.now() - overallStart),
  };
}
