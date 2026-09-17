// src/toolGroups.ts
export const TOOL_GROUPS = [
  'docs',
  'drive',
  'sheets',
  'slides',
  'gmail',
  'sheets-advanced',
  'docs-chips',
] as const;

export type ToolGroup = (typeof TOOL_GROUPS)[number];

export const DEFAULT_TOOL_GROUPS: readonly ToolGroup[] = [
  'docs',
  'drive',
  'sheets',
  'slides',
  'gmail',
];

const TOOL_GROUP_SET = new Set<string>(TOOL_GROUPS);

export function parseEnabledToolGroups(
  raw: string | undefined = process.env.MCP_TOOL_GROUPS,
): ToolGroup[] {
  if (!raw?.trim()) {
    return [...DEFAULT_TOOL_GROUPS];
  }

  const requested = raw
    .split(',')
    .map((group) => group.trim().toLowerCase())
    .filter(Boolean);

  if (requested.length === 0) {
    return [...DEFAULT_TOOL_GROUPS];
  }

  if (requested.includes('all')) {
    return [...TOOL_GROUPS];
  }

  const unknown = requested.filter((group) => !TOOL_GROUP_SET.has(group));
  if (unknown.length > 0) {
    throw new Error(
      `Unknown MCP_TOOL_GROUPS value(s): ${unknown.join(', ')}. Valid groups: ${TOOL_GROUPS.join(', ')}`,
    );
  }

  const selected = new Set(requested);
  return TOOL_GROUPS.filter((group) => selected.has(group));
}
