// src/googleDriveApiHelpers.ts

/**
 * Returns a single-quoted Drive `q` literal with `\` then `'` escaped.
 */
export function driveQueryQuoted(value: string): string {
  return "'" + value.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
}
