// src/googleDriveApiHelpers.ts
import { drive_v3 } from 'googleapis';

/**
 * Returns a single-quoted Drive `q` literal with `\` then `'` escaped.
 */
export function driveQueryQuoted(value: string): string {
  return "'" + value.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
}

/**
 * Lists every permission on a file, including inherited shared-drive roles,
 * paging through nextPageToken so results are not truncated at 100 ACLs.
 */
export async function listAllFilePermissions(
  drive: drive_v3.Drive,
  fileId: string
): Promise<drive_v3.Schema$Permission[]> {
  const permissions: drive_v3.Schema$Permission[] = [];
  let pageToken: string | undefined;
  do {
    const response = await drive.permissions.list({
      fileId,
      fields: 'nextPageToken,permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)',
      supportsAllDrives: true,
      pageToken,
    });
    if (response.data.permissions) {
      permissions.push(...response.data.permissions);
    }
    pageToken = response.data.nextPageToken || undefined;
  } while (pageToken);
  return permissions;
}

export function formatFilePermissions(permissions: drive_v3.Schema$Permission[]): string {
  if (permissions.length === 0) {
    return 'No permissions found for this file.';
  }
  let result = `File has ${permissions.length} permission(s):\n\n`;
  permissions.forEach((perm, index) => {
    result += `${index + 1}. **${perm.role}** — `;
    if (perm.type === 'anyone') {
      result += 'Anyone with the link';
    } else if (perm.type === 'domain') {
      result += `Domain: ${perm.domain}`;
    } else {
      result += `${perm.displayName || 'Unknown'} (${perm.emailAddress || 'no email'})`;
    }
    result += ` [type: ${perm.type}]\n`;
    if (perm.permissionDetails && perm.permissionDetails.length > 0) {
      const details = perm.permissionDetails.map((detail) => {
        const origin = detail.inherited
          ? (detail.inheritedFrom ? `inherited from ${detail.inheritedFrom}` : 'inherited')
          : 'direct';
        return `${detail.role || 'unknown'} (${detail.permissionType || 'unknown'}, ${origin})`;
      }).join('; ');
      result += `   permissionDetails: ${details}\n`;
    }
  });
  return result;
}
