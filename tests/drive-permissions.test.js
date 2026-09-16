// tests/drive-permissions.test.js
import { formatFilePermissions, listAllFilePermissions } from '../dist/googleDriveApiHelpers.js';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it, mock } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const serverSrc = readFileSync(join(srcDir, 'server.ts'), 'utf8');

function toolBody(source, toolName) {
  const start = source.search(new RegExp(`name:\\s*['"]${toolName}['"]`));
  assert.notEqual(start, -1, `missing tool ${toolName}`);
  const rest = source.slice(start);
  const next = rest.indexOf('\nserver.addTool({');
  return next === -1 ? rest : rest.slice(0, next);
}

describe('listAllFilePermissions', () => {
  it('requests permissionDetails and pages through nextPageToken', async () => {
    const list = mock.fn(async (params) => {
      if (!params.pageToken) {
        return {
          data: {
            nextPageToken: 'page2',
            permissions: [
              {
                id: 'p1',
                type: 'user',
                role: 'writer',
                emailAddress: 'a@example.com',
                displayName: 'A',
                permissionDetails: [
                  { role: 'writer', permissionType: 'file', inherited: false },
                ],
              },
            ],
          },
        };
      }
      return {
        data: {
          permissions: [
            {
              id: 'p2',
              type: 'user',
              role: 'reader',
              emailAddress: 'b@example.com',
              displayName: 'B',
              permissionDetails: [
                { role: 'reader', permissionType: 'member', inherited: true, inheritedFrom: 'drive1' },
              ],
            },
          ],
        },
      };
    });
    const drive = { permissions: { list } };

    const permissions = await listAllFilePermissions(drive, 'file1');

    assert.strictEqual(list.mock.calls.length, 2);
    assert.deepStrictEqual(list.mock.calls[0].arguments[0], {
      fileId: 'file1',
      fields: 'nextPageToken,permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)',
      supportsAllDrives: true,
      pageToken: undefined,
    });
    assert.deepStrictEqual(list.mock.calls[1].arguments[0], {
      fileId: 'file1',
      fields: 'nextPageToken,permissions(id,type,role,emailAddress,displayName,domain,permissionDetails)',
      supportsAllDrives: true,
      pageToken: 'page2',
    });
    assert.strictEqual(permissions.length, 2);
    assert.strictEqual(permissions[0].id, 'p1');
    assert.strictEqual(permissions[1].id, 'p2');
    assert.strictEqual(permissions[1].permissionDetails[0].inherited, true);
  });
});

describe('formatFilePermissions', () => {
  it('prints inherited shared-drive roles from permissionDetails', () => {
    const text = formatFilePermissions([
      {
        id: 'p1',
        type: 'user',
        role: 'writer',
        displayName: 'A',
        emailAddress: 'a@example.com',
        permissionDetails: [
          { role: 'writer', permissionType: 'file', inherited: false },
          { role: 'reader', permissionType: 'member', inherited: true, inheritedFrom: 'drive1' },
        ],
      },
    ]);
    assert.match(text, /permissionDetails: writer \(file, direct\); reader \(member, inherited from drive1\)/);
  });

  it('prints inherited without inheritedFrom and returns empty-list copy', () => {
    const withInherited = formatFilePermissions([
      {
        type: 'user',
        role: 'reader',
        displayName: 'B',
        emailAddress: 'b@example.com',
        permissionDetails: [
          { role: 'reader', permissionType: 'member', inherited: true },
        ],
      },
    ]);
    assert.match(withInherited, /permissionDetails: reader \(member, inherited\)/);
    assert.strictEqual(formatFilePermissions([]), 'No permissions found for this file.');
  });
});

describe('listFilePermissions tool source', () => {
  it('calls listAllFilePermissions and formatFilePermissions', () => {
    const body = toolBody(serverSrc, 'listFilePermissions');
    assert.ok(body.includes('DriveHelpers.listAllFilePermissions'));
    assert.ok(body.includes('DriveHelpers.formatFilePermissions'));
  });
});

describe('getFolderInfo mimeType', () => {
  it('includes mimeType in files.get fields', () => {
    const body = toolBody(serverSrc, 'getFolderInfo');
    assert.match(body, /fields:\s*'[^']*mimeType[^']*'/);
    assert.ok(body.includes("folder.mimeType !== 'application/vnd.google-apps.folder'"));
  });
});

describe('list/search owner line', () => {
  it('does not print Unknown when Drive omits owners', () => {
    assert.equal(serverSrc.includes("owners?.[0]?.displayName || 'Unknown'"), false);
    for (const toolName of [
      'listGoogleDocs',
      'searchGoogleDocs',
      'getRecentGoogleDocs',
      'listGoogleSheets',
      'listGoogleSlides',
    ]) {
      const body = toolBody(serverSrc, toolName);
      assert.match(body, /if \(owner\) \{/);
      assert.match(body, /Owner: \$\{owner\}/);
    }
    const folderBody = toolBody(serverSrc, 'listFolderContents');
    assert.match(folderBody, /Modified: \$\{modifiedDate\}\$\{owner \? ` by \$\{owner\}` : ''\}/);
  });
});
