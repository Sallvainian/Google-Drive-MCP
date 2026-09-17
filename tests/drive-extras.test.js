// tests/drive-extras.test.js
import { FastMCP, UserError } from 'fastmcp';
import { google } from 'googleapis';
import { registerHooks } from 'node:module';
import assert from 'node:assert';
import { dirname, join } from 'node:path';
import { beforeEach, describe, it, mock } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');

const driveList = mock.fn(async () => ({
  data: {
    files: [
      {
        id: 'f1',
        name: 'Report',
        mimeType: 'application/pdf',
        size: '12',
        modifiedTime: '2026-01-01T00:00:00.000Z',
        createdTime: '2025-01-01T00:00:00.000Z',
        webViewLink: 'https://example.test/f1',
        owners: [{ displayName: 'Ada' }],
      },
    ],
  },
}));
const sheetsBatchUpdate = mock.fn(async () => ({ data: {} }));
const sheetsGet = mock.fn(async () => ({
  data: {
    sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }],
  },
}));

google.docs = () => ({});
google.drive = () => ({ files: { list: driveList } });
google.sheets = () => ({
  spreadsheets: {
    get: sheetsGet,
    batchUpdate: sheetsBatchUpdate,
  },
});
google.slides = () => ({});
google.gmail = () => ({});

const tools = new Map();
const originalAddTool = FastMCP.prototype.addTool;
FastMCP.prototype.addTool = function addTool(tool) {
  tools.set(tool.name, tool);
  return originalAddTool.call(this, tool);
};

const stubAuth = pathToFileURL(join(repoRoot, 'scripts', 'stub-authorize.js')).href;
const transportSrc = [
  'export function resolveHttpStreamBind(){return {host:"127.0.0.1",port:8787};}',
  'export function createBearerAuthenticate(){return async()=>({authenticated:true});}',
  'export async function startConfiguredTransport(){return;}',
].join('\n');
const transportUrl = `data:text/javascript;charset=utf-8,${encodeURIComponent(transportSrc)}`;

registerHooks({
  resolve(specifier, context, nextResolve) {
    const parent = context.parentURL || '';
    if (specifier === './auth.js' && parent.endsWith('/dist/server.js')) {
      return { shortCircuit: true, url: stubAuth };
    }
    if (specifier === './mcpTransport.js' && parent.endsWith('/dist/server.js')) {
      return { shortCircuit: true, url: transportUrl };
    }
    return nextResolve(specifier, context);
  },
});

process.env.MCP_TRANSPORT = 'httpStream';
delete process.env.MCP_HTTP_TOKEN;
delete process.env.MCP_TOOL_GROUPS;

await import('../dist/server.js');

const silentLog = { info() {}, warn() {}, error() {}, debug() {} };

function callTool(name, args) {
  const tool = tools.get(name);
  assert.ok(tool, `missing tool ${name}`);
  return tool.execute(args, { log: silentLog });
}

function parseTool(name, args) {
  return tools.get(name).parameters.parse(args);
}

describe('drive extras registration', () => {
  it('registers listDriveFiles and searchDriveFiles and keeps type-specific tools', () => {
    assert.ok(tools.has('listDriveFiles'));
    assert.ok(tools.has('searchDriveFiles'));
    assert.ok(tools.has('listGoogleDocs'));
    assert.ok(tools.has('searchGoogleDocs'));
    assert.equal(tools.has('setFilePermission'), false);
  });
});

describe('listDriveFiles', () => {
  beforeEach(() => {
    driveList.mock.resetCalls();
    driveList.mock.mockImplementation(async () => ({
      data: {
        files: [
          {
            id: 'f1',
            name: 'Report',
            mimeType: 'application/pdf',
            size: '12',
            modifiedTime: '2026-01-01T00:00:00.000Z',
            createdTime: '2025-01-01T00:00:00.000Z',
            webViewLink: 'https://example.test/f1',
            owners: [{ displayName: 'Ada' }],
          },
        ],
      },
    }));
  });

  it('resolves mime shortcuts, quotes folderId, and sets shared-drive flags', async () => {
    const args = parseTool('listDriveFiles', {
      mimeType: 'pdf',
      folderId: "root",
      maxResults: 5,
    });
    const result = JSON.parse(await callTool('listDriveFiles', args));
    const params = driveList.mock.calls[0].arguments[0];
    assert.match(params.q, /mimeType='application\/pdf'/);
    assert.match(params.q, /'root' in parents/);
    assert.strictEqual(params.supportsAllDrives, true);
    assert.strictEqual(params.includeItemsFromAllDrives, true);
    assert.equal(Object.hasOwn(params, 'corpora'), false);
    assert.equal(Object.hasOwn(params, 'driveId'), false);
    assert.strictEqual(result.total, 1);
    assert.strictEqual(result.files[0].id, 'f1');
  });

  it('rejects ownedByMe+sharedWithMe before listing', async () => {
    const args = parseTool('listDriveFiles', { ownedByMe: true, sharedWithMe: true });
    await assert.rejects(
      () => callTool('listDriveFiles', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'ownedByMe and sharedWithMe cannot both be true.');
        return true;
      }
    );
    assert.equal(driveList.mock.calls.length, 0);
  });

  it('maps 403 to a Drive permission UserError', async () => {
    driveList.mock.mockImplementation(async () => {
      const error = new Error('forbidden');
      error.code = 403;
      throw error;
    });
    await assert.rejects(
      () => callTool('listDriveFiles', parseTool('listDriveFiles', {})),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Permission denied. Make sure you have granted Google Drive access to the application.'
        );
        return true;
      }
    );
  });
});

describe('searchDriveFiles', () => {
  beforeEach(() => {
    driveList.mock.resetCalls();
    driveList.mock.mockImplementation(async () => ({
      data: {
        nextPageToken: 'page-2',
        files: [{ id: 'f2', name: 'Notes', mimeType: 'application/vnd.google-apps.document' }],
      },
    }));
  });

  it('drops orderBy for searchIn both and uses ancestors for folderId', async () => {
    const args = parseTool('searchDriveFiles', {
      query: "O'Reilly",
      searchIn: 'both',
      folderId: 'fld1',
    });
    const result = JSON.parse(await callTool('searchDriveFiles', args));
    const params = driveList.mock.calls[0].arguments[0];
    assert.equal(Object.hasOwn(params, 'orderBy'), false);
    assert.match(params.q, /fullText contains/);
    assert.ok(params.q.includes("O\\'Reilly"), params.q);
    assert.match(params.q, /'fld1' in ancestors/);
    assert.strictEqual(params.supportsAllDrives, true);
    assert.strictEqual(params.includeItemsFromAllDrives, true);
    assert.strictEqual(result.hasMore, true);
    assert.strictEqual(result.nextPageToken, 'page-2');
  });

  it('keeps orderBy when searchIn is name', async () => {
    const args = parseTool('searchDriveFiles', { query: 'notes', searchIn: 'name' });
    await callTool('searchDriveFiles', args);
    const params = driveList.mock.calls[0].arguments[0];
    assert.ok(params.orderBy);
    assert.equal(params.q.includes('fullText'), false);
  });

  it('maps 403 to a Drive permission UserError', async () => {
    driveList.mock.mockImplementation(async () => {
      const error = new Error('forbidden');
      error.code = 403;
      throw error;
    });
    await assert.rejects(
      () => callTool('searchDriveFiles', parseTool('searchDriveFiles', { query: 'notes' })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Permission denied. Make sure you have granted Google Drive access to the application.'
        );
        return true;
      }
    );
  });
});

describe('formatSpreadsheetCells wrap and numberFormat', () => {
  beforeEach(() => {
    sheetsGet.mock.resetCalls();
    sheetsBatchUpdate.mock.resetCalls();
    sheetsGet.mock.mockImplementation(async () => ({
      data: { sheets: [{ properties: { sheetId: 0, title: 'Sheet1' } }] },
    }));
    sheetsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('sends wrapStrategy and numberFormat on the existing sheets tool', async () => {
    const args = parseTool('formatSpreadsheetCells', {
      spreadsheetId: 's1',
      range: 'A1:B2',
      wrapStrategy: 'WRAP',
      numberFormat: { type: 'CURRENCY', pattern: '$#,##0.00' },
    });
    const result = await callTool('formatSpreadsheetCells', args);
    const request = sheetsBatchUpdate.mock.calls[0].arguments[0].requestBody.requests[0].repeatCell;
    assert.deepStrictEqual(request.range, {
      sheetId: 0,
      startRowIndex: 0,
      endRowIndex: 2,
      startColumnIndex: 0,
      endColumnIndex: 2,
    });
    assert.strictEqual(request.cell.userEnteredFormat.wrapStrategy, 'WRAP');
    assert.deepStrictEqual(request.cell.userEnteredFormat.numberFormat, {
      type: 'CURRENCY',
      pattern: '$#,##0.00',
    });
    assert.strictEqual(request.fields, 'userEnteredFormat(wrapStrategy,numberFormat)');
    assert.strictEqual(result, 'Successfully formatted cells A1:B2.');
  });
});
