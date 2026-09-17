// tests/sheets-advanced.test.js
import { FastMCP, UserError } from 'fastmcp';
import { google } from 'googleapis';
import { registerHooks } from 'node:module';
import assert from 'node:assert';
import { dirname, join } from 'node:path';
import { beforeEach, describe, it, mock } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');

const sheetsGet = mock.fn(async () => ({
  data: {
    sheets: [
      {
        properties: { sheetId: 0, title: 'Sheet1' },
        tables: [{ tableId: 't1', name: 'People', range: { startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 }, columnProperties: [{ columnIndex: 0, columnName: 'Name' }] }],
        conditionalFormats: [{ booleanRule: { condition: { type: 'BLANK' }, format: { textFormat: { bold: true } } }, ranges: [{ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }] }],
      },
    ],
  },
}));
const sheetsBatchUpdate = mock.fn(async () => ({
  data: {
    replies: [
      { addSheet: { properties: { sheetId: 9, title: 'Copy' } } },
      { duplicateSheet: { properties: { sheetId: 8, title: 'Copy of Sheet1' } } },
      { addProtectedRange: { protectedRange: { protectedRangeId: 44 } } },
      { addChart: { chart: { chartId: 77 } } },
      { addTable: { table: { tableId: 't2', name: 'New', columnProperties: [] } } },
    ],
  },
}));
const valuesBatchUpdate = mock.fn(async () => ({
  data: { totalUpdatedCells: 4, totalUpdatedRows: 2, totalUpdatedColumns: 2, totalUpdatedSheets: 1 },
}));
const valuesAppend = mock.fn(async () => ({ data: { updates: { updatedRange: 'Sheet1!A3:B3' } } }));
const valuesGet = mock.fn(async () => ({ data: { values: [['Ada']] } }));
const sheetsCopyTo = mock.fn(async () => ({ data: { sheetId: 3, title: 'Copied' } }));
const commentsCreate = mock.fn(async () => ({ data: { id: 'c1' } }));

google.docs = () => ({});
google.drive = () => ({ comments: { create: commentsCreate } });
google.sheets = () => ({
  spreadsheets: {
    get: sheetsGet,
    batchUpdate: sheetsBatchUpdate,
    sheets: { copyTo: sheetsCopyTo },
    values: {
      batchUpdate: valuesBatchUpdate,
      append: valuesAppend,
      get: valuesGet,
    },
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
process.env.MCP_TOOL_GROUPS = 'sheets-advanced';

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

const ADVANCED_NAMES = [
  'batchWrite',
  'deleteSheet',
  'duplicateSheet',
  'copySheetTo',
  'renameSheet',
  'copyFormatting',
  'readCellFormat',
  'setCellBorders',
  'freezeRowsAndColumns',
  'setDropdownValidation',
  'setColumnWidths',
  'setRowHeights',
  'autoResizeColumns',
  'autoResizeRows',
  'protectRange',
  'addConditionalFormatting',
  'getConditionalFormatting',
  'deleteConditionalFormatting',
  'groupRows',
  'ungroupAllRows',
  'insertChart',
  'deleteChart',
  'createTable',
  'listTables',
  'getTable',
  'deleteTable',
  'updateTableRange',
  'appendTableRows',
  'createSheetsComment',
  'createSheetsCellNote',
];

describe('sheets-advanced catalog', () => {
  it('registers the 30 advanced names and not default sheets/docs tools', () => {
    assert.deepStrictEqual([...tools.keys()], ADVANCED_NAMES);
    assert.equal(tools.has('readSpreadsheet'), false);
    assert.equal(tools.has('formatSpreadsheetCells'), false);
    assert.equal(tools.has('insertDateChip'), false);
  });
});

describe('sheets-advanced I/O', () => {
  beforeEach(() => {
    sheetsGet.mock.resetCalls();
    sheetsBatchUpdate.mock.resetCalls();
    valuesBatchUpdate.mock.resetCalls();
    commentsCreate.mock.resetCalls();
    sheetsGet.mock.mockImplementation(async () => ({
      data: {
        sheets: [
          {
            properties: { sheetId: 0, title: 'Sheet1' },
            tables: [{
              tableId: 't1',
              name: 'People',
              range: { startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 },
              columnProperties: [{ columnIndex: 0, columnName: 'Name' }],
            }],
            conditionalFormats: [{
              booleanRule: { condition: { type: 'BLANK' }, format: { textFormat: { bold: true } } },
              ranges: [{ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }],
            }],
          },
        ],
      },
    }));
    sheetsBatchUpdate.mock.mockImplementation(async () => ({
      data: {
        replies: [
          { duplicateSheet: { properties: { sheetId: 8, title: 'Copy of Sheet1' } } },
          { addProtectedRange: { protectedRange: { protectedRangeId: 44 } } },
          { addChart: { chart: { chartId: 77 } } },
          { addTable: { table: { tableId: 't2', name: 'New', columnProperties: [] } } },
        ],
      },
    }));
  });

  it('batchWrite returns the upstream success template', async () => {
    const args = parseTool('batchWrite', {
      spreadsheetId: 's1',
      data: [{ range: 'A1:B2', values: [[1, 2], [3, 4]] }],
    });
    const result = await callTool('batchWrite', args);
    assert.strictEqual(
      result,
      'Successfully batch-wrote 4 cells (2 rows, 2 columns) across 1 sheet(s) in 1 range(s).'
    );
  });

  it('deleteSheet / duplicateSheet / copySheetTo / renameSheet succeed', async () => {
    assert.strictEqual(
      await callTool('deleteSheet', parseTool('deleteSheet', { spreadsheetId: 's1', sheetId: 0 })),
      'Successfully deleted sheet (ID: 0) from spreadsheet.'
    );
    assert.strictEqual(
      await callTool('duplicateSheet', parseTool('duplicateSheet', { spreadsheetId: 's1', sheetId: 0 })),
      'Successfully duplicated sheet as "Copy of Sheet1" (Sheet ID: 8).'
    );
    assert.strictEqual(
      await callTool('copySheetTo', parseTool('copySheetTo', {
        sourceSpreadsheetId: 's1',
        sheetId: 0,
        destinationSpreadsheetId: 's2',
      })),
      'Successfully copied sheet to destination as "Copied" (Sheet ID: 3).'
    );
    assert.strictEqual(
      await callTool('renameSheet', parseTool('renameSheet', {
        spreadsheetId: 's1',
        sheetId: 0,
        newName: 'Renamed',
      })),
      'Successfully renamed sheet (ID: 0) to "Renamed".'
    );
  });

  it('createTable and createSheetsCellNote use helpers', async () => {
    sheetsBatchUpdate.mock.mockImplementation(async () => ({
      data: { replies: [{ addTable: { table: { tableId: 't2', name: 'New', columnProperties: [] } } }] },
    }));
    const created = JSON.parse(await callTool('createTable', parseTool('createTable', {
      spreadsheetId: 's1',
      name: 'New',
      range: 'Sheet1!A1:B2',
    })));
    assert.strictEqual(created.message, 'Table "New" created successfully.');
    assert.strictEqual(
      await callTool('createSheetsCellNote', parseTool('createSheetsCellNote', {
        spreadsheetId: 's1',
        range: 'Sheet1!A1',
        content: 'note',
      })),
      'Cell note added successfully to range "Sheet1!A1".'
    );
  });

  it('createSheetsComment rejects cell+range before write', async () => {
    assert.throws(
      () => parseTool('createSheetsComment', {
        spreadsheetId: 's1',
        content: 'hi',
        cell: 'A1',
        range: 'A1:B2',
        sheetName: 'Sheet1',
      }),
      /either cell or range/
    );
    assert.equal(commentsCreate.mock.calls.length, 0);
  });
});
