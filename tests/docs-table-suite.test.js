// tests/docs-table-suite.test.js
import { TABLE_CONTENT_BASIC_BODY_FIELDS } from '../dist/googleDocsApiHelpers.js';
import { FastMCP, UserError } from 'fastmcp';
import { google } from 'googleapis';
import { registerHooks } from 'node:module';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { beforeEach, describe, it, mock } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');
const srcDir = join(repoRoot, 'src');
const serverSrc = readFileSync(join(srcDir, 'server.ts'), 'utf8');

function toolBody(source, toolName) {
  const start = source.search(new RegExp(`name:\\s*['"]${toolName}['"]`));
  assert.notEqual(start, -1, `missing tool ${toolName}`);
  const rest = source.slice(start);
  const next = rest.indexOf('\nserver.addTool({');
  return next === -1 ? rest : rest.slice(0, next);
}

function tableElement({ startIndex = 1, endIndex = 20, rows = [['A', 'B'], ['C', 'D']] } = {}) {
  return {
    startIndex,
    endIndex,
    table: {
      rows: rows.length,
      columns: rows[0]?.length ?? 0,
      tableRows: rows.map((row, rowIndex) => ({
        tableCells: row.map((text, columnIndex) => {
          const contentStart = startIndex + 4 + rowIndex * (1 + 2 * row.length) + 2 * columnIndex;
          const contentEnd = contentStart + text.length + 1;
          return {
            startIndex: contentStart - 1,
            endIndex: contentEnd + 1,
            content: [
              {
                paragraph: {
                  elements: [
                    {
                      startIndex: contentStart,
                      endIndex: contentEnd,
                      textRun: { content: `${text}\n` },
                    },
                  ],
                },
              },
            ],
          };
        }),
      })),
    },
  };
}

function bodyTablesDoc(tables = [tableElement()]) {
  return { data: { body: { content: tables } } };
}

function tabTablesDoc(tables = [tableElement()], tabId = 'tab1') {
  return {
    data: {
      tabs: [
        {
          tabProperties: { tabId, title: 'Tab One', index: 0 },
          documentTab: { body: { content: tables } },
        },
      ],
    },
  };
}

const docsGet = mock.fn(async () => bodyTablesDoc());
const docsBatchUpdate = mock.fn(async () => ({ data: {} }));

google.docs = () => ({
  documents: {
    get: docsGet,
    batchUpdate: docsBatchUpdate,
  },
});
google.drive = () => ({});
google.sheets = () => ({});
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

function getParams(callIndex = 0) {
  return docsGet.mock.calls[callIndex].arguments[0];
}

function batchRequests(callIndex = 0) {
  return docsBatchUpdate.mock.calls[callIndex].arguments[0].requestBody.requests;
}

describe('#114 table suite registration', () => {
  it('registers table tools and keeps editTableCell', () => {
    for (const name of [
      'listDocumentTables',
      'cloneTable',
      'replaceTableRowData',
      'updateTableBorders',
      'updateTableColumnWidth',
      'updateTableRowStyle',
      'updateTableCellStyle',
      'editTableCell',
    ]) {
      assert.ok(tools.has(name), name);
    }
  });

  it('does not set includeTabsContent in table tool bodies', () => {
    for (const name of [
      'listDocumentTables',
      'cloneTable',
      'replaceTableRowData',
      'updateTableBorders',
      'updateTableColumnWidth',
      'updateTableRowStyle',
      'updateTableCellStyle',
    ]) {
      assert.equal(toolBody(serverSrc, name).includes('includeTabsContent'), false, name);
    }
  });
});

describe('listDocumentTables', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
  });

  it('returns table:body:N ids when tabId is omitted', async () => {
    docsGet.mock.mockImplementation(async () => bodyTablesDoc());
    const args = parseTool('listDocumentTables', { documentId: 'doc1' });
    const result = JSON.parse(await callTool('listDocumentTables', args));
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      fields: TABLE_CONTENT_BASIC_BODY_FIELDS,
    });
    assert.equal(Object.hasOwn(getParams(0), 'includeTabsContent'), false);
    assert.strictEqual(result.tables[0].tableId, 'table:body:0');
    assert.strictEqual(result.tables[0].rowCount, 2);
    assert.strictEqual(result.tables[0].columnCount, 2);
  });

  it('returns table:<tabId>:N ids when tabId is set', async () => {
    docsGet.mock.mockImplementation(async () => tabTablesDoc());
    const args = parseTool('listDocumentTables', { documentId: 'doc1', tabId: 'tab1' });
    const result = JSON.parse(await callTool('listDocumentTables', args));
    assert.strictEqual(getParams(0).includeTabsContent, true);
    assert.strictEqual(result.tables[0].tableId, 'table:tab1:0');
  });
});

describe('cloneTable and replaceTableRowData', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => bodyTablesDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('clones a source table into the target index with copy flags default true', async () => {
    const args = parseTool('cloneTable', {
      documentId: 'target',
      sourceDocumentId: 'source',
      sourceTableId: 'table:body:0',
      index: 1,
    });
    assert.strictEqual(args.copyColumnWidths, true);
    assert.strictEqual(args.copyRowStyles, true);
    assert.strictEqual(args.copyCellStyles, true);
    const result = await callTool('cloneTable', args);
    assert.ok(docsBatchUpdate.mock.calls.length >= 1);
    assert.strictEqual(
      result,
      'Successfully cloned table:body:0 into target at index 1.'
    );
  });

  it('throws when the source table is missing, without inserting', async () => {
    docsGet.mock.mockImplementation(async () => ({ data: { body: { content: [] } } }));
    const args = parseTool('cloneTable', {
      documentId: 'target',
      sourceDocumentId: 'source',
      sourceTableId: 'table:body:0',
      index: 1,
    });
    await assert.rejects(
      () => callTool('cloneTable', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /was not found/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('throws when the source table is empty, without inserting', async () => {
    docsGet.mock.mockImplementation(async () => ({
      data: {
        body: {
          content: [
            {
              startIndex: 1,
              endIndex: 2,
              table: { rows: 0, columns: 0, tableRows: [] },
            },
          ],
        },
      },
    }));
    const args = parseTool('cloneTable', {
      documentId: 'target',
      sourceDocumentId: 'source',
      sourceTableId: 'table:body:0',
      index: 1,
    });
    await assert.rejects(
      () => callTool('cloneTable', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /is empty and cannot be cloned/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('throws when the cloned table cannot be re-located after insert', async () => {
    let getCount = 0;
    docsGet.mock.mockImplementation(async () => {
      getCount += 1;
      if (getCount === 1) return bodyTablesDoc();
      return { data: { body: { content: [] } } };
    });
    const args = parseTool('cloneTable', {
      documentId: 'target',
      sourceDocumentId: 'source',
      sourceTableId: 'table:body:0',
      index: 1,
    });
    await assert.rejects(
      () => callTool('cloneTable', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /could not be re-located safely/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 1);
  });

  it('replaces a whole row by tableId', async () => {
    const args = parseTool('replaceTableRowData', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowIndex: 0,
      values: ['X', 'Y'],
    });
    const result = await callTool('replaceTableRowData', args);
    assert.ok(docsBatchUpdate.mock.calls.length >= 1);
    assert.strictEqual(result, 'Successfully replaced row 0 in table table:body:0.');
  });

  it('rejects an out-of-bounds row before writing', async () => {
    const args = parseTool('replaceTableRowData', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowIndex: 9,
      values: ['X'],
    });
    await assert.rejects(
      () => callTool('replaceTableRowData', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /out of bounds/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });
});

describe('table chrome', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => bodyTablesDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('updates borders, column width, row style, and cell style at tableStartLocation', async () => {
    const borderArgs = parseTool('updateTableBorders', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowStart: 0,
      rowEnd: 1,
      columnStart: 0,
      columnEnd: 1,
      top: { widthPt: 1, dashStyle: 'SOLID' },
    });
    const borderResult = await callTool('updateTableBorders', borderArgs);
    assert.ok(batchRequests(0)[0].updateTableCellStyle.tableRange.tableCellLocation.tableStartLocation.index);
    assert.match(borderResult, /Successfully updated table borders/);

    const widthArgs = parseTool('updateTableColumnWidth', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      columnIndices: [0],
      widthPt: 72,
    });
    const widthResult = await callTool('updateTableColumnWidth', widthArgs);
    assert.strictEqual(
      batchRequests(1)[0].updateTableColumnProperties.tableStartLocation.index,
      1
    );
    assert.match(widthResult, /Successfully updated width/);

    const rowArgs = parseTool('updateTableRowStyle', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowIndices: [0],
      minRowHeightPt: 20,
    });
    const rowResult = await callTool('updateTableRowStyle', rowArgs);
    assert.match(rowResult, /Successfully updated row style/);

    const cellArgs = parseTool('updateTableCellStyle', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowStart: 0,
      rowEnd: 0,
      columnStart: 0,
      columnEnd: 0,
      contentAlignment: 'MIDDLE',
    });
    const cellResult = await callTool('updateTableCellStyle', cellArgs);
    assert.match(cellResult, /Successfully updated table cell style/);
  });

  it('rejects missing border options and unknown table ids before write', async () => {
    const noOpts = parseTool('updateTableBorders', {
      documentId: 'doc1',
      tableId: 'table:body:0',
      rowStart: 0,
      rowEnd: 0,
      columnStart: 0,
      columnEnd: 0,
    });
    await assert.rejects(
      () => callTool('updateTableBorders', noOpts),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'No border style options were provided.');
        return true;
      }
    );
    const missing = parseTool('updateTableColumnWidth', {
      documentId: 'doc1',
      tableId: 'table:body:9',
      columnIndices: [0],
      widthPt: 40,
    });
    await assert.rejects(
      () => callTool('updateTableColumnWidth', missing),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Table "table:body:9" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });
});
