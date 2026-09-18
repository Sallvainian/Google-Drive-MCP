// tests/docs-remaining-extras.test.js
import {
  FIND_TEXT_RANGE_FIELDS,
  TAB_VERIFY_FIELDS,
  buildTabsFieldMask,
} from '../dist/googleDocsApiHelpers.js';
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

function bodyDoc(text = 'Hello world target') {
  return {
    data: {
      body: {
        content: [
          {
            paragraph: {
              elements: [
                {
                  startIndex: 1,
                  endIndex: 1 + text.length,
                  textRun: { content: text },
                },
              ],
            },
          },
        ],
      },
    },
  };
}

function tabDoc(text = 'Hello world target', tabId = 'tab1') {
  return {
    data: {
      tabs: [
        {
          tabProperties: { tabId, title: 'Tab One', index: 0 },
          documentTab: {
            body: {
              content: [
                {
                  paragraph: {
                    elements: [
                      {
                        startIndex: 1,
                        endIndex: 1 + text.length,
                        textRun: { content: text },
                      },
                    ],
                  },
                },
              ],
            },
          },
        },
      ],
    },
  };
}

const docsGet = mock.fn(async () => bodyDoc());
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

const silentLog = {
  info() {},
  warn() {},
  error() {},
  debug() {},
};

function callTool(name, args) {
  const tool = tools.get(name);
  assert.ok(tool, `missing tool ${name}`);
  return tool.execute(args, { log: silentLog });
}

function parseTool(name, args) {
  const tool = tools.get(name);
  assert.ok(tool, `missing tool ${name}`);
  return tool.parameters.parse(args);
}

function getParams(callIndex = 0) {
  assert.ok(docsGet.mock.calls.length > callIndex, 'documents.get was not called enough times');
  return docsGet.mock.calls[callIndex].arguments[0];
}

function batchRequests(callIndex = 0) {
  assert.ok(
    docsBatchUpdate.mock.calls.length > callIndex,
    'documents.batchUpdate was not called enough times'
  );
  return docsBatchUpdate.mock.calls[callIndex].arguments[0].requestBody.requests;
}

describe('docs remaining extras registration', () => {
  it('registers extras and keeps empty insertTable; no findAndReplace', () => {
    assert.ok(tools.has('insertTableWithData'));
    assert.ok(tools.has('insertTable'));
    assert.ok(tools.has('insertSectionBreak'));
    assert.ok(tools.has('updateSectionStyle'));
    assert.ok(tools.has('modifyText'));
    assert.ok(tools.has('batchApplyTextStyle'));
    assert.equal(tools.has('findAndReplace'), false);
  });

  it('does not set includeTabsContent in extra tool bodies', () => {
    for (const name of [
      'insertTableWithData',
      'insertSectionBreak',
      'updateSectionStyle',
      'modifyText',
      'batchApplyTextStyle',
    ]) {
      const body = toolBody(serverSrc, name);
      assert.equal(body.includes('includeTabsContent'), false, name);
    }
  });
});

describe('insertTableWithData', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('inserts a 2x2 table, bolds the header, and keeps empty insertTable', async () => {
    const args = parseTool('insertTableWithData', {
      documentId: 'doc1',
      data: [
        ['Name', 'Age'],
        ['Ada', '36'],
      ],
      index: 1,
      hasHeaderRow: true,
    });
    const result = await callTool('insertTableWithData', args);
    const requests = batchRequests(0);
    assert.equal(requests[0].insertTable.rows, 2);
    assert.equal(requests[0].insertTable.columns, 2);
    assert.equal(requests[0].insertTable.location.index, 1);
    assert.equal(Object.hasOwn(requests[0].insertTable.location, 'tabId'), false);
    const insertTexts = requests.filter((r) => r.insertText);
    assert.deepStrictEqual(
      insertTexts.map((r) => r.insertText.location.index),
      [5, 11, 17, 22]
    );
    const headerBolds = requests.filter((r) => r.updateTextStyle?.textStyle?.bold === true);
    assert.deepStrictEqual(
      headerBolds.map((r) => [r.updateTextStyle.range.startIndex, r.updateTextStyle.range.endIndex]),
      [
        [5, 9],
        [11, 14],
      ]
    );
    assert.match(result, /Successfully inserted a 2x2 table with data at index 1/);
    assert.match(result, /Header row bolded/);
    const empty = parseTool('insertTable', { documentId: 'doc1', rows: 2, columns: 2, index: 1 });
    await callTool('insertTable', empty);
    const emptyReq = batchRequests(1);
    assert.deepStrictEqual(emptyReq[0].insertTable, {
      location: { index: 1 },
      rows: 2,
      columns: 2,
    });
  });

  it('rejects empty rows before writing', async () => {
    const args = parseTool('insertTableWithData', {
      documentId: 'doc1',
      data: [[]],
      index: 1,
    });
    await assert.rejects(
      () => callTool('insertTableWithData', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Table data must contain at least one non-empty row with at least one cell.'
        );
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('sets location.tabId after verifying the tab', async () => {
    const args = parseTool('insertTableWithData', {
      documentId: 'doc1',
      data: [['A']],
      index: 1,
      tabId: 'tab1',
    });
    const result = await callTool('insertTableWithData', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.strictEqual(batchRequests(0)[0].insertTable.location.tabId, 'tab1');
    assert.match(result, / in tab tab1/);
  });

  it('rejects a missing tab before writing', async () => {
    docsGet.mock.mockImplementation(async () => ({ data: { tabs: [] } }));
    const args = parseTool('insertTableWithData', {
      documentId: 'doc1',
      data: [['A']],
      index: 1,
      tabId: 'missing',
    });
    await assert.rejects(
      () => callTool('insertTableWithData', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });
});

describe('insertSectionBreak and updateSectionStyle', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('defaults to NEXT_PAGE and accepts CONTINUOUS', async () => {
    const next = parseTool('insertSectionBreak', { documentId: 'doc1', index: 5 });
    assert.strictEqual(next.sectionType, 'NEXT_PAGE');
    const result = await callTool('insertSectionBreak', next);
    assert.deepStrictEqual(batchRequests(0), [
      { insertSectionBreak: { location: { index: 5 }, sectionType: 'NEXT_PAGE' } },
    ]);
    assert.strictEqual(result, 'Successfully inserted NEXT_PAGE section break at index 5.');
    const cont = parseTool('insertSectionBreak', {
      documentId: 'doc1',
      index: 8,
      sectionType: 'CONTINUOUS',
    });
    const contResult = await callTool('insertSectionBreak', cont);
    assert.strictEqual(batchRequests(1)[0].insertSectionBreak.sectionType, 'CONTINUOUS');
    assert.strictEqual(contResult, 'Successfully inserted CONTINUOUS section break at index 8.');
  });

  it('sets location.tabId after verifying the tab', async () => {
    const args = parseTool('insertSectionBreak', {
      documentId: 'doc1',
      index: 5,
      tabId: 'tab1',
    });
    const result = await callTool('insertSectionBreak', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.strictEqual(batchRequests(0)[0].insertSectionBreak.location.tabId, 'tab1');
    assert.match(result, / in tab tab1/);
  });

  it('sets range.tabId on updateSectionStyle after verifying the tab', async () => {
    const args = parseTool('updateSectionStyle', {
      documentId: 'doc1',
      startIndex: 2,
      endIndex: 10,
      flipPageOrientation: true,
      tabId: 'tab1',
    });
    const result = await callTool('updateSectionStyle', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.strictEqual(batchRequests(0)[0].updateSectionStyle.range.tabId, 'tab1');
    assert.match(result, / in tab tab1/);
  });

  it('rejects a missing tab on insertSectionBreak before writing', async () => {
    docsGet.mock.mockImplementation(async () => ({ data: { tabs: [] } }));
    const args = parseTool('insertSectionBreak', {
      documentId: 'doc1',
      index: 5,
      tabId: 'missing',
    });
    await assert.rejects(
      () => callTool('insertSectionBreak', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('updates flipPageOrientation and rejects no-options before write', async () => {
    const args = parseTool('updateSectionStyle', {
      documentId: 'doc1',
      startIndex: 2,
      endIndex: 10,
      flipPageOrientation: true,
    });
    const result = await callTool('updateSectionStyle', args);
    const request = batchRequests(0)[0].updateSectionStyle;
    assert.strictEqual(request.sectionStyle.flipPageOrientation, true);
    assert.ok(request.fields.includes('flipPageOrientation'));
    assert.equal(Object.hasOwn(request.range, 'tabId'), false);
    assert.strictEqual(
      result,
      'Successfully updated section style (flipPageOrientation) for range 2-10.'
    );
    const parsed = parseTool('updateSectionStyle', {
      documentId: 'doc1',
      startIndex: 2,
      endIndex: 10,
    });
    await assert.rejects(
      () => callTool('updateSectionStyle', parsed),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /flipPageOrientation/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 1);
  });
});

describe('modifyText and batchApplyTextStyle tabId', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async (params) => {
      if (params.includeTabsContent) return tabDoc();
      return bodyDoc();
    });
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('uses findTextRange 5th arg (tab get) when tabId is set', async () => {
    const args = parseTool('modifyText', {
      documentId: 'doc1',
      target: { textToFind: 'world' },
      text: 'earth',
      tabId: 'tab1',
    });
    const result = await callTool('modifyText', args);
    assert.ok(docsGet.mock.calls.length >= 2);
    const findGet = docsGet.mock.calls.find((call) =>
      call.arguments[0].fields.includes('textRun(content)')
    );
    assert.ok(findGet);
    assert.strictEqual(findGet.arguments[0].includeTabsContent, true);
    assert.strictEqual(
      findGet.arguments[0].fields,
      buildTabsFieldMask(`documentTab(${FIND_TEXT_RANGE_FIELDS})`)
    );
    const requests = batchRequests(0);
    assert.ok(requests[0].deleteContentRange.range.tabId === 'tab1');
    assert.match(result, /replaced text/);
  });

  it('uses the 4-arg body get when tabId is omitted', async () => {
    const args = parseTool('modifyText', {
      documentId: 'doc1',
      target: { startIndex: 1, endIndex: 6 },
      style: { bold: true },
    });
    assert.equal(args.tabId, undefined);
    const result = await callTool('modifyText', args);
    assert.equal(docsGet.mock.calls.length, 0);
    const request = batchRequests(0)[0].updateTextStyle;
    assert.equal(Object.hasOwn(request.range, 'tabId'), false);
    assert.match(result, /applied formatting/);
  });

  it('finds text with the 4-arg body get when tabId is omitted', async () => {
    const args = parseTool('modifyText', {
      documentId: 'doc1',
      target: { textToFind: 'world' },
      text: 'earth',
    });
    const result = await callTool('modifyText', args);
    const findGet = docsGet.mock.calls.find((call) =>
      call.arguments[0].fields.includes('textRun(content)')
    );
    assert.ok(findGet);
    assert.equal(Object.hasOwn(findGet.arguments[0], 'includeTabsContent'), false);
    assert.strictEqual(findGet.arguments[0].fields, FIND_TEXT_RANGE_FIELDS);
    assert.equal(Object.hasOwn(batchRequests(0)[0].deleteContentRange.range, 'tabId'), false);
    assert.match(result, /replaced text/);
  });

  it('uses findTextRange 5th arg and sets updateTextStyle.range.tabId when tabId is set', async () => {
    const args = parseTool('batchApplyTextStyle', {
      documentId: 'doc1',
      tabId: 'tab1',
      operations: [{ target: { textToFind: 'world' }, style: { bold: true } }],
    });
    const result = await callTool('batchApplyTextStyle', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    const findGet = docsGet.mock.calls.find((call) =>
      call.arguments[0].fields.includes('textRun(content)')
    );
    assert.ok(findGet);
    assert.strictEqual(findGet.arguments[0].includeTabsContent, true);
    assert.strictEqual(
      findGet.arguments[0].fields,
      buildTabsFieldMask(`documentTab(${FIND_TEXT_RANGE_FIELDS})`)
    );
    const request = batchRequests(0)[0].updateTextStyle;
    assert.strictEqual(request.range.tabId, 'tab1');
    assert.match(result, /Applied 1\/1 text style operation\(s\)/);
  });

  it('skips unresolved find ops and reports applied/skipped', async () => {
    const args = parseTool('batchApplyTextStyle', {
      documentId: 'doc1',
      operations: [
        { target: { startIndex: 1, endIndex: 5 }, style: { bold: true } },
        { target: { textToFind: 'missing' }, style: { italic: true } },
      ],
    });
    const result = await callTool('batchApplyTextStyle', args);
    assert.equal(docsBatchUpdate.mock.calls.length, 1);
    assert.match(result, /Applied 1\/2 text style operation\(s\)/);
    assert.match(result, /Skipped/);
  });

  it('throws when every operation is skipped, without writing', async () => {
    const args = parseTool('batchApplyTextStyle', {
      documentId: 'doc1',
      operations: [{ target: { textToFind: 'missing' }, style: { bold: true } }],
    });
    await assert.rejects(
      () => callTool('batchApplyTextStyle', args),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.match(error.message, /No operations could be applied/);
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });
});
