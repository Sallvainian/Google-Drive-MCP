// tests/docs-tab-writes.test.js
import { TAB_VERIFY_FIELDS } from '../dist/googleDocsApiHelpers.js';
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

function tabDoc({
  tabId = 'tab1',
  documentTab,
  extraTabs = [],
} = {}) {
  const tab = {
    tabProperties: { tabId, title: 'Tab One', index: 0 },
  };
  if (documentTab !== undefined) {
    tab.documentTab = documentTab;
  } else {
    tab.documentTab = {
      body: { content: [{ startIndex: 1, endIndex: 20 }] },
    };
  }
  return {
    data: {
      tabs: [tab, ...extraTabs],
    },
  };
}

const docsGet = mock.fn(async () => tabDoc());
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

function parseTool(name, args) {
  const tool = tools.get(name);
  assert.ok(tool, `missing tool ${name}`);
  return tool.parameters.parse(args);
}

function parseAddTab(args) {
  return parseTool('addTab', args);
}

function parseRenameTab(args) {
  return parseTool('renameTab', args);
}

function parseReplaceAllText(args) {
  return parseTool('replaceAllText', args);
}

describe('tab write tools are registered', () => {
  it('registers addTab and renameTab and does not register findAndReplace', () => {
    assert.ok(tools.has('addTab'));
    assert.ok(tools.has('renameTab'));
    assert.equal(tools.has('findAndReplace'), false);
  });

  it('does not set includeTabsContent in addTab, renameTab, or replaceAllText bodies', () => {
    for (const name of ['addTab', 'renameTab', 'replaceAllText']) {
      const body = toolBody(serverSrc, name);
      assert.equal(body.includes('includeTabsContent'), false, name);
    }
  });
});

describe('addTab', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({
      data: {
        replies: [
          {
            addDocumentTab: {
              tabProperties: {
                tabId: 'new1',
                title: 'Notes',
                index: 1,
                nestingLevel: 0,
              },
            },
          },
        ],
      },
    }));
  });

  it('sends only title in tabProperties and returns JSON with the new tabId', async () => {
    const result = await callTool('addTab', parseAddTab({ documentId: 'doc1', title: 'Notes' }));
    assert.deepStrictEqual(batchRequests(0), [
      { addDocumentTab: { tabProperties: { title: 'Notes' } } },
    ]);
    assert.equal(docsGet.mock.calls.length, 0);
    const parsed = JSON.parse(result);
    assert.strictEqual(parsed.message, 'Successfully added new tab "Notes"');
    assert.strictEqual(parsed.tabId, 'new1');
    assert.strictEqual(parsed.title, 'Notes');
    assert.strictEqual(parsed.index, 1);
    assert.strictEqual(parsed.nestingLevel, 0);
  });

  it('verifies parentTabId then sets parent, index, and iconEmoji', async () => {
    docsBatchUpdate.mock.mockImplementation(async () => ({
      data: {
        replies: [
          {
            addDocumentTab: {
              tabProperties: {
                tabId: 'child1',
                title: 'Child',
                index: 0,
                parentTabId: 'tab1',
                nestingLevel: 1,
              },
            },
          },
        ],
      },
    }));
    const args = parseAddTab({
      documentId: 'doc1',
      title: 'Child',
      parentTabId: 'tab1',
      index: 0,
      iconEmoji: '📋',
    });
    assert.strictEqual(args.index, 0);
    const result = await callTool('addTab', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.deepStrictEqual(batchRequests(0), [
      {
        addDocumentTab: {
          tabProperties: {
            title: 'Child',
            parentTabId: 'tab1',
            index: 0,
            iconEmoji: '📋',
          },
        },
      },
    ]);
    const parsed = JSON.parse(result);
    assert.strictEqual(parsed.tabId, 'child1');
    assert.strictEqual(parsed.parentTabId, 'tab1');
    assert.strictEqual(parsed.nestingLevel, 1);
  });

  it('throws tab UserErrors for a missing or non-document parent', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () => callTool('addTab', parseAddTab({ documentId: 'doc1', parentTabId: 'missing' })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);

    docsGet.mock.mockImplementation(async () => tabDoc({ documentTab: null }));
    await assert.rejects(
      () => callTool('addTab', parseAddTab({ documentId: 'doc1', parentTabId: 'tab1' })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Tab "tab1" does not have content (may not be a document tab).'
        );
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('rejects index: -1 at the schema', () => {
    assert.throws(() => parseAddTab({ documentId: 'doc1', index: -1 }));
  });

  it('returns the fallback string when replies are empty or tabId is missing', async () => {
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
    const empty = await callTool('addTab', parseAddTab({ documentId: 'doc1' }));
    assert.deepStrictEqual(batchRequests(0), [{ addDocumentTab: { tabProperties: {} } }]);
    assert.strictEqual(
      empty,
      'Tab created successfully, but could not retrieve the new tab details.'
    );

    docsBatchUpdate.mock.resetCalls();
    docsBatchUpdate.mock.mockImplementation(async () => ({
      data: { replies: [{ addDocumentTab: { tabProperties: { title: 'Notes' } } }] },
    }));
    const missingId = await callTool('addTab', parseAddTab({ documentId: 'doc1' }));
    assert.strictEqual(
      missingId,
      'Tab created successfully, but could not retrieve the new tab details.'
    );
  });

  it('rethrows batch 404 from executeBatchUpdate', async () => {
    const notFound = new Error('not found');
    notFound.code = 404;
    docsBatchUpdate.mock.mockImplementation(async () => {
      throw notFound;
    });
    await assert.rejects(
      () => callTool('addTab', parseAddTab({ documentId: 'doc1', title: 'Notes' })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Document not found (ID: doc1). Check the ID.');
        return true;
      }
    );
  });

  it('wraps unexpected errors as Failed to add tab', async () => {
    docsGet.mock.mockImplementation(async () => {
      throw new Error('boom');
    });
    await assert.rejects(
      () => callTool('addTab', parseAddTab({ documentId: 'doc1', parentTabId: 'tab1' })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to add tab: boom');
        return true;
      }
    );
  });
});

describe('renameTab', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('updates title after verifying the tab', async () => {
    const result = await callTool('renameTab', parseRenameTab({
      documentId: 'doc1',
      tabId: 'tab1',
      newTitle: 'Renamed',
    }));
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.deepStrictEqual(batchRequests(0), [
      {
        updateDocumentTabProperties: {
          tabProperties: { tabId: 'tab1', title: 'Renamed' },
          fields: 'title',
        },
      },
    ]);
    assert.strictEqual(result, 'Successfully renamed tab "tab1" to "Renamed".');
  });

  it('throws tab UserErrors for a missing or non-document tab', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () =>
        callTool('renameTab', parseRenameTab({
          documentId: 'doc1',
          tabId: 'missing',
          newTitle: 'Renamed',
        })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);

    docsGet.mock.mockImplementation(async () => tabDoc({ documentTab: null }));
    await assert.rejects(
      () =>
        callTool('renameTab', parseRenameTab({
          documentId: 'doc1',
          tabId: 'tab1',
          newTitle: 'Renamed',
        })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Tab "tab1" does not have content (may not be a document tab).'
        );
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('rejects newTitle: \'\' at the schema', () => {
    assert.throws(() =>
      parseRenameTab({ documentId: 'doc1', tabId: 'tab1', newTitle: '' })
    );
  });

  it('wraps unexpected errors as Failed to rename tab', async () => {
    docsGet.mock.mockImplementation(async () => {
      throw new Error('boom');
    });
    await assert.rejects(
      () =>
        callTool('renameTab', parseRenameTab({
          documentId: 'doc1',
          tabId: 'tab1',
          newTitle: 'Renamed',
        })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to rename tab: boom');
        return true;
      }
    );
  });
});

describe('replaceAllText tabId', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({
      data: {
        replies: [{ replaceAllText: { occurrencesChanged: 2 } }],
      },
    }));
  });

  it('sets tabsCriteria.tabIds and matchCase true when tabId is set', async () => {
    const args = parseReplaceAllText({
      documentId: 'doc1',
      findText: 'old',
      replaceText: 'new',
      tabId: 'tab1',
    });
    assert.strictEqual(args.matchCase, true);
    const result = await callTool('replaceAllText', args);
    assert.deepStrictEqual(getParams(0), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    const request = batchRequests(0)[0].replaceAllText;
    assert.deepStrictEqual(request.tabsCriteria, { tabIds: ['tab1'] });
    assert.strictEqual(request.containsText.matchCase, true);
    assert.equal(Object.hasOwn(request, 'range'), false);
    assert.strictEqual(
      result,
      'Replaced 2 occurrence(s) of "old" with "new".'
    );
  });

  it('omits tabsCriteria and still uses matchCase true when tabId is omitted', async () => {
    const args = parseReplaceAllText({
      documentId: 'doc1',
      findText: 'old',
      replaceText: 'new',
    });
    assert.strictEqual(args.matchCase, true);
    assert.equal(args.tabId, undefined);
    const result = await callTool('replaceAllText', args);
    assert.equal(docsGet.mock.calls.length, 0);
    const request = batchRequests(0)[0].replaceAllText;
    assert.equal(Object.hasOwn(request, 'tabsCriteria'), false);
    assert.equal(request.tabsCriteria, undefined);
    assert.strictEqual(request.containsText.matchCase, true);
    assert.strictEqual(
      result,
      'Replaced 2 occurrence(s) of "old" with "new".'
    );
  });

  it('sends matchCase false when matchCase: false is parsed', async () => {
    const args = parseReplaceAllText({
      documentId: 'doc1',
      findText: 'old',
      replaceText: 'new',
      matchCase: false,
    });
    assert.strictEqual(args.matchCase, false);
    await callTool('replaceAllText', args);
    assert.strictEqual(batchRequests(0)[0].replaceAllText.containsText.matchCase, false);
  });

  it('throws tab UserErrors when tabId is missing or not a document tab', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () =>
        callTool('replaceAllText', parseReplaceAllText({
          documentId: 'doc1',
          findText: 'old',
          replaceText: 'new',
          tabId: 'missing',
        })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);

    docsGet.mock.mockImplementation(async () => tabDoc({ documentTab: null }));
    await assert.rejects(
      () =>
        callTool('replaceAllText', parseReplaceAllText({
          documentId: 'doc1',
          findText: 'old',
          replaceText: 'new',
          tabId: 'tab1',
        })),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Tab "tab1" does not have content (may not be a document tab).'
        );
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 0);
  });

  it('wraps unexpected errors as Failed to replace text', async () => {
    docsGet.mock.mockImplementation(async () => {
      throw new Error('boom');
    });
    await assert.rejects(
      () =>
        callTool('replaceAllText', {
          documentId: 'doc1',
          findText: 'old',
          replaceText: 'new',
          tabId: 'tab1',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to replace text: boom');
        return true;
      }
    );
  });
});
