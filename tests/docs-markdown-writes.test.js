// tests/docs-markdown-writes.test.js
import {
  SUGGESTIONS_VIEW_MODE,
  TAB_BODY_END_INDEX_FIELDS,
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

function bodyDoc({ endIndex = 20, content } = {}) {
  return {
    data: {
      body: {
        content: content || [{ startIndex: 1, endIndex }],
      },
    },
  };
}

function tabDoc({
  tabId = 'tab1',
  endIndex = 20,
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
      body: { content: [{ startIndex: 1, endIndex }] },
    };
  }
  return {
    data: {
      tabs: [tab, ...extraTabs],
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

function assertNoIncludeTabsContent(params) {
  assert.equal(Object.hasOwn(params, 'includeTabsContent'), false);
  assert.equal(params.includeTabsContent, undefined);
  assert.equal(String(params.fields || '').includes('includeTabsContent'), false);
}

const REPLACE_BODY_FIELDS = 'body(content(startIndex,endIndex,paragraph))';
const REPLACE_TAB_FIELDS = buildTabsFieldMask(
  'documentTab(body(content(startIndex,endIndex,paragraph)))'
);

function assertInsertsHaveTabId(requests, tabId) {
  const inserts = requests.filter((request) => request.insertText);
  assert.ok(inserts.length > 0, 'expected insertText requests');
  for (const request of inserts) {
    assert.strictEqual(request.insertText.location.tabId, tabId);
  }
}

describe('markdown write tools are registered', () => {
  it('does not register appendMarkdown or findAndReplace', () => {
    assert.ok(tools.has('replaceDocumentWithMarkdown'));
    assert.ok(tools.has('appendMarkdownToGoogleDoc'));
    assert.ok(tools.has('replaceRangeWithMarkdown'));
    assert.equal(tools.has('appendMarkdown'), false);
    assert.equal(tools.has('findAndReplace'), false);
  });

  it('does not use includeTabsContent or wide tab masks in the three tools', () => {
    for (const name of [
      'replaceDocumentWithMarkdown',
      'appendMarkdownToGoogleDoc',
      'replaceRangeWithMarkdown',
    ]) {
      const body = toolBody(serverSrc, name);
      assert.equal(body.includes('includeTabsContent'), false, name);
      assert.equal(body.includes('tabs(tabProperties,documentTab)'), false, name);
      assert.ok(body.includes('tabId'), name);
    }
  });
});

describe('replaceDocumentWithMarkdown', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => bodyDoc({ endIndex: 20 }));
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('deletes [1, last.endIndex-1), runs survivor cleanup, then inserts', async () => {
    let getCalls = 0;
    docsGet.mock.mockImplementation(async () => {
      getCalls += 1;
      if (getCalls === 1) {
        return bodyDoc({
          content: [
            {
              startIndex: 1,
              endIndex: 20,
              paragraph: { bullet: { listId: 'k1' } },
            },
          ],
        });
      }
      return bodyDoc({ endIndex: 2 });
    });

    const markdown = '# Hello';
    const result = await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown,
    });

    assert.ok(
      result.startsWith(
        `Successfully replaced document content with ${markdown.length} characters of markdown.`
      )
    );

    const firstGet = getParams(0);
    assert.strictEqual(firstGet.documentId, 'doc1');
    assert.strictEqual(firstGet.suggestionsViewMode, SUGGESTIONS_VIEW_MODE);
    assert.strictEqual(firstGet.fields, REPLACE_BODY_FIELDS);
    assert.ok(firstGet.fields.includes('paragraph'));
    assertNoIncludeTabsContent(firstGet);

    const deleteReqs = batchRequests(0);
    assert.deepStrictEqual(deleteReqs[0].deleteContentRange.range, {
      startIndex: 1,
      endIndex: 19,
    });

    const cleanupReqs = batchRequests(1);
    assert.ok(cleanupReqs[0].deleteParagraphBullets);
    assert.deepStrictEqual(cleanupReqs[0].deleteParagraphBullets.range, {
      startIndex: 1,
      endIndex: 2,
    });
    assert.ok(cleanupReqs[1].updateTextStyle);
    assert.strictEqual(
      cleanupReqs[1].updateTextStyle.fields,
      'underline,bold,italic,strikethrough,foregroundColor,backgroundColor'
    );

    const insertReqs = batchRequests(2);
    assert.ok(insertReqs.some((request) => request.insertText));
    assert.equal(
      insertReqs.some((request) => request.deleteParagraphBullets),
      false
    );
    assert.ok(
      insertReqs.some(
        (request) => request.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'TITLE'
      )
    );
  });

  it('defaults firstHeadingAsTitle to true when the field is omitted', () => {
    const tool = tools.get('replaceDocumentWithMarkdown');
    const parsed = tool.parameters.parse({ documentId: 'doc1', markdown: '# Hello' });
    assert.strictEqual(parsed.firstHeadingAsTitle, true);
  });

  it('preserveTitle starts the replace after the first paragraph', async () => {
    docsGet.mock.mockImplementation(async () =>
      bodyDoc({
        content: [
          { startIndex: 1, endIndex: 10, paragraph: { elements: [] } },
          { startIndex: 10, endIndex: 30 },
        ],
      })
    );

    await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown: 'Body',
      preserveTitle: true,
    });

    const firstGet = getParams(0);
    assert.strictEqual(firstGet.fields, REPLACE_BODY_FIELDS);
    assert.ok(firstGet.fields.includes('paragraph'));
    assert.deepStrictEqual(batchRequests(0)[0].deleteContentRange.range, {
      startIndex: 10,
      endIndex: 29,
    });
  });

  it('skips delete and cleanup on an empty body and inserts at startIndex', async () => {
    docsGet.mock.mockImplementation(async () => bodyDoc({ endIndex: 2 }));

    const result = await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown: 'Hi',
    });

    assert.ok(result.startsWith('Successfully replaced document content with 2 characters of markdown.'));
    assert.equal(
      docsBatchUpdate.mock.calls.some((call) =>
        call.arguments[0].requestBody.requests.some((request) => request.deleteContentRange)
      ),
      false
    );
    assert.equal(
      docsBatchUpdate.mock.calls.some((call) =>
        call.arguments[0].requestBody.requests.some((request) => request.deleteParagraphBullets)
      ),
      false
    );
    const insertReqs = batchRequests(0);
    assert.ok(insertReqs.some((request) => request.insertText));
  });

  it('uses TAB_BODY_RANGE_FIELDS and SVM without includeTabsContent for tabId', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ endIndex: 12 }));

    await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown: 'Hi',
      tabId: 'tab1',
    });

    const firstGet = getParams(0);
    assert.strictEqual(firstGet.fields, REPLACE_TAB_FIELDS);
    assert.ok(firstGet.fields.includes('paragraph'));
    assert.strictEqual(firstGet.suggestionsViewMode, SUGGESTIONS_VIEW_MODE);
    assertNoIncludeTabsContent(firstGet);
    assert.deepStrictEqual(batchRequests(0)[0].deleteContentRange.range, {
      startIndex: 1,
      endIndex: 11,
      tabId: 'tab1',
    });
    const cleanupReqs = batchRequests(1);
    assert.strictEqual(cleanupReqs[0].deleteParagraphBullets.range.tabId, 'tab1');
    assert.strictEqual(cleanupReqs[1].updateTextStyle.range.tabId, 'tab1');
    assertInsertsHaveTabId(batchRequests(2), 'tab1');
  });

  it('throws when tabId is missing', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () =>
        callTool('replaceDocumentWithMarkdown', {
          documentId: 'doc1',
          markdown: 'Hi',
          tabId: 'missing',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
  });

  it('throws when the tab has no documentTab', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ documentTab: null }));
    await assert.rejects(
      () =>
        callTool('replaceDocumentWithMarkdown', {
          documentId: 'doc1',
          markdown: 'Hi',
          tabId: 'tab1',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Tab "tab1" does not have content (may not be a document tab).'
        );
        return true;
      }
    );
  });

  it('throws when body.content is missing', async () => {
    docsGet.mock.mockImplementation(async () => ({ data: { body: {} } }));
    await assert.rejects(
      () =>
        callTool('replaceDocumentWithMarkdown', {
          documentId: 'doc1',
          markdown: 'Hi',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'No content found in document/tab');
        return true;
      }
    );
  });

  it('wraps unexpected errors as Failed to apply markdown', async () => {
    docsGet.mock.mockImplementation(async () => {
      throw new Error('boom');
    });
    await assert.rejects(
      () =>
        callTool('replaceDocumentWithMarkdown', {
          documentId: 'doc1',
          markdown: 'Hi',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to apply markdown: boom');
        return true;
      }
    );
  });

  it('rethrows batch 404 from executeBatchUpdate', async () => {
    const notFound = new Error('not found');
    notFound.code = 404;
    docsBatchUpdate.mock.mockImplementation(async () => {
      throw notFound;
    });
    await assert.rejects(
      () =>
        callTool('replaceDocumentWithMarkdown', {
          documentId: 'doc1',
          markdown: 'Hi',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Document not found (ID: doc1). Check the ID.');
        return true;
      }
    );
  });
});

describe('appendMarkdownToGoogleDoc', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => bodyDoc({ endIndex: 10 }));
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('inserts spacing then markdown at startIndex+2 when the body is not empty', async () => {
    const markdown = 'More';
    const result = await callTool('appendMarkdownToGoogleDoc', {
      documentId: 'doc1',
      markdown,
      addNewlineIfNeeded: true,
    });

    assert.ok(
      result.startsWith(`Successfully appended ${markdown.length} characters of markdown.`)
    );

    const firstGet = getParams(0);
    assert.strictEqual(firstGet.fields, 'body(content(endIndex))');
    assert.strictEqual(firstGet.suggestionsViewMode, SUGGESTIONS_VIEW_MODE);
    assertNoIncludeTabsContent(firstGet);

    const spacing = batchRequests(0)[0].insertText;
    assert.deepStrictEqual(spacing.location, { index: 9 });
    assert.strictEqual(spacing.text, '\n\n');

    const insertReqs = batchRequests(1);
    const firstInsert = insertReqs.find((request) => request.insertText);
    assert.strictEqual(firstInsert.insertText.location.index, 11);
  });

  it('uses TAB_BODY_END_INDEX_FIELDS for tabId and puts tabId on the spacing insert', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ endIndex: 10 }));
    await callTool('appendMarkdownToGoogleDoc', {
      documentId: 'doc1',
      markdown: 'More',
      tabId: 'tab1',
      addNewlineIfNeeded: true,
    });

    const firstGet = getParams(0);
    assert.strictEqual(firstGet.fields, TAB_BODY_END_INDEX_FIELDS);
    assert.strictEqual(firstGet.suggestionsViewMode, SUGGESTIONS_VIEW_MODE);
    assertNoIncludeTabsContent(firstGet);
    assert.deepStrictEqual(batchRequests(0)[0].insertText.location, {
      index: 9,
      tabId: 'tab1',
    });
    assertInsertsHaveTabId(batchRequests(1), 'tab1');
  });

  it('does not insert spacing when addNewlineIfNeeded is false', async () => {
    const markdown = 'More';
    await callTool('appendMarkdownToGoogleDoc', {
      documentId: 'doc1',
      markdown,
      addNewlineIfNeeded: false,
    });
    assert.equal(
      docsBatchUpdate.mock.calls.some((call) =>
        call.arguments[0].requestBody.requests.some(
          (request) => request.insertText && request.insertText.text === '\n\n'
        )
      ),
      false
    );
    const firstInsert = batchRequests(0).find((request) => request.insertText);
    assert.strictEqual(firstInsert.insertText.location.index, 9);
  });

  it('throws the tab and no-content errors', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () =>
        callTool('appendMarkdownToGoogleDoc', {
          documentId: 'doc1',
          markdown: 'Hi',
          tabId: 'missing',
        }),
      (error) => {
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );

    docsGet.mock.mockImplementation(async () => tabDoc({ documentTab: null }));
    await assert.rejects(
      () =>
        callTool('appendMarkdownToGoogleDoc', {
          documentId: 'doc1',
          markdown: 'Hi',
          tabId: 'tab1',
        }),
      (error) => {
        assert.strictEqual(
          error.message,
          'Tab "tab1" does not have content (may not be a document tab).'
        );
        return true;
      }
    );

    docsGet.mock.mockImplementation(async () => ({ data: { body: {} } }));
    await assert.rejects(
      () =>
        callTool('appendMarkdownToGoogleDoc', {
          documentId: 'doc1',
          markdown: 'Hi',
        }),
      (error) => {
        assert.strictEqual(error.message, 'No content found in document/tab');
        return true;
      }
    );
  });

  it('wraps unexpected errors as Failed to append markdown', async () => {
    docsGet.mock.mockImplementation(async () => {
      throw new Error('nope');
    });
    await assert.rejects(
      () =>
        callTool('appendMarkdownToGoogleDoc', {
          documentId: 'doc1',
          markdown: 'Hi',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to append markdown: nope');
        return true;
      }
    );
  });
});

describe('replaceRangeWithMarkdown', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('deletes the range then inserts at startIndex with no survivor cleanup', async () => {
    const markdown = 'New';
    const result = await callTool('replaceRangeWithMarkdown', {
      documentId: 'doc1',
      startIndex: 5,
      endIndex: 12,
      markdown,
    });

    assert.ok(
      result.startsWith(
        `Successfully replaced range 5-12 with ${markdown.length} characters of markdown.`
      )
    );
    assert.strictEqual(docsGet.mock.calls.length, 0);
    assert.deepStrictEqual(batchRequests(0)[0].deleteContentRange.range, {
      startIndex: 5,
      endIndex: 12,
    });
    assert.equal(
      docsBatchUpdate.mock.calls.some((call) =>
        call.arguments[0].requestBody.requests.some((request) => request.deleteParagraphBullets)
      ),
      false
    );
    const insertReqs = batchRequests(1);
    const firstInsert = insertReqs.find((request) => request.insertText);
    assert.strictEqual(firstInsert.insertText.location.index, 5);
  });

  it('puts tabId on delete and insert after verifying the tab', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ endIndex: 20 }));
    await callTool('replaceRangeWithMarkdown', {
      documentId: 'doc1',
      startIndex: 5,
      endIndex: 12,
      markdown: 'New',
      tabId: 'tab1',
    });
    assert.strictEqual(getParams(0).fields, TAB_VERIFY_FIELDS);
    assertNoIncludeTabsContent(getParams(0));
    assert.deepStrictEqual(batchRequests(0)[0].deleteContentRange.range, {
      startIndex: 5,
      endIndex: 12,
      tabId: 'tab1',
    });
    assertInsertsHaveTabId(batchRequests(1), 'tab1');
  });

  it('throws when tabId is missing', async () => {
    docsGet.mock.mockImplementation(async () => tabDoc({ tabId: 'other' }));
    await assert.rejects(
      () =>
        callTool('replaceRangeWithMarkdown', {
          documentId: 'doc1',
          startIndex: 5,
          endIndex: 12,
          markdown: 'New',
          tabId: 'missing',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
    assert.strictEqual(docsBatchUpdate.mock.calls.length, 0);
  });

  it('throws when endIndex is not greater than startIndex', async () => {
    await assert.rejects(
      () =>
        callTool('replaceRangeWithMarkdown', {
          documentId: 'doc1',
          startIndex: 5,
          endIndex: 5,
          markdown: 'x',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'endIndex must be greater than startIndex');
        return true;
      }
    );
  });

  it('keeps the Failed to replace range with markdown catch-all in source', () => {
    const body = toolBody(serverSrc, 'replaceRangeWithMarkdown');
    assert.ok(body.includes('Failed to replace range with markdown:'));
    assert.ok(body.includes('endIndex must be greater than startIndex'));
  });

  it('rethrows batch 403 from executeBatchUpdate', async () => {
    const denied = new Error('denied');
    denied.code = 403;
    docsBatchUpdate.mock.mockImplementation(async () => {
      throw denied;
    });
    await assert.rejects(
      () =>
        callTool('replaceRangeWithMarkdown', {
          documentId: 'doc1',
          startIndex: 5,
          endIndex: 12,
          markdown: 'x',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Permission denied for document (ID: doc1). Ensure the authenticated user has edit access.'
        );
        return true;
      }
    );
  });
});

describe('whitespace and image markdown', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => bodyDoc({ endIndex: 2 }));
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('whitespace markdown does not insert', async () => {
    await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown: '   \n\n   ',
    });
    assert.equal(
      docsBatchUpdate.mock.calls.some((call) =>
        call.arguments[0].requestBody.requests.some((request) => request.insertText)
      ),
      false
    );
  });

  it('image markdown is omitted from insert requests', async () => {
    await callTool('replaceDocumentWithMarkdown', {
      documentId: 'doc1',
      markdown: '![alt](https://example.com/img.png)',
    });
    const serialized = JSON.stringify(
      docsBatchUpdate.mock.calls.map((call) => call.arguments[0].requestBody.requests)
    );
    assert.equal(serialized.includes('insertInlineImage'), false);
    assert.equal(serialized.includes('https://example.com/img.png'), false);
  });
});
