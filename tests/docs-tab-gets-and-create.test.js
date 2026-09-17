// tests/docs-tab-gets-and-create.test.js
import {
  FIND_TEXT_RANGE_FIELDS,
  GET_PARAGRAPH_RANGE_FIELDS,
  GET_TABLE_CELL_RANGE_FIELDS,
  TAB_BODY_END_INDEX_FIELDS,
  TAB_LIST_FIELDS,
  TAB_LIST_WITH_CONTENT_FIELDS,
  TAB_READ_CONTENT_FIELDS,
  TAB_VERIFY_FIELDS,
  buildDocumentGetFields,
  buildTabsFieldMask,
  getDocumentTab,
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

const CREATED_DOC = {
  id: 'doc1',
  name: 'My Doc',
  webViewLink: 'https://docs.google.com/d/doc1',
};

function defaultTabDoc({ tabId = 'tab1', lists, extraParagraph } = {}) {
  const content = [
    {
      endIndex: 10,
      paragraph: {
        elements: [{ textRun: { content: 'Hello\n' } }],
        ...(extraParagraph || {}),
      },
    },
  ];
  return {
    data: {
      title: 'Doc Title',
      lists,
      tabs: [
        {
          tabProperties: { tabId, title: 'Tab One', index: 0 },
          documentTab: {
            body: { content },
            lists,
          },
        },
      ],
    },
  };
}

const docsGet = mock.fn(async () => defaultTabDoc());
const docsBatchUpdate = mock.fn(async () => ({ data: {} }));
const driveCreate = mock.fn(async () => ({ data: CREATED_DOC }));
const driveDelete = mock.fn(async () => ({}));

google.docs = () => ({
  documents: {
    get: docsGet,
    batchUpdate: docsBatchUpdate,
  },
});
google.drive = () => ({
  files: {
    create: driveCreate,
    delete: driveDelete,
  },
});
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

function lastGetParams() {
  assert.ok(docsGet.mock.calls.length >= 1, 'documents.get was not called');
  return docsGet.mock.calls[0].arguments[0];
}

describe('tab field mask builders', () => {
  it('does not change the locked body masks', () => {
    assert.strictEqual(
      FIND_TEXT_RANGE_FIELDS,
      'body(content(paragraph(elements(startIndex,endIndex,textRun(content))),table,sectionBreak,tableOfContents,startIndex,endIndex))'
    );
    assert.strictEqual(
      GET_PARAGRAPH_RANGE_FIELDS,
      'body(content(startIndex,endIndex,paragraph,table,sectionBreak,tableOfContents))'
    );
    assert.strictEqual(
      GET_TABLE_CELL_RANGE_FIELDS,
      'body(content(startIndex,endIndex,table(tableRows(tableCells(startIndex,endIndex,content(paragraph(elements(startIndex,endIndex))))))))'
    );
  });

  it('nests childTabs three deep and names tabProperties.tabId', () => {
    const mask = buildTabsFieldMask('documentTab(body(content(endIndex)))');
    assert.strictEqual((mask.match(/childTabs\(/g) || []).length, 3);
    assert.ok(mask.includes('tabProperties(tabId)'));
    assert.equal(mask.includes('tabProperties,'), false);
    assert.equal(mask.includes('tabs)'), false);
    assert.strictEqual(mask, TAB_VERIFY_FIELDS);
  });

  it('exports list, endIndex, and read-content masks including lists', () => {
    assert.ok(TAB_LIST_FIELDS.startsWith('title,tabs('));
    assert.equal(TAB_LIST_FIELDS === 'title,tabs', false);
    assert.strictEqual((TAB_LIST_FIELDS.match(/childTabs\(/g) || []).length, 3);
    assert.ok(TAB_LIST_FIELDS.includes('tabProperties(tabId'));
    assert.strictEqual((TAB_LIST_WITH_CONTENT_FIELDS.match(/childTabs\(/g) || []).length, 3);
    assert.ok(TAB_LIST_WITH_CONTENT_FIELDS.includes('tabProperties(tabId'));
    assert.ok(
      (TAB_LIST_WITH_CONTENT_FIELDS.match(/documentTab\(body\(content\(endIndex\)\)\)/g) || []).length > 1
    );
    assert.ok(TAB_BODY_END_INDEX_FIELDS.includes('documentTab(body(content(endIndex)))'));
    assert.ok(TAB_READ_CONTENT_FIELDS.includes('lists'));
    assert.ok(TAB_READ_CONTENT_FIELDS.includes('documentTab(body'));
    assert.equal(TAB_READ_CONTENT_FIELDS.includes('*'), false);
  });

  it('buildDocumentGetFields wraps only when tabId is set', () => {
    const bodyFields = 'body(content(endIndex))';
    assert.strictEqual(buildDocumentGetFields(bodyFields), bodyFields);
    assert.equal(buildDocumentGetFields(bodyFields).includes('tabs('), false);
    const withTab = buildDocumentGetFields(bodyFields, 't.abc');
    assert.ok(withTab.startsWith('tabs('));
    assert.ok(withTab.includes('documentTab(body(content(endIndex)))'));
    assert.equal(/^body\(/.test(withTab), false);
  });
});

describe('server.ts no longer contains the wide tab masks', () => {
  it('does not contain tabs(tabProperties,documentTab) or needsTabsContent ? \'*\'', () => {
    assert.equal(serverSrc.includes('tabs(tabProperties,documentTab)'), false);
    assert.equal(serverSrc.includes("needsTabsContent ? '*'"), false);
    assert.equal(serverSrc.includes('needsTabsContent ? "*"'), false);
  });

  it('createDocument rethrows UserError and does not swallow insert failure', () => {
    const body = toolBody(serverSrc, 'createDocument');
    assert.ok(body.includes('if (error instanceof UserError) throw error'));
    assert.ok(body.includes('it is currently EMPTY'));
    assert.ok(body.includes('do not create a new one'));
    assert.equal(body.includes('You can add content manually'), false);
    assert.ok(body.includes('docs.documents.batchUpdate'));
  });
});

describe('getDocumentTab', () => {
  it('requests the verify mask and suggestionsViewMode', async () => {
    const get = mock.fn(async () => defaultTabDoc());
    const tab = await getDocumentTab({ documents: { get } }, 'doc1', 'tab1');
    assert.strictEqual(tab.tabProperties.tabId, 'tab1');
    assert.deepStrictEqual(get.mock.calls[0].arguments[0], {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
  });

  it('finds a tab nested three levels down', async () => {
    const get = mock.fn(async () => ({
      data: {
        tabs: [
          {
            tabProperties: { tabId: 't0' },
            documentTab: {},
            childTabs: [
              {
                tabProperties: { tabId: 't1' },
                documentTab: {},
                childTabs: [
                  {
                    tabProperties: { tabId: 't2' },
                    documentTab: { body: { content: [] } },
                  },
                ],
              },
            ],
          },
        ],
      },
    }));
    const tab = await getDocumentTab({ documents: { get } }, 'doc1', 't2');
    assert.strictEqual(tab.tabProperties.tabId, 't2');
  });

  it('throws the existing UserError when the tab is missing', async () => {
    const get = mock.fn(async () => ({ data: { tabs: [] } }));
    await assert.rejects(
      () => getDocumentTab({ documents: { get } }, 'doc1', 'missing'),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
  });
});

describe('createDocument and tab-targeted documents.get', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    driveCreate.mock.resetCalls();
    driveDelete.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => defaultTabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
    driveCreate.mock.mockImplementation(async () => ({ data: CREATED_DOC }));
    driveDelete.mock.mockImplementation(async () => ({}));
  });

  it('returns a success string with id, link, and Initial content added when insert succeeds', async () => {
    const result = await callTool('createDocument', {
      title: 'My Doc',
      initialContent: 'hello',
    });
    assert.strictEqual(
      result,
      `Successfully created document "${CREATED_DOC.name}" (ID: ${CREATED_DOC.id})\nView Link: ${CREATED_DOC.webViewLink}\n\nInitial content added to document.`
    );
    assert.strictEqual(driveCreate.mock.calls.length, 1);
    assert.strictEqual(docsBatchUpdate.mock.calls.length, 1);
    assert.strictEqual(driveDelete.mock.calls.length, 0);
  });

  it('throws a UserError naming the empty created doc when insert fails', async () => {
    docsBatchUpdate.mock.mockImplementation(async () => {
      throw new Error('quota exceeded');
    });
    await assert.rejects(
      () => callTool('createDocument', { title: 'My Doc', initialContent: 'hello' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Document "My Doc" was created (id: doc1, url: https://docs.google.com/d/doc1) but inserting the initial content FAILED, so it is currently EMPTY. Append the content to this existing document (do not create a new one). Underlying error: quota exceeded'
        );
        assert.equal(error.message.includes('Failed to create document'), false);
        return true;
      }
    );
    assert.strictEqual(driveCreate.mock.calls.length, 1);
    assert.strictEqual(driveDelete.mock.calls.length, 0);
  });

  it('rethrows that insert-failure UserError from the outer catch unchanged', async () => {
    docsBatchUpdate.mock.mockImplementation(async () => {
      throw new Error('quota exceeded');
    });
    await assert.rejects(
      () => callTool('createDocument', { title: 'My Doc', initialContent: 'hello' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.equal(error.message.startsWith('Failed to create document'), false);
        assert.ok(error.message.includes('EMPTY'));
        return true;
      }
    );
  });

  it('insertText with tabId uses the verify mask and PREVIEW_WITHOUT_SUGGESTIONS', async () => {
    await callTool('insertText', {
      documentId: 'doc1',
      textToInsert: 'x',
      index: 1,
      tabId: 'tab1',
    });
    assert.deepStrictEqual(lastGetParams(), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
    assert.equal(lastGetParams().fields.includes('tabs(tabProperties,documentTab)'), false);
  });

  it('deleteRange with tabId uses the verify mask and PREVIEW_WITHOUT_SUGGESTIONS', async () => {
    await callTool('deleteRange', {
      documentId: 'doc1',
      startIndex: 1,
      endIndex: 2,
      tabId: 'tab1',
    });
    assert.deepStrictEqual(lastGetParams(), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_VERIFY_FIELDS,
    });
  });

  it('appendToGoogleDoc with tabId uses the tab endIndex mask, not bare tabs', async () => {
    await callTool('appendToGoogleDoc', {
      documentId: 'doc1',
      textToAppend: 'more',
      addNewlineIfNeeded: false,
      tabId: 'tab1',
    });
    assert.deepStrictEqual(lastGetParams(), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_BODY_END_INDEX_FIELDS,
    });
    assert.equal(lastGetParams().fields === 'tabs', false);
    assert.deepStrictEqual(
      docsBatchUpdate.mock.calls[0].arguments[0].requestBody.requests[0].insertText.location,
      { index: 9, tabId: 'tab1' }
    );
  });

  it('readGoogleDoc with tabId uses the content mask including lists, not *', async () => {
    await callTool('readGoogleDoc', {
      documentId: 'doc1',
      format: 'text',
      tabId: 'tab1',
    });
    const params = lastGetParams();
    assert.strictEqual(params.suggestionsViewMode, 'PREVIEW_WITHOUT_SUGGESTIONS');
    assert.strictEqual(params.fields, TAB_READ_CONTENT_FIELDS);
    assert.ok(params.fields.includes('lists'));
    assert.equal(params.fields === '*', false);
    assert.equal(params.fields.includes('tabs(tabProperties,documentTab)'), false);
  });

  it('listDocumentTabs does not use bare title,tabs and sets suggestionsViewMode', async () => {
    await callTool('listDocumentTabs', { documentId: 'doc1', includeContent: false });
    assert.deepStrictEqual(lastGetParams(), {
      documentId: 'doc1',
      includeTabsContent: true,
      suggestionsViewMode: 'PREVIEW_WITHOUT_SUGGESTIONS',
      fields: TAB_LIST_FIELDS,
    });
    assert.equal(lastGetParams().fields === 'title,tabs', false);

    docsGet.mock.resetCalls();
    const listed = await callTool('listDocumentTabs', { documentId: 'doc1', includeContent: true });
    assert.strictEqual(lastGetParams().fields, TAB_LIST_WITH_CONTENT_FIELDS);
    assert.strictEqual(lastGetParams().suggestionsViewMode, 'PREVIEW_WITHOUT_SUGGESTIONS');
    assert.ok(listed.includes('9 characters'));
    assert.equal(listed.includes('Empty'), false);
  });

  it('markdown readGoogleDoc with tabId still converts using tab body and lists', async () => {
    const lists = {
      k1: { listProperties: { nestingLevels: [{ glyphType: 'DECIMAL' }] } },
    };
    docsGet.mock.mockImplementation(async () =>
      defaultTabDoc({
        lists,
        extraParagraph: { bullet: { listId: 'k1', nestingLevel: 0 } },
      })
    );
    const md = await callTool('readGoogleDoc', {
      documentId: 'doc1',
      format: 'markdown',
      tabId: 'tab1',
    });
    assert.ok(md.includes('1. Hello'));
    assert.equal(md.includes('- Hello'), false);
  });

  it('insertText with a missing tabId throws the existing UserError', async () => {
    docsGet.mock.mockImplementation(async () => ({ data: { tabs: [] } }));
    await assert.rejects(
      () =>
        callTool('insertText', {
          documentId: 'doc1',
          textToInsert: 'x',
          index: 1,
          tabId: 'missing',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Tab with ID "missing" not found in document.');
        return true;
      }
    );
  });
});
