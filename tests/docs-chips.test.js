// tests/docs-chips.test.js
import { TAB_VERIFY_FIELDS } from '../dist/googleDocsApiHelpers.js';
import { FastMCP, UserError } from 'fastmcp';
import { google } from 'googleapis';
import { registerHooks } from 'node:module';
import assert from 'node:assert';
import { dirname, join } from 'node:path';
import { beforeEach, describe, it, mock } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');

function tabDoc(tabId = 'tab1') {
  return {
    data: {
      tabs: [
        {
          tabProperties: { tabId, title: 'Tab One', index: 0 },
          documentTab: { body: { content: [{ startIndex: 1, endIndex: 20 }] } },
        },
      ],
    },
  };
}

const docsGet = mock.fn(async () => tabDoc());
const docsBatchUpdate = mock.fn(async () => ({ data: {} }));

google.docs = () => ({
  documents: { get: docsGet, batchUpdate: docsBatchUpdate },
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
process.env.MCP_TOOL_GROUPS = 'docs-chips';

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

function batchRequests(callIndex = 0) {
  return docsBatchUpdate.mock.calls[callIndex].arguments[0].requestBody.requests;
}

describe('docs-chips catalog', () => {
  it('registers only the three chip tools', () => {
    assert.deepStrictEqual([...tools.keys()].sort(), [
      'insertDateChip',
      'insertPerson',
      'insertRichLink',
    ].sort());
  });
});

describe('insertDateChip', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('inserts a date chip and rejects invalid dates before write', async () => {
    const args = parseTool('insertDateChip', {
      documentId: 'doc1',
      index: 3,
      date: '2026-05-07T15:30:00+09:00',
    });
    const result = await callTool('insertDateChip', args);
    assert.ok(batchRequests(0)[0].insertDate);
    assert.strictEqual(result, 'Successfully inserted a date chip at index 3.');
    const bad = parseTool('insertDateChip', {
      documentId: 'doc1',
      index: 3,
      date: 'not-a-date',
    });
    await assert.rejects(
      () => callTool('insertDateChip', bad),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Invalid date/time value: "not-a-date"');
        return true;
      }
    );
    assert.equal(docsBatchUpdate.mock.calls.length, 1);
  });
});

describe('insertPerson and insertRichLink', () => {
  beforeEach(() => {
    docsGet.mock.resetCalls();
    docsBatchUpdate.mock.resetCalls();
    docsGet.mock.mockImplementation(async () => tabDoc());
    docsBatchUpdate.mock.mockImplementation(async () => ({ data: {} }));
  });

  it('inserts person and rich link chips, verifying tabId when set', async () => {
    const person = parseTool('insertPerson', {
      documentId: 'doc1',
      index: 4,
      email: 'ada@example.com',
      name: 'Ada',
      tabId: 'tab1',
    });
    const personResult = await callTool('insertPerson', person);
    assert.deepStrictEqual(docsGet.mock.calls[0].arguments[0].fields, TAB_VERIFY_FIELDS);
    assert.strictEqual(batchRequests(0)[0].insertPerson.personProperties.email, 'ada@example.com');
    assert.strictEqual(personResult, 'Successfully inserted a person chip at index 4 in tab tab1.');

    const link = parseTool('insertRichLink', {
      documentId: 'doc1',
      index: 5,
      uri: 'https://docs.google.com/document/d/abc',
    });
    const linkResult = await callTool('insertRichLink', link);
    assert.strictEqual(
      batchRequests(1)[0].insertRichLink.richLinkProperties.uri,
      'https://docs.google.com/document/d/abc'
    );
    assert.strictEqual(linkResult, 'Successfully inserted a rich link at index 5.');
  });
});
