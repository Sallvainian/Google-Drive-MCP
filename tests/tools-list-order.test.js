// tests/tools-list-order.test.js
import {
  assertCreateFromTemplateReplacementsSchema,
  listToolsOverHttpStream,
  listToolsOverStdio,
} from './live-fastmcp-helpers.js';
import assert from 'node:assert';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

const ADD_TOOL_NAMES = [
  'readGoogleDoc',
  'listDocumentTabs',
  'appendToGoogleDoc',
  'insertText',
  'replaceAllText',
  'deleteRange',
  'applyTextStyle',
  'applyParagraphStyle',
  'insertTable',
  'editTableCell',
  'insertDocTableRow',
  'deleteDocTableRow',
  'insertPageBreak',
  'insertImageFromUrl',
  'insertLocalImage',
  'fixListFormatting',
  'listComments',
  'getComment',
  'addComment',
  'replyToComment',
  'resolveComment',
  'deleteComment',
  'findElement',
  'formatMatchingText',
  'listGoogleDocs',
  'searchGoogleDocs',
  'getRecentGoogleDocs',
  'getDocumentInfo',
  'createFolder',
  'listFolderContents',
  'listAllFolders',
  'getFolderInfo',
  'moveFile',
  'copyFile',
  'renameFile',
  'deleteFile',
  'uploadFile',
  'downloadFile',
  'shareFile',
  'makeFilePublic',
  'listFilePermissions',
  'createDocument',
  'createFromTemplate',
  'readSpreadsheet',
  'writeSpreadsheet',
  'appendSpreadsheetRows',
  'clearSpreadsheetRange',
  'getSpreadsheetInfo',
  'addSpreadsheetSheet',
  'formatSpreadsheetCells',
  'createSpreadsheet',
  'listGoogleSheets',
  'createFormattedDocument',
  'insertFormattedContent',
  'replaceDocumentContent',
  'updateDocumentSection',
  'listGoogleSlides',
  'getPresentation',
  'listSlides',
  'getSlide',
  'mapSlide',
  'createPresentation',
  'addSlide',
  'duplicateSlide',
  'addTextBox',
  'addShape',
  'addImage',
  'addTable',
  'editSlideTableCell',
  'deleteSlide',
  'deleteElement',
  'updateSpeakerNotes',
  'moveSlide',
  'insertTextInElement',
  'send_email',
  'draft_email',
  'read_email',
  'search_emails',
  'modify_email',
  'delete_email',
  'download_attachment',
  'list_email_labels',
  'create_label',
  'update_label',
  'delete_label',
  'get_or_create_label',
  'batch_modify_emails',
  'batch_delete_emails',
  'create_filter',
  'list_filters',
  'get_filter',
  'delete_filter',
  'create_filter_from_template',
  'get_thread',
  'list_threads',
  'reply_to_email',
  'forward_email',
  'trash_email',
  'untrash_email',
  'archive_email',
  'mark_as_read',
  'mark_as_unread',
  'get_user_profile',
  'list_drafts',
  'get_draft',
  'update_draft',
  'delete_draft',
  'send_draft',
];

function addToolNames(source) {
  const names = [];
  const marker = 'server.addTool({';
  let from = 0;
  while (true) {
    const start = source.indexOf(marker, from);
    if (start === -1) {
      break;
    }
    const rest = source.slice(start);
    const next = rest.indexOf('\nserver.addTool({', marker.length);
    const body = next === -1 ? rest : rest.slice(0, next);
    const match = body.match(/name:\s*['"]([^'"]+)['"]/);
    assert.notEqual(match, null, 'addTool is missing a name');
    names.push(match[1]);
    from = start + marker.length;
  }
  return names;
}

describe('tools/list registration order', () => {
  it('lists 108 unique addTool names in file order, first readGoogleDoc last send_draft', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    const names = addToolNames(source);
    assert.deepStrictEqual(names, ADD_TOOL_NAMES);
  });

  it('does not put ttlMs or cacheScope in src/', () => {
    for (const file of readdirSync(srcDir)) {
      const full = join(srcDir, file);
      if (!statSync(full).isFile()) continue;
      const source = readFileSync(full, 'utf8');
      assert.equal(source.includes('ttlMs'), false, `${file} contains ttlMs`);
      assert.equal(source.includes('cacheScope'), false, `${file} contains cacheScope`);
    }
  });
});

function assertLiveToolsList(listed, expectedNames) {
  const names = listed.tools.map((tool) => tool.name);
  assert.deepStrictEqual(names, expectedNames);
  for (const response of listed.listResults) {
    assert.deepStrictEqual(Object.keys(response.result), ['tools']);
  }
  const create = listed.tools.find((tool) => tool.name === 'createFromTemplate');
  assert.notEqual(create, undefined);
  assertCreateFromTemplateReplacementsSchema(create.inputSchema.properties.replacements);
}

describe('live FastMCP tools/list', () => {
  it(
    'stdio tools/list is byte-identical across two process starts',
    { timeout: 300000 },
    async () => {
      const first = await listToolsOverStdio();
      const second = await listToolsOverStdio();
      assertLiveToolsList(first, ADD_TOOL_NAMES);
      assertLiveToolsList(second, ADD_TOOL_NAMES);
      assert.deepStrictEqual(
        first.tools.map((tool) => tool.name),
        second.tools.map((tool) => tool.name),
      );
    },
  );

  it(
    'httpStream tools/list matches ADD_TOOL_NAMES and live replacements schema',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverHttpStream();
      assertLiveToolsList(listed, ADD_TOOL_NAMES);
      assert.ok(
        listed.stderr.includes(
          `MCP Server running on httpStream transport, port ${listed.port}. Endpoint: /mcp`,
        ),
      );
    },
  );
});
