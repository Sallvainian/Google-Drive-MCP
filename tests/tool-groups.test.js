// tests/tool-groups.test.js
import {
  DEFAULT_TOOL_GROUPS,
  parseEnabledToolGroups,
  TOOL_GROUPS,
} from '../dist/toolGroups.js';
import {
  createStubAuthorizeEnv,
  killLiveServer,
  listToolsOverStdio,
  spawnLiveServer,
} from './live-fastmcp-helpers.js';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

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

function groupForIndex(index) {
  if (index < 41) return 'docs';
  if (index < 44) return 'docs-chips';
  if (index < 65) return 'drive';
  if (index < 74) return 'sheets';
  if (index < 104) return 'sheets-advanced';
  if (index < 108) return 'docs';
  if (index < 126) return 'slides';
  return 'gmail';
}

function expectedNamesForGroups(groups) {
  const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
  const names = addToolNames(source);
  assert.equal(names.length, 160);
  const enabled = new Set(groups);
  return names.filter((_name, index) => enabled.has(groupForIndex(index)));
}

function withToolGroupsEnv(value, fn) {
  const previous = process.env.MCP_TOOL_GROUPS;
  if (value === undefined) {
    delete process.env.MCP_TOOL_GROUPS;
  } else {
    process.env.MCP_TOOL_GROUPS = value;
  }
  try {
    return fn();
  } finally {
    if (previous === undefined) {
      delete process.env.MCP_TOOL_GROUPS;
    } else {
      process.env.MCP_TOOL_GROUPS = previous;
    }
  }
}

describe('parseEnabledToolGroups', () => {
  it('unset, whitespace, and empty tokens return the five DEFAULT_TOOL_GROUPS families', () => {
    withToolGroupsEnv(undefined, () => {
      assert.deepStrictEqual(parseEnabledToolGroups(), [...DEFAULT_TOOL_GROUPS]);
      assert.deepStrictEqual(parseEnabledToolGroups(undefined), [...DEFAULT_TOOL_GROUPS]);
    });
    assert.deepStrictEqual(parseEnabledToolGroups(''), [...DEFAULT_TOOL_GROUPS]);
    assert.deepStrictEqual(parseEnabledToolGroups('  '), [...DEFAULT_TOOL_GROUPS]);
    assert.deepStrictEqual(parseEnabledToolGroups(',,,'), [...DEFAULT_TOOL_GROUPS]);
    assert.deepStrictEqual(DEFAULT_TOOL_GROUPS, [
      'docs',
      'drive',
      'sheets',
      'slides',
      'gmail',
    ]);
    assert.equal(DEFAULT_TOOL_GROUPS.includes('sheets-advanced'), false);
    assert.equal(DEFAULT_TOOL_GROUPS.includes('docs-chips'), false);
  });

  it('docs,drive,gmail keeps canonical order', () => {
    assert.deepStrictEqual(parseEnabledToolGroups('docs,drive,gmail'), [
      'docs',
      'drive',
      'gmail',
    ]);
  });

  it('docs only is docs', () => {
    assert.deepStrictEqual(parseEnabledToolGroups('docs'), ['docs']);
  });

  it('all, any case, alone or in the list, enables every family including opt-in', () => {
    assert.deepStrictEqual(parseEnabledToolGroups('all'), [...TOOL_GROUPS]);
    assert.deepStrictEqual(parseEnabledToolGroups('ALL'), [...TOOL_GROUPS]);
    assert.deepStrictEqual(parseEnabledToolGroups('docs,all'), [...TOOL_GROUPS]);
    assert.deepStrictEqual(parseEnabledToolGroups('unknown,all'), [...TOOL_GROUPS]);
    assert.deepStrictEqual(TOOL_GROUPS, [
      'docs',
      'drive',
      'sheets',
      'slides',
      'gmail',
      'sheets-advanced',
      'docs-chips',
    ]);
  });

  it('opt-in empty families are valid and stay in the list', () => {
    assert.deepStrictEqual(parseEnabledToolGroups('docs,sheets-advanced'), [
      'docs',
      'sheets-advanced',
    ]);
    assert.deepStrictEqual(parseEnabledToolGroups('docs-chips'), ['docs-chips']);
  });

  it('trims, lower-cases, drops empties, collapses duplicates, canonical order', () => {
    assert.deepStrictEqual(parseEnabledToolGroups('Sheets, docs, sheets'), [
      'docs',
      'sheets',
    ]);
    assert.deepStrictEqual(parseEnabledToolGroups(' gmail , , Drive,docs '), [
      'docs',
      'drive',
      'gmail',
    ]);
  });

  it('unknown names without all throw Unknown MCP_TOOL_GROUPS value(s)', () => {
    assert.throws(
      () => parseEnabledToolGroups('docs,unknown'),
      (error) => {
        assert.equal(error instanceof Error, true);
        assert.equal(
          error.message,
          'Unknown MCP_TOOL_GROUPS value(s): unknown. Valid groups: docs, drive, sheets, slides, gmail, sheets-advanced, docs-chips',
        );
        return true;
      },
    );
    assert.throws(
      () => parseEnabledToolGroups('utils,calendar'),
      /Unknown MCP_TOOL_GROUPS value\(s\): utils, calendar\. Valid groups: /,
    );
  });
});

describe('group membership vs ADD_TOOL_NAMES', () => {
  it('default-on families map to 127 names and all maps to 160', () => {
    const unsetNames = expectedNamesForGroups([...DEFAULT_TOOL_GROUPS]);
    const allNames = expectedNamesForGroups([...TOOL_GROUPS]);
    assert.equal(unsetNames.length, 127);
    assert.equal(allNames.length, 160);
    assert.equal(unsetNames.includes('batchWrite'), false);
    assert.equal(unsetNames.includes('insertDateChip'), false);
    assert.equal(allNames.includes('batchWrite'), true);
    assert.equal(allNames.includes('insertDateChip'), true);
  });

  it('docs is 45, sheets-advanced adds 30, docs-chips is the 3 chip names', () => {
    const docsNames = expectedNamesForGroups(['docs']);
    const withAdvanced = expectedNamesForGroups(['docs', 'sheets-advanced']);
    const chipsOnly = expectedNamesForGroups(['docs-chips']);
    const advancedOnly = expectedNamesForGroups(['sheets-advanced']);
    assert.equal(docsNames.length, 45);
    assert.equal(withAdvanced.length, 75);
    assert.equal(advancedOnly.length, 30);
    assert.ok(advancedOnly.includes('batchWrite'));
    assert.ok(advancedOnly.includes('createTable'));
    assert.equal(advancedOnly.includes('readSpreadsheet'), false);
    assert.deepStrictEqual(chipsOnly, ['insertDateChip', 'insertPerson', 'insertRichLink']);
  });

  it('src/server.ts assigns currentToolGroup to sheets-advanced and docs-chips', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.equal(source.includes("currentToolGroup = 'sheets-advanced'"), true);
    assert.equal(source.includes("currentToolGroup = 'docs-chips'"), true);
  });
});

describe('in-process dist/server.js imports', () => {
  it('docs-tab-gets-and-create.test.js deletes MCP_TOOL_GROUPS before importing dist/server.js', () => {
    const source = readFileSync(
      join(dirname(fileURLToPath(import.meta.url)), 'docs-tab-gets-and-create.test.js'),
      'utf8',
    );
    const deleteAt = source.indexOf('delete process.env.MCP_TOOL_GROUPS');
    const importAt = source.indexOf("await import('../dist/server.js')");
    assert.notEqual(deleteAt, -1);
    assert.notEqual(importAt, -1);
    assert.ok(deleteAt < importAt);
  });
});

describe('createStubAuthorizeEnv MCP_TOOL_GROUPS', () => {
  it('deletes inherited MCP_TOOL_GROUPS unless the caller override sets it', () => {
    const previous = process.env.MCP_TOOL_GROUPS;
    process.env.MCP_TOOL_GROUPS = 'docs';
    try {
      const stripped = createStubAuthorizeEnv();
      assert.equal(Object.hasOwn(stripped, 'MCP_TOOL_GROUPS'), false);
      const kept = createStubAuthorizeEnv({ MCP_TOOL_GROUPS: 'gmail' });
      assert.equal(kept.MCP_TOOL_GROUPS, 'gmail');
    } finally {
      if (previous === undefined) {
        delete process.env.MCP_TOOL_GROUPS;
      } else {
        process.env.MCP_TOOL_GROUPS = previous;
      }
    }
  });
});

describe('live tools/list MCP_TOOL_GROUPS', () => {
  it(
    'docs,drive,gmail lists 100 tools including createFormattedDocument and no Sheets/Slides',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverStdio({
        env: { MCP_TOOL_GROUPS: 'docs,drive,gmail' },
      });
      const names = listed.tools.map((tool) => tool.name);
      const expected = expectedNamesForGroups(['docs', 'drive', 'gmail']);
      assert.equal(expected.length, 100);
      assert.deepStrictEqual(names, expected);
      assert.equal(names.includes('createFormattedDocument'), true);
      assert.equal(names.includes('addTab'), true);
      assert.equal(names.includes('renameTab'), true);
      assert.equal(names.includes('replaceDocumentWithMarkdown'), true);
      assert.equal(names.includes('appendMarkdownToGoogleDoc'), true);
      assert.equal(names.includes('replaceRangeWithMarkdown'), true);
      assert.equal(names.includes('readSpreadsheet'), false);
      assert.equal(names.includes('listGoogleSlides'), false);
      assert.match(listed.stderr, /Registered tool groups: docs, drive, gmail(?:\n|$)/);
    },
  );

  it(
    'docs,sheets-advanced lists 75 tools including the 30 advanced names',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverStdio({
        env: { MCP_TOOL_GROUPS: 'docs,sheets-advanced' },
      });
      const names = listed.tools.map((tool) => tool.name);
      const expected = expectedNamesForGroups(['docs', 'sheets-advanced']);
      assert.equal(expected.length, 75);
      assert.deepStrictEqual(names, expected);
      assert.equal(names.includes('batchWrite'), true);
      assert.equal(names.includes('createTable'), true);
      assert.equal(names.includes('insertDateChip'), false);
      assert.match(
        listed.stderr,
        /Registered tool groups: docs, sheets-advanced(?:\n|$)/,
      );
    },
  );

  it(
    'docs-chips lists exactly insertDateChip, insertPerson, insertRichLink',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverStdio({
        env: { MCP_TOOL_GROUPS: 'docs-chips' },
      });
      const names = listed.tools.map((tool) => tool.name);
      assert.deepStrictEqual(names, ['insertDateChip', 'insertPerson', 'insertRichLink']);
      assert.match(listed.stderr, /Registered tool groups: docs-chips(?:\n|$)/);
    },
  );

  it(
    'all lists 160 tools',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverStdio({
        env: { MCP_TOOL_GROUPS: 'all' },
      });
      const names = listed.tools.map((tool) => tool.name);
      const expected = expectedNamesForGroups([...TOOL_GROUPS]);
      assert.equal(expected.length, 160);
      assert.deepStrictEqual(names, expected);
    },
  );

  it(
    'docs lists 45 tools including the four formatting tools and no Drive/Sheets/Slides/Gmail',
    { timeout: 180000 },
    async () => {
      const listed = await listToolsOverStdio({
        env: { MCP_TOOL_GROUPS: 'docs' },
      });
      const names = listed.tools.map((tool) => tool.name);
      const expected = expectedNamesForGroups(['docs']);
      assert.equal(expected.length, 45);
      assert.deepStrictEqual(names, expected);
      assert.equal(names[0], 'readGoogleDoc');
      assert.equal(names.includes('addTab'), true);
      assert.equal(names.includes('renameTab'), true);
      assert.equal(names.includes('formatMatchingText'), true);
      assert.equal(names.includes('createFormattedDocument'), true);
      assert.equal(names.includes('updateDocumentSection'), true);
      assert.equal(names.includes('listGoogleDocs'), false);
      assert.equal(names.includes('readSpreadsheet'), false);
      assert.equal(names.includes('listGoogleSlides'), false);
      assert.equal(names.includes('send_email'), false);
      assert.match(listed.stderr, /Registered tool groups: docs(?:\n|$)/);
    },
  );

  it('docs,unknown exits 1 with FATAL and does not listen', { timeout: 30000 }, async () => {
    const env = createStubAuthorizeEnv({
      MCP_TRANSPORT: 'stdio',
      MCP_TOOL_GROUPS: 'docs,unknown',
    });
    const child = spawnLiveServer(env);
    const stderrChunks = [];
    child.stderr.on('data', (chunk) => stderrChunks.push(chunk));
    child.stdout.on('data', () => {});
    let result;
    try {
      result = await new Promise((resolve, reject) => {
        const timer = setTimeout(() => {
          reject(new Error('invalid MCP_TOOL_GROUPS child did not exit within 30000ms'));
        }, 30000);
        child.on('error', (err) => {
          clearTimeout(timer);
          reject(err);
        });
        child.on('exit', (code, signal) => {
          clearTimeout(timer);
          resolve({
            code,
            signal,
            stderr: Buffer.concat(stderrChunks).toString('utf8'),
          });
        });
      });
    } finally {
      await killLiveServer(child);
    }
    assert.strictEqual(result.code, 1);
    assert.match(
      result.stderr,
      /FATAL: Invalid MCP_TOOL_GROUPS: Unknown MCP_TOOL_GROUPS value\(s\): unknown/,
    );
    assert.equal(result.stderr.includes('Starting Ultimate'), false);
    assert.equal(result.stderr.includes('MCP Server running'), false);
    assert.equal(result.stderr.includes('Registered tool groups:'), false);
  });
});
