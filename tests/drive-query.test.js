// tests/drive-query.test.js
import { driveQueryQuoted } from '../dist/googleDriveApiHelpers.js';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

const INTERPOLATING_TOOLS = [
  'listGoogleDocs',
  'searchGoogleDocs',
  'getRecentGoogleDocs',
  'listFolderContents',
  'listAllFolders',
  'listGoogleSheets',
  'listGoogleSlides',
];

function toolBody(source, toolName) {
  const start = source.search(new RegExp(`name:\\s*['"]${toolName}['"]`));
  assert.notEqual(start, -1, `missing tool ${toolName}`);
  const rest = source.slice(start);
  const next = rest.indexOf('\nserver.addTool({');
  return next === -1 ? rest : rest.slice(0, next);
}

describe('driveQueryQuoted', () => {
  it('escapes apostrophes inside the quoted literal', () => {
    assert.strictEqual(driveQueryQuoted("O'Brien"), "'O\\'Brien'");
  });

  it('escapes backslashes inside the quoted literal', () => {
    assert.strictEqual(driveQueryQuoted('\\authors'), "'\\\\authors'");
  });

  it('escapes backslash first, then apostrophe', () => {
    assert.strictEqual(
      driveQueryQuoted("quinn's paper\\essay"),
      "'quinn\\'s paper\\\\essay'",
    );
  });

  it('keeps an injection payload as one contains literal', () => {
    const q = `name contains ${driveQueryQuoted("x' or trashed=true or name contains 'y")}`;
    assert.strictEqual(
      q,
      "name contains 'x\\' or trashed=true or name contains \\'y'",
    );
    assert.notStrictEqual(
      q,
      "name contains 'x' or trashed=true or name contains 'y'",
    );
  });

  it('quotes a plain term without extra escaping', () => {
    assert.strictEqual(driveQueryQuoted('budget'), "'budget'");
  });

  it('quotes a folder id', () => {
    assert.strictEqual(driveQueryQuoted('root'), "'root'");
  });
});

describe('Drive q interpolations in server.ts', () => {
  it('calls driveQueryQuoted in each interpolating tool and has no raw quoted interpolations', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.equal(source.includes("'${"), false);
    for (const toolName of INTERPOLATING_TOOLS) {
      const body = toolBody(source, toolName);
      assert.ok(
        body.includes('DriveHelpers.driveQueryQuoted'),
        `${toolName} must call DriveHelpers.driveQueryQuoted`,
      );
    }
  });
});
