// tests/throw-classification.test.js
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function throwNewErrorCount(source) {
  return (source.match(/throw new Error/g) || []).length;
}

describe('internal-fault throw sites remain Error', () => {
  it('should keep eight throw new Error sites in auth.ts and no UserError import', () => {
    const source = readFileSync(join(srcDir, 'auth.ts'), 'utf8');
    assert.strictEqual(throwNewErrorCount(source), 8);
    assert.equal(source.includes("from 'fastmcp'"), false);
    assert.equal(source.includes('UserError'), false);
  });

  it('should keep the two initializeGoogleClient throw new Error strings in server.ts', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.ok(source.includes('throw new Error("Google client initialization failed. Cannot start server tools.")'));
    assert.ok(source.includes('throw new Error("Google Docs, Drive, Sheets, Slides, and Gmail clients could not be initialized.")'));
    assert.strictEqual(throwNewErrorCount(source), 2);
  });
});
