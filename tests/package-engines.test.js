// tests/package-engines.test.js
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const pkg = JSON.parse(
  readFileSync(join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json'), 'utf8')
);

describe('package.json engines and pretest', () => {
  it('declares engines.node >=22 and pretest tsc', () => {
    assert.strictEqual(pkg.engines.node, '>=22');
    assert.strictEqual(pkg.scripts.pretest, 'tsc');
  });
});
