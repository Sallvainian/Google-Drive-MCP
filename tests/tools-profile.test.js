// tests/tools-profile.test.js
import assert from 'node:assert';
import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const metaGroup = /_meta[\s\S]{0,80}\bgroup\s*:/;

describe('src/ source text has no profile filter', () => {
  it('does not put MCP_TOOL_PROFILE or toolsets in src/', () => {
    for (const file of readdirSync(srcDir)) {
      const source = readFileSync(join(srcDir, file), 'utf8');
      assert.equal(source.includes('MCP_TOOL_PROFILE'), false, `${file} contains MCP_TOOL_PROFILE`);
      assert.equal(source.includes('toolsets'), false, `${file} contains toolsets`);
    }
  });

  it('does not put grouping metadata or FastMCP hide APIs in src/', () => {
    for (const file of readdirSync(srcDir)) {
      const source = readFileSync(join(srcDir, file), 'utf8');
      assert.equal(source.includes('_meta.group'), false, `${file} contains _meta.group`);
      assert.equal(metaGroup.test(source), false, `${file} contains a _meta object with a group key`);
      assert.equal(source.includes('annotations'), false, `${file} contains annotations`);
      assert.equal(source.includes('toolset'), false, `${file} contains toolset`);
      assert.equal(source.includes('canAccess'), false, `${file} contains canAccess`);
      assert.equal(source.includes('removeTool'), false, `${file} contains removeTool`);
      assert.equal(source.includes('removeTools'), false, `${file} contains removeTools`);
    }
  });
});
