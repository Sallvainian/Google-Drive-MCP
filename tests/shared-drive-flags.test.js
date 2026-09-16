// tests/shared-drive-flags.test.js
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const serverSrc = readFileSync(join(srcDir, 'server.ts'), 'utf8');
const helpersSrc = readFileSync(join(srcDir, 'googleDocsApiHelpers.ts'), 'utf8');
const driveHelpersSrc = readFileSync(join(srcDir, 'googleDriveApiHelpers.ts'), 'utf8');
const combinedSrc = `${helpersSrc}\n${driveHelpersSrc}\n${serverSrc}`;

function skipWsAndComments(source, i) {
  while (i < source.length) {
    if (/\s/.test(source[i])) {
      i++;
      continue;
    }
    if (source.startsWith('//', i)) {
      const nl = source.indexOf('\n', i);
      i = nl === -1 ? source.length : nl + 1;
      continue;
    }
    if (source.startsWith('/*', i)) {
      const end = source.indexOf('*/', i + 2);
      i = end === -1 ? source.length : end + 2;
      continue;
    }
    break;
  }
  return i;
}

function extractBalanced(source, start, open, close) {
  let depth = 0;
  let inString = null;
  let escaped = false;
  for (let i = start; i < source.length; i++) {
    const ch = source[i];
    if (inString) {
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === '`') {
      inString = ch;
      continue;
    }
    if (ch === open) depth++;
    else if (ch === close) {
      depth--;
      if (depth === 0) return { text: source.slice(start, i + 1), end: i + 1 };
    }
  }
  throw new Error(`unbalanced ${open}${close} at ${start}`);
}

function objectHasKey(objText, key) {
  const re = new RegExp(`(?:^|[,{\\s])(?:${key}|['"]${key}['"])\\s*:`);
  return re.test(objText);
}

function objectKeys(objText) {
  const inner = objText.slice(1, -1);
  const keys = [];
  let depth = 0;
  let inString = null;
  let escaped = false;
  let token = '';
  let expectingKey = true;
  for (let i = 0; i < inner.length; i++) {
    const ch = inner[i];
    if (inString) {
      token += ch;
      if (escaped) {
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        escaped = true;
        continue;
      }
      if (ch === inString) inString = null;
      continue;
    }
    if (ch === '"' || ch === "'" || ch === '`') {
      if (expectingKey) token += ch;
      inString = ch;
      continue;
    }
    if (ch === '{' || ch === '(' || ch === '[') {
      depth++;
      continue;
    }
    if (ch === '}' || ch === ')' || ch === ']') {
      depth--;
      continue;
    }
    if (depth !== 0) continue;
    if (expectingKey) {
      if (ch === ':') {
        const key = token.trim().replace(/^['"]|['"]$/g, '');
        if (key) keys.push(key);
        token = '';
        expectingKey = false;
      } else if (ch === ',' || /\s/.test(ch)) {
        if (ch === ',') token = '';
        else token += ch;
      } else {
        token += ch;
      }
    } else if (ch === ',') {
      expectingKey = true;
      token = '';
    }
  }
  return keys;
}

function resolveIdentObject(source, ident, callIndex) {
  const before = source.slice(0, callIndex);
  const re = new RegExp(`(?:let|const|var)\\s+${ident}\\b(?:\\s*:\\s*[^=]+)?\\s*=\\s*\\{`, 'g');
  let last = -1;
  let match;
  while ((match = re.exec(before))) last = match.index;
  assert.notEqual(last, -1, `no assignment for ${ident}`);
  const brace = before.indexOf('{', last);
  return extractBalanced(source, brace, '{', '}').text;
}

function collectDriveCalls(source) {
  const calls = [];
  const re = /drive\.(files|permissions)\.([A-Za-z]+)\s*\(/g;
  let match;
  while ((match = re.exec(source))) {
    const resource = match[1];
    const method = match[2];
    const openParen = match.index + match[0].length - 1;
    let i = skipWsAndComments(source, openParen + 1);
    let paramsText;
    if (source[i] === '{') {
      paramsText = extractBalanced(source, i, '{', '}').text;
    } else {
      const identMatch = source.slice(i).match(/^([A-Za-z_][A-Za-z0-9_]*)/);
      assert.ok(identMatch, `expected object or ident after drive.${resource}.${method}(`);
      paramsText = resolveIdentObject(source, identMatch[1], match.index);
    }
    calls.push({ resource, method, paramsText, index: match.index });
  }
  return calls;
}

const calls = collectDriveCalls(combinedSrc);
const fileCalls = calls.filter((c) => c.resource === 'files');
const permissionCalls = calls.filter((c) => c.resource === 'permissions');
const listCalls = fileCalls.filter((c) => c.method === 'list');
const exportCalls = fileCalls.filter((c) => c.method === 'export');
const flaggedFileCalls = fileCalls.filter((c) => c.method !== 'export');

describe('shared-drive flags', () => {
  it('keeps 31 drive.files.* calls and 5 drive.permissions.* calls', () => {
    assert.strictEqual(fileCalls.length, 31);
    assert.strictEqual(permissionCalls.length, 5);
    const permissionMethods = permissionCalls.map((c) => c.method).sort();
    assert.deepStrictEqual(permissionMethods, ['create', 'create', 'create', 'delete', 'list'].sort());
  });

  it('puts supportsAllDrives and includeItemsFromAllDrives on every files.list params object', () => {
    assert.strictEqual(listCalls.length, 7);
    for (const call of listCalls) {
      assert.equal(
        objectHasKey(call.paramsText, 'supportsAllDrives'),
        true,
        `files.list missing supportsAllDrives: ${call.paramsText}`
      );
      assert.match(call.paramsText, /supportsAllDrives:\s*true/);
      assert.equal(
        objectHasKey(call.paramsText, 'includeItemsFromAllDrives'),
        true,
        `files.list missing includeItemsFromAllDrives: ${call.paramsText}`
      );
      assert.match(call.paramsText, /includeItemsFromAllDrives:\s*true/);
      assert.equal(objectHasKey(call.paramsText, 'corpora'), false);
      assert.equal(objectHasKey(call.paramsText, 'driveId'), false);
    }
  });

  it('puts supportsAllDrives on every files.* params object except files.export', () => {
    assert.strictEqual(flaggedFileCalls.length, 30);
    for (const call of flaggedFileCalls) {
      assert.equal(
        objectHasKey(call.paramsText, 'supportsAllDrives'),
        true,
        `files.${call.method} missing supportsAllDrives: ${call.paramsText}`
      );
      assert.match(call.paramsText, /supportsAllDrives:\s*true/);
      assert.equal(objectHasKey(call.paramsText, 'corpora'), false);
      assert.equal(objectHasKey(call.paramsText, 'driveId'), false);
    }
  });

  it('leaves files.export params as fileId and mimeType only', () => {
    assert.strictEqual(exportCalls.length, 1);
    const params = exportCalls[0].paramsText;
    assert.equal(objectHasKey(params, 'supportsAllDrives'), false);
    assert.equal(objectHasKey(params, 'includeItemsFromAllDrives'), false);
    const keys = objectKeys(params);
    assert.deepStrictEqual(keys.sort(), ['fileId', 'mimeType'].sort());
    const exportCall = serverSrc.slice(
      serverSrc.indexOf('drive.files.export('),
      serverSrc.indexOf('drive.files.export(') + 220
    );
    assert.match(exportCall, /responseType:\s*'arraybuffer'/);
  });

  it('puts supportsAllDrives on every permissions.* params object', () => {
    assert.strictEqual(permissionCalls.length, 5);
    for (const call of permissionCalls) {
      assert.equal(
        objectHasKey(call.paramsText, 'supportsAllDrives'),
        true,
        `permissions.${call.method} missing supportsAllDrives: ${call.paramsText}`
      );
      assert.match(call.paramsText, /supportsAllDrives:\s*true/);
      assert.equal(objectHasKey(call.paramsText, 'corpora'), false);
      assert.equal(objectHasKey(call.paramsText, 'driveId'), false);
    }
  });

  it('keeps the exportError.code === 403 exportViaWebUrl branch', () => {
    assert.match(
      serverSrc,
      /if \(exportError\.code === 403\)[\s\S]{0,400}?exportViaWebUrl\(args\.fileId, fileMimeType, args\.exportFormat\)/
    );
    const definitions = serverSrc.match(/async function exportViaWebUrl/g) || [];
    const callsites = serverSrc.match(/exportViaWebUrl\(/g) || [];
    assert.strictEqual(definitions.length, 1);
    assert.ok(callsites.length >= 2);
  });

  it('does not add corpora or driveId tool parameters', () => {
    assert.doesNotMatch(serverSrc, /\bcorpora\b/);
    assert.doesNotMatch(serverSrc, /\bdriveId\b/);
    assert.doesNotMatch(helpersSrc, /\bcorpora\b/);
    assert.doesNotMatch(helpersSrc, /\bdriveId\b/);
    assert.doesNotMatch(driveHelpersSrc, /\bcorpora\b/);
    assert.doesNotMatch(driveHelpersSrc, /\bdriveId\b/);
  });
});
