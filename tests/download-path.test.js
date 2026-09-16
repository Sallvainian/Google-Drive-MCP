// tests/download-path.test.js
import {
  downloadAttachment,
  resolveSafeDownloadPath,
} from '../dist/googleGmailApiHelpers.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import {
  existsSync,
  lstatSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, sep } from 'node:path';
import { afterEach, beforeEach, describe, it } from 'node:test';

function stubGmail() {
  return {
    users: {
      messages: {
        attachments: {
          get: async () => ({
            data: { data: Buffer.from('x').toString('base64') },
          }),
        },
      },
    },
  };
}

describe('download path containment', () => {
  let tmpRoot;
  let savePath;

  beforeEach(() => {
    tmpRoot = mkdtempSync(join(tmpdir(), 'download-path-'));
    savePath = join(tmpRoot, 'save');
    mkdirSync(savePath);
  });

  afterEach(() => {
    rmSync(tmpRoot, { recursive: true, force: true });
  });

  it('writes report.pdf inside savePath', async () => {
    const result = await downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf');
    const expected = resolve(savePath, 'report.pdf');
    assert.strictEqual(result.savedTo, expected);
    assert.ok(existsSync(expected));
    assert.strictEqual(readFileSync(expected, 'utf8'), 'x');
  });

  it('treats ../outside.txt as a contained write of outside.txt', async () => {
    const naive = join(savePath, '../outside.txt');
    assert.ok(!resolve(naive).startsWith(resolve(savePath) + sep));

    const result = await downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, '../outside.txt');
    const expected = resolve(savePath, 'outside.txt');
    assert.strictEqual(result.savedTo, expected);
    assert.ok(existsSync(expected));
    assert.strictEqual(readFileSync(expected, 'utf8'), 'x');
    assert.equal(existsSync(join(tmpRoot, 'outside.txt')), false);
    assert.equal(existsSync(naive), false);
  });

  it('rejects filename .. with UserError and writes nothing', async () => {
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, '..'),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Download path escapes the requested directory.');
        return true;
      }
    );
    assert.ok(statSync(savePath).isDirectory());
    assert.ok(statSync(tmpRoot).isDirectory());
    assert.deepStrictEqual(readdirSync(savePath), []);
    assert.deepStrictEqual(readdirSync(tmpRoot), ['save']);
  });

  it('writes only the basename of a path-separator filename', async () => {
    const result = await downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'sub/evil.txt');
    const expected = resolve(savePath, 'evil.txt');
    assert.strictEqual(result.savedTo, expected);
    assert.ok(existsSync(expected));
    assert.equal(existsSync(join(savePath, 'sub')), false);
  });

  it('resolves a Drive-side name under savePath', () => {
    const resolved = resolveSafeDownloadPath(savePath, '../../.ssh/authorized_keys');
    const expected = resolve(savePath, 'authorized_keys');
    assert.strictEqual(resolved, expected);
    assert.ok(resolved.startsWith(resolve(savePath) + sep));
    assert.equal(existsSync(join(tmpRoot, '.ssh')), false);
  });

  it('refuses an existing file when overwrite is omitted or false', async () => {
    const dest = join(savePath, 'report.pdf');
    writeFileSync(dest, 'hello');
    const assertOverwriteError = (error) => {
      assert.ok(error instanceof UserError);
      assert.ok(error.message.startsWith('File already exists:'));
      assert.ok(error.message.includes('Pass overwrite: true to replace it.'));
      return true;
    };
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf'),
      assertOverwriteError
    );
    assert.strictEqual(readFileSync(dest, 'utf8'), 'hello');
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf', false),
      assertOverwriteError
    );
    assert.strictEqual(readFileSync(dest, 'utf8'), 'hello');
  });

  it('replaces an existing file when overwrite is true', async () => {
    const dest = join(savePath, 'report.pdf');
    writeFileSync(dest, 'hello');
    const result = await downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf', true);
    assert.strictEqual(result.savedTo, resolve(savePath, 'report.pdf'));
    assert.strictEqual(readFileSync(dest, 'utf8'), 'x');
  });

  it('refuses a dangling symlink destination even when overwrite is true', async () => {
    const dest = join(savePath, 'report.pdf');
    const outside = join(tmpRoot, 'pwned');
    symlinkSync(outside, dest);
    const assertSymlinkError = (error) => {
      assert.ok(error instanceof UserError);
      assert.ok(error.message.startsWith('Refusing to write through a symlink:'));
      return true;
    };
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf'),
      assertSymlinkError
    );
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf', true),
      assertSymlinkError
    );
    assert.throws(
      () => resolveSafeDownloadPath(savePath, 'report.pdf', true),
      assertSymlinkError
    );
    assert.ok(lstatSync(dest).isSymbolicLink());
    assert.equal(existsSync(outside), false);
  });

  it('refuses a live symlink so overwrite cannot truncate the target', async () => {
    const dest = join(savePath, 'report.pdf');
    const outside = join(tmpRoot, 'pwned');
    writeFileSync(outside, 'secret');
    symlinkSync(outside, dest);
    const assertSymlinkError = (error) => {
      assert.ok(error instanceof UserError);
      assert.ok(error.message.startsWith('Refusing to write through a symlink:'));
      return true;
    };
    await assert.rejects(
      () => downloadAttachment(stubGmail(), 'msg1', 'att1', savePath, 'report.pdf', true),
      assertSymlinkError
    );
    assert.throws(
      () => resolveSafeDownloadPath(savePath, 'report.pdf', true),
      assertSymlinkError
    );
    assert.strictEqual(readFileSync(outside, 'utf8'), 'secret');
    assert.ok(lstatSync(dest).isSymbolicLink());
  });

  it('keeps the default attachment name inside savePath', async () => {
    const attachmentId = '/../../xxxxxxxx';
    assert.strictEqual(attachmentId.substring(0, 8), '/../../x');
    const defaultName = `attachment_${attachmentId.substring(0, 8)}`;
    assert.strictEqual(defaultName, 'attachment_/../../x');
    const naive = join(savePath, defaultName);
    assert.ok(!resolve(naive).startsWith(resolve(savePath) + sep));

    const result = await downloadAttachment(stubGmail(), 'msg1', attachmentId, savePath);
    assert.ok(result.savedTo.startsWith(resolve(savePath) + sep));
    assert.ok(existsSync(result.savedTo));
    assert.equal(existsSync(join(tmpRoot, 'x')), false);
    assert.equal(existsSync(naive), false);
  });
});
