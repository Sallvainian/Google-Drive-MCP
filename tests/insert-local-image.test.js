// tests/insert-local-image.test.js
import {
  insertLocalImageFromPath,
} from '../dist/googleDocsApiHelpers.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import {
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { afterEach, beforeEach, describe, it, mock } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const WEB_CONTENT_LINK = 'https://drive.google.com/uc?id=img1';

function toolBody(source, toolName) {
  const start = source.search(new RegExp(`name:\\s*['"]${toolName}['"]`));
  assert.notEqual(start, -1, `missing tool ${toolName}`);
  const rest = source.slice(start);
  const next = rest.indexOf('\nserver.addTool({');
  return next === -1 ? rest : rest.slice(0, next);
}

function consumeUploadStream(params) {
  const body = params?.media?.body;
  if (!body || typeof body.on !== 'function') {
    return Promise.resolve();
  }
  return new Promise((resolve, reject) => {
    body.on('error', reject);
    body.on('end', resolve);
    body.resume();
  });
}

function makeStubs(overrides = {}) {
  const order = [];
  const drive = {
    files: {
      create: mock.fn(async (params) => {
        order.push('files.create');
        await consumeUploadStream(params);
        if (overrides.create) {
          return overrides.create();
        }
        return { data: { id: 'img1' } };
      }),
      get: mock.fn(async () => {
        order.push('files.get');
        if (overrides.get) {
          return overrides.get();
        }
        return { data: { webContentLink: WEB_CONTENT_LINK } };
      }),
    },
    permissions: {
      create: mock.fn(async () => {
        order.push('permissions.create');
        if (overrides.permissionCreate) {
          return overrides.permissionCreate();
        }
        return { data: { id: 'anyoneWithLink' } };
      }),
      delete: mock.fn(async () => {
        order.push('permissions.delete');
        if (overrides.delete) {
          return overrides.delete();
        }
        return {};
      }),
    },
  };
  const docs = {
    documents: {
      batchUpdate: mock.fn(async () => {
        order.push('documents.batchUpdate');
        if (overrides.batchUpdate) {
          return overrides.batchUpdate();
        }
        return { data: {} };
      }),
    },
  };
  return { drive, docs, order };
}

describe('insertLocalImageFromPath', () => {
  let tmpDir;
  let pngPath;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'insert-local-image-'));
    pngPath = join(tmpDir, 'image.png');
    writeFileSync(pngPath, Buffer.from([0x89, 0x50, 0x4e, 0x47]));
  });

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('grants anyone/reader, inserts, then revokes the grant', async () => {
    const { drive, docs, order } = makeStubs();

    const result = await insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1);

    assert.deepStrictEqual(result, {
      fileId: 'img1',
      webContentLink: WEB_CONTENT_LINK,
    });
    assert.deepStrictEqual(order, [
      'files.create',
      'permissions.create',
      'files.get',
      'documents.batchUpdate',
      'permissions.delete',
    ]);
    assert.deepStrictEqual(drive.permissions.create.mock.calls[0].arguments[0], {
      fileId: 'img1',
      requestBody: { role: 'reader', type: 'anyone' },
      fields: 'id',
    });
    assert.deepStrictEqual(drive.permissions.delete.mock.calls[0].arguments[0], {
      fileId: 'img1',
      permissionId: 'anyoneWithLink',
    });
    assert.strictEqual(drive.permissions.create.mock.calls.length, 1);
    assert.ok(order.indexOf('permissions.create') < order.indexOf('permissions.delete'));
  });

  it('revokes the grant when insert fails after the grant', async () => {
    const { drive, docs } = makeStubs({
      batchUpdate: () => {
        throw { code: 400, message: 'bad index' };
      },
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Google API Error (400): bad index');
        return true;
      }
    );
    assert.strictEqual(drive.files.create.mock.calls.length, 1);
    assert.deepStrictEqual(drive.permissions.delete.mock.calls[0].arguments[0], {
      fileId: 'img1',
      permissionId: 'anyoneWithLink',
    });
  });

  it('does not return success when revoke fails after insert', async () => {
    const { drive, docs } = makeStubs({
      delete: () => {
        throw { code: 500, message: 'backend' };
      },
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Failed to revoke the temporary anyone/reader grant on file img1; the anyone grant may still exist: backend'
        );
        return true;
      }
    );
  });

  it('throws UserError for a missing local file and does not upload', async () => {
    const missingPath = join(tmpDir, 'missing.png');
    const { drive, docs } = makeStubs();

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, missingPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, `Image file not found: ${missingPath}`);
        return true;
      }
    );
    assert.strictEqual(drive.files.create.mock.calls.length, 0);
    assert.strictEqual(drive.permissions.create.mock.calls.length, 0);
  });

  it('uploads to parentFolderId and inserts with objectSize when width and height are set', async () => {
    const { drive, docs } = makeStubs();

    const result = await insertLocalImageFromPath(
      docs,
      drive,
      pngPath,
      'doc1',
      1,
      120,
      80,
      'folder1'
    );

    assert.strictEqual(result.fileId, 'img1');
    assert.deepStrictEqual(
      drive.files.create.mock.calls[0].arguments[0].requestBody.parents,
      ['folder1']
    );
    const insert = docs.documents.batchUpdate.mock.calls[0].arguments[0]
      .requestBody.requests[0].insertInlineImage;
    assert.strictEqual(insert.uri, WEB_CONTENT_LINK);
    assert.deepStrictEqual(insert.objectSize, {
      height: { magnitude: 80, unit: 'PT' },
      width: { magnitude: 120, unit: 'PT' },
    });
  });

  it('revokes the grant when files.get fails after the grant', async () => {
    const getError = { code: 500, message: 'backend' };
    const { drive, docs } = makeStubs({
      get: () => {
        throw getError;
      },
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.strictEqual(error, getError);
        assert.equal(error instanceof UserError, false);
        return true;
      }
    );
    assert.deepStrictEqual(drive.permissions.delete.mock.calls[0].arguments[0], {
      fileId: 'img1',
      permissionId: 'anyoneWithLink',
    });
  });

  it('revokes the grant when webContentLink is missing after the grant', async () => {
    const { drive, docs } = makeStubs({
      get: () => ({ data: {} }),
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Failed to get public URL for uploaded image');
        return true;
      }
    );
    assert.deepStrictEqual(drive.permissions.delete.mock.calls[0].arguments[0], {
      fileId: 'img1',
      permissionId: 'anyoneWithLink',
    });
  });

  it('throws UserError when permissions.create returns no id and does not revoke', async () => {
    const { drive, docs } = makeStubs({
      permissionCreate: () => ({ data: {} }),
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Failed to upload image to Drive - no permission ID returned'
        );
        return true;
      }
    );
    assert.strictEqual(drive.permissions.delete.mock.calls.length, 0);
  });

  it('throws UserError when create returns no file id and does not grant', async () => {
    const { drive, docs } = makeStubs({
      create: () => ({ data: {} }),
    });

    await assert.rejects(
      () => insertLocalImageFromPath(docs, drive, pngPath, 'doc1', 1),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Failed to upload image to Drive - no file ID returned'
        );
        return true;
      }
    );
    assert.strictEqual(drive.permissions.create.mock.calls.length, 0);
  });
});

describe('insertLocalImage tool source', () => {
  it('discloses the temporary anyone/reader grant, calls the sequencer, and does not print Image URL', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    const body = toolBody(source, 'insertLocalImage');
    const descMatch = body.match(/description:\s*'([^']*)'/);
    assert.ok(descMatch, 'insertLocalImage description string');
    const description = descMatch[1];
    assert.match(description, /anyone\/reader|anyone-with-link/);
    assert.match(description, /revok/i);
    assert.match(description, /temporar/i);
    assert.ok(body.includes('GDocsHelpers.insertLocalImageFromPath'));
    assert.ok(body.includes(`GDocsHelpers.insertLocalImageFromPath(
docs,
drive,
args.localImagePath,
args.documentId,
args.index,
args.width,
args.height,
parentFolderId
)`));
    assert.ok(body.includes('uploaded.fileId'));
    assert.match(body, /grant was revoked/);
    assert.equal(body.includes('Image URL:'), false);
  });
});
