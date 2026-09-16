// tests/docs-destructive-writes.test.js
import {
  buildFormattedContentRequests,
  createFormattedDocument,
  executeBatchUpdate,
  replaceFormattedDocumentContent,
  updateFormattedDocumentSection,
} from '../dist/googleDocsApiHelpers.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it, mock } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function toolBody(source, toolName) {
  const start = source.search(new RegExp(`name:\\s*['"]${toolName}['"]`));
  assert.notEqual(start, -1, `missing tool ${toolName}`);
  const rest = source.slice(start);
  const next = rest.indexOf('\nserver.addTool({');
  return next === -1 ? rest : rest.slice(0, next);
}

function makeReplaceDocs({ endIndex = 20, batchUpdate, includeRevisionId = true, revisionId = 'rev1' } = {}) {
  const data = {
    body: { content: [{ endIndex }] },
  };
  if (includeRevisionId) {
    data.revisionId = revisionId;
  }
  return {
    documents: {
      get: mock.fn(async () => ({ data })),
      batchUpdate: mock.fn(batchUpdate ?? (async () => ({ data: {} }))),
    },
  };
}

function headingDocData({ includeRevisionId = true, revisionId = 'rev1' } = {}) {
  const data = {
    body: {
      content: [
        {
          startIndex: 1,
          endIndex: 15,
          paragraph: {
            elements: [{ textRun: { content: 'Introduction\n' } }],
            paragraphStyle: { namedStyleType: 'HEADING_1' },
          },
        },
        {
          startIndex: 15,
          endIndex: 40,
          paragraph: {
            elements: [{ textRun: { content: 'Old body\n' } }],
            paragraphStyle: { namedStyleType: 'NORMAL_TEXT' },
          },
        },
      ],
    },
  };
  if (includeRevisionId) {
    data.revisionId = revisionId;
  }
  return data;
}

function makeSectionDocs({ revisionId = 'rev1', batchUpdate, includeRevisionId = true } = {}) {
  return {
    documents: {
      get: mock.fn(async () => ({ data: headingDocData({ revisionId, includeRevisionId }) })),
      batchUpdate: mock.fn(batchUpdate ?? (async () => ({ data: {} }))),
    },
  };
}

function firstStyleIndex(requests) {
  return requests.findIndex(
    (request) =>
      request.updateParagraphStyle ||
      request.createParagraphBullets ||
      request.updateTextStyle
  );
}

function normalSections(count) {
  return Array.from({ length: count }, (_, i) => ({
    type: 'normal',
    text: `Section ${i}`,
  }));
}

describe('replaceFormattedDocumentContent', () => {
  it('sends delete then insert in one batchUpdate pinned to the get revisionId', async () => {
    const docs = makeReplaceDocs();
    const sections = [{ type: 'normal', text: 'Hello' }];

    await replaceFormattedDocumentContent(docs, 'doc123', sections);

    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.deepStrictEqual(docs.documents.get.mock.calls[0].arguments[0], {
      documentId: 'doc123',
      fields: 'revisionId,body(content(endIndex))',
    });
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
    const callArgs = docs.documents.batchUpdate.mock.calls[0].arguments[0];
    assert.strictEqual(callArgs.documentId, 'doc123');
    const requests = callArgs.requestBody.requests;
    assert.ok(requests[0].deleteContentRange);
    assert.deepStrictEqual(requests[0].deleteContentRange.range, {
      startIndex: 1,
      endIndex: 19,
    });
    const insertIdx = requests.findIndex((request) => request.insertText);
    assert.ok(insertIdx > 0);
    assert.ok(requests[insertIdx].insertText.text.startsWith('Hello'));
    assert.strictEqual(requests[insertIdx].insertText.location.index, 1);
    const lastInsertIdx = requests.reduce((last, request, i) => (request.insertText ? i : last), -1);
    const styleIdx = firstStyleIndex(requests);
    assert.ok(styleIdx > lastInsertIdx);
    assert.strictEqual(callArgs.requestBody.writeControl.requiredRevisionId, 'rev1');
  });

  it('throws UserError for empty content without get or batchUpdate', async () => {
    const docs = makeReplaceDocs();

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', []),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'content must not be empty.');
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 0);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('throws UserError for empty-text items without get or batchUpdate', async () => {
    const docs = makeReplaceDocs();

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', [{ type: 'normal', text: '' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'content must not be empty.');
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 0);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('skips deleteContentRange when endIndex is 2', async () => {
    const docs = makeReplaceDocs({ endIndex: 2 });
    const sections = [{ type: 'normal', text: 'Hello' }];

    await replaceFormattedDocumentContent(docs, 'doc123', sections);

    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
    const requests = docs.documents.batchUpdate.mock.calls[0].arguments[0].requestBody.requests;
    assert.equal(requests.some((request) => request.deleteContentRange), false);
    assert.ok(requests[0].insertText);
    assert.ok(requests[0].insertText.text.startsWith('Hello'));
    assert.strictEqual(requests[0].insertText.location.index, 1);
  });

  it('throws UserError when lastElement.endIndex is missing and does not write', async () => {
    const docs = {
      documents: {
        get: mock.fn(async () => ({
          data: { revisionId: 'rev1', body: { content: [{}] } },
        })),
        batchUpdate: mock.fn(async () => ({ data: {} })),
      },
    };

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', [{ type: 'normal', text: 'Hello' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Document endIndex was not returned; cannot replace content safely.'
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('keeps a rejected combined batch as a single batchUpdate call', async () => {
    const docs = makeReplaceDocs({
      batchUpdate: async () => {
        const error = new Error('Invalid requests');
        error.code = 400;
        throw error;
      },
    });

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', [{ type: 'normal', text: 'Hello' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.ok(error.message.includes('Invalid request sent to Google Docs API'));
        return true;
      }
    );
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
  });

  it('names data loss when a later chunk fails after a committed delete', async () => {
    let callCount = 0;
    const docs = makeReplaceDocs({
      batchUpdate: async () => {
        callCount += 1;
        if (callCount === 1) {
          return { data: { writeControl: { requiredRevisionId: 'rev2' } } };
        }
        const error = new Error('backend');
        error.code = 500;
        throw error;
      },
    });
    const sections = normalSections(50);

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', sections),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.ok(error.message.includes('Original content was already deleted and was not restored.'));
        return true;
      }
    );
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 2);
    const firstArgs = docs.documents.batchUpdate.mock.calls[0].arguments[0];
    const secondArgs = docs.documents.batchUpdate.mock.calls[1].arguments[0];
    assert.ok(firstArgs.requestBody.requests.some((request) => request.deleteContentRange));
    assert.ok(secondArgs.requestBody.requests.some((request) => request.insertText));
    const expectedLaterInsert = buildFormattedContentRequests(sections, 1).textRequests[49];
    const laterInsert = secondArgs.requestBody.requests.find((request) => request.insertText);
    assert.strictEqual(
      laterInsert.insertText.location.index,
      expectedLaterInsert.insertText.location.index
    );
    assert.strictEqual(firstArgs.requestBody.writeControl.requiredRevisionId, 'rev1');
    assert.strictEqual(secondArgs.requestBody.writeControl.requiredRevisionId, 'rev2');
  });

  it('does not name data loss when the first split chunk rejects', async () => {
    const docs = makeReplaceDocs({
      batchUpdate: async () => {
        const error = new Error('Invalid requests');
        error.code = 400;
        throw error;
      },
    });

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', normalSections(50)),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.ok(error.message.includes('Invalid request sent to Google Docs API'));
        assert.equal(
          error.message.includes('Original content was already deleted and was not restored.'),
          false
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
  });

  it('throws UserError when revisionId is missing and does not write', async () => {
    const docs = makeReplaceDocs({ includeRevisionId: false });

    await assert.rejects(
      () => replaceFormattedDocumentContent(docs, 'doc123', [{ type: 'normal', text: 'Hello' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Document revision ID was not returned; cannot pin this write.'
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });
});

describe('updateFormattedDocumentSection', () => {
  it('sends delete then insert in one batchUpdate pinned to the get revisionId', async () => {
    const docs = makeSectionDocs();
    const sections = [{ type: 'normal', text: 'Hello' }];

    await updateFormattedDocumentSection(docs, 'doc123', 'Introduction', sections);

    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.deepStrictEqual(docs.documents.get.mock.calls[0].arguments[0], {
      documentId: 'doc123',
      fields: 'revisionId,body(content(paragraph(elements(textRun(content)),paragraphStyle(namedStyleType)),startIndex,endIndex))',
    });
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
    const callArgs = docs.documents.batchUpdate.mock.calls[0].arguments[0];
    const requests = callArgs.requestBody.requests;
    const deleteIdx = requests.findIndex((request) => request.deleteContentRange);
    const insertIdx = requests.findIndex((request) => request.insertText);
    assert.ok(deleteIdx >= 0);
    assert.ok(insertIdx > deleteIdx);
    assert.deepStrictEqual(requests[deleteIdx].deleteContentRange.range, {
      startIndex: 15,
      endIndex: 39,
    });
    assert.ok(requests[insertIdx].insertText.text.startsWith('Hello'));
    assert.strictEqual(requests[insertIdx].insertText.location.index, 15);
    const lastInsertIdx = requests.reduce((last, request, i) => (request.insertText ? i : last), -1);
    const styleIdx = firstStyleIndex(requests);
    assert.ok(styleIdx > lastInsertIdx);
    assert.strictEqual(callArgs.requestBody.writeControl.requiredRevisionId, 'rev1');
  });

  it('throws UserError for empty content without get or batchUpdate', async () => {
    const docs = makeSectionDocs();

    await assert.rejects(
      () => updateFormattedDocumentSection(docs, 'doc123', 'Introduction', []),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'content must not be empty.');
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 0);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('throws UserError when the heading is not found and does not write', async () => {
    const docs = makeSectionDocs();

    await assert.rejects(
      () => updateFormattedDocumentSection(docs, 'doc123', 'Missing Heading', [{ type: 'normal', text: 'Hello' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Could not find a heading matching "Missing Heading" in the document. Make sure the heading text is an exact match.'
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('throws UserError when revisionId is missing and does not write', async () => {
    const docs = makeSectionDocs({ includeRevisionId: false });

    await assert.rejects(
      () => updateFormattedDocumentSection(docs, 'doc123', 'Introduction', [{ type: 'normal', text: 'Hello' }]),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'Document revision ID was not returned; cannot pin this write.'
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.get.mock.calls.length, 1);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });
});

describe('formatted-content tool wiring in server.ts', () => {
  it('routes the four formatted-content tools through helpers and leaves createDocument/createFromTemplate on documents.batchUpdate', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    for (const toolName of [
      'createFormattedDocument',
      'insertFormattedContent',
      'replaceDocumentContent',
      'updateDocumentSection',
    ]) {
      const body = toolBody(source, toolName);
      assert.equal(
        body.includes('docs.documents.batchUpdate'),
        false,
        `${toolName} still calls docs.documents.batchUpdate`
      );
    }
    assert.ok(toolBody(source, 'replaceDocumentContent').includes('replaceFormattedDocumentContent'));
    assert.ok(toolBody(source, 'updateDocumentSection').includes('updateFormattedDocumentSection'));
    assert.ok(toolBody(source, 'createFormattedDocument').includes('GDocsHelpers.createFormattedDocument'));
    assert.ok(toolBody(source, 'insertFormattedContent').includes('[...textRequests, ...styleRequests]'));
    assert.ok(toolBody(source, 'insertFormattedContent').includes('executeBatchUpdate'));
    assert.ok(toolBody(source, 'createDocument').includes('docs.documents.batchUpdate'));
    assert.ok(toolBody(source, 'createFromTemplate').includes('docs.documents.batchUpdate'));
  });
});

describe('buildFormattedContentRequests color', () => {
  it('expands #F00 through hexToRgbColor', () => {
    const { styleRequests } = buildFormattedContentRequests(
      [{ type: 'normal', text: 'Hi', color: '#F00' }],
      1
    );
    const style = styleRequests.find((request) => request.updateTextStyle);
    assert.ok(style);
    assert.deepStrictEqual(
      style.updateTextStyle.textStyle.foregroundColor.color.rgbColor,
      { red: 1, green: 0, blue: 0 }
    );
    assert.ok(style.updateTextStyle.fields.includes('foregroundColor'));
  });

  it('omits foregroundColor when color is invalid so a style request cannot 400', () => {
    const { styleRequests } = buildFormattedContentRequests(
      [{ type: 'normal', text: 'Hi', color: '#XYZ', bold: true }],
      1
    );
    const style = styleRequests.find((request) => request.updateTextStyle);
    assert.ok(style);
    assert.strictEqual(style.updateTextStyle.textStyle.bold, true);
    assert.equal('foregroundColor' in style.updateTextStyle.textStyle, false);
    assert.equal(style.updateTextStyle.fields.includes('foregroundColor'), false);
  });

  it('does not emit updateTextStyle when the only style is an invalid or null color', () => {
    const invalid = buildFormattedContentRequests(
      [{ type: 'normal', text: 'Hi', color: 'not-a-color' }],
      1
    );
    const nulled = buildFormattedContentRequests(
      [{ type: 'normal', text: 'Hi', color: null }],
      1
    );
    assert.equal(invalid.styleRequests.some((request) => request.updateTextStyle), false);
    assert.equal(nulled.styleRequests.some((request) => request.updateTextStyle), false);
  });
});

describe('executeBatchUpdate insert-only split', () => {
  it('names leftover content when a later insert-only chunk fails', async () => {
    let callCount = 0;
    const docs = {
      documents: {
        batchUpdate: mock.fn(async () => {
          callCount += 1;
          if (callCount === 1) {
            return { data: { writeControl: { requiredRevisionId: 'rev2' } } };
          }
          const error = new Error('backend');
          error.code = 500;
          throw error;
        }),
      },
    };
    const requests = Array.from({ length: 51 }, (_, i) => ({
      insertText: { location: { index: 1 }, text: `s${i}` },
    }));

    await assert.rejects(
      () => executeBatchUpdate(docs, 'doc123', requests),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.ok(error.message.includes('Partial write'));
        assert.ok(error.message.includes('leftover'));
        assert.equal(
          error.message.includes('Original content was already deleted and was not restored.'),
          false
        );
        return true;
      }
    );
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 2);
  });
});

describe('createFormattedDocument', () => {
  function makeDriveAndDocs({ batchUpdate, deleteImpl } = {}) {
    const drive = {
      files: {
        create: mock.fn(async () => ({
          data: { id: 'doc1', name: 'T', webViewLink: 'http://x' },
        })),
        delete: mock.fn(deleteImpl ?? (async () => ({}))),
      },
    };
    const docs = {
      documents: {
        batchUpdate: mock.fn(batchUpdate ?? (async () => ({ data: {} }))),
      },
    };
    return { drive, docs };
  }

  it('deletes the Drive file when batchUpdate fails after create', async () => {
    const { drive, docs } = makeDriveAndDocs({
      batchUpdate: async () => {
        const error = new Error('backend');
        error.code = 500;
        throw error;
      },
    });

    await assert.rejects(
      () => createFormattedDocument(drive, docs, {
        title: 'T',
        content: [{ type: 'normal', text: 'Hi' }],
      }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.equal(error.message.includes('Leftover empty document was not deleted'), false);
        return true;
      }
    );
    assert.strictEqual(drive.files.create.mock.calls.length, 1);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
    assert.strictEqual(drive.files.delete.mock.calls.length, 1);
    assert.deepStrictEqual(drive.files.delete.mock.calls[0].arguments[0], {
      fileId: 'doc1',
      supportsAllDrives: true,
    });
  });

  it('includes leftover id when delete after failed batch also fails', async () => {
    const { drive, docs } = makeDriveAndDocs({
      batchUpdate: async () => {
        const error = new Error('backend');
        error.code = 500;
        throw error;
      },
      deleteImpl: async () => {
        throw new Error('delete failed');
      },
    });

    await assert.rejects(
      () => createFormattedDocument(drive, docs, {
        title: 'T',
        content: [{ type: 'normal', text: 'Hi' }],
      }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.ok(error.message.includes('Leftover empty document was not deleted (ID: doc1).'));
        return true;
      }
    );
    assert.strictEqual(drive.files.delete.mock.calls.length, 1);
  });

  it('does not delete when files.create fails', async () => {
    const drive = {
      files: {
        create: mock.fn(async () => {
          const error = new Error('not found');
          error.code = 404;
          throw error;
        }),
        delete: mock.fn(async () => ({})),
      },
    };
    const docs = {
      documents: {
        batchUpdate: mock.fn(async () => ({ data: {} })),
      },
    };

    await assert.rejects(
      () => createFormattedDocument(drive, docs, {
        title: 'T',
        content: [{ type: 'normal', text: 'Hi' }],
      })
    );
    assert.strictEqual(drive.files.delete.mock.calls.length, 0);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 0);
  });

  it('does not delete after a successful batch', async () => {
    const { drive, docs } = makeDriveAndDocs();

    const document = await createFormattedDocument(drive, docs, {
      title: 'T',
      content: [{ type: 'normal', text: 'Hi' }],
    });

    assert.strictEqual(document.id, 'doc1');
    assert.strictEqual(drive.files.delete.mock.calls.length, 0);
    assert.strictEqual(docs.documents.batchUpdate.mock.calls.length, 1);
  });
});
