// tests/find-element.test.js
import {
  FIND_ELEMENT_FIELDS,
  SUGGESTIONS_VIEW_MODE,
  findElements,
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

const mockDocsWith = (content) => ({
  documents: {
    get: mock.fn(async () => ({ data: { body: { content } } })),
  },
});

function lastGetParams(docs) {
  assert.strictEqual(docs.documents.get.mock.calls.length >= 1, true);
  return docs.documents.get.mock.calls[0].arguments[0];
}

describe('FIND_ELEMENT_FIELDS', () => {
  it('is the upstream body mask with no tabs wrap', () => {
    assert.strictEqual(
      FIND_ELEMENT_FIELDS,
      'body(content(startIndex,endIndex,table(rows,columns,tableRows(tableCells(content(paragraph(elements(startIndex,endIndex,textRun(content))))))),paragraph(elements(startIndex,endIndex,textRun(content)))))'
    );
    assert.equal(FIND_ELEMENT_FIELDS.includes('tabs('), false);
    assert.equal(FIND_ELEMENT_FIELDS.includes('tabId'), false);
    assert.equal(FIND_ELEMENT_FIELDS.includes('includeTabsContent'), false);
  });
});

describe('findElements', () => {
  describe('documents.get contract', () => {
    it('requests FIND_ELEMENT_FIELDS and PREVIEW_WITHOUT_SUGGESTIONS', async () => {
      const docs = mockDocsWith([
        { paragraph: { elements: [{ startIndex: 1, endIndex: 6, textRun: { content: 'abcd\n' } }] } },
      ]);
      await findElements(docs, 'doc123', { textQuery: 'ab' });
      assert.deepStrictEqual(lastGetParams(docs), {
        documentId: 'doc123',
        fields: FIND_ELEMENT_FIELDS,
        suggestionsViewMode: SUGGESTIONS_VIEW_MODE,
      });
      assert.strictEqual(lastGetParams(docs).suggestionsViewMode, 'PREVIEW_WITHOUT_SUGGESTIONS');
      assert.equal(lastGetParams(docs).fields.includes('tabs('), false);
      assert.equal(Object.prototype.hasOwnProperty.call(lastGetParams(docs), 'includeTabsContent'), false);
      assert.equal(Object.prototype.hasOwnProperty.call(lastGetParams(docs), 'tabId'), false);
    });
  });

  describe('textQuery', () => {
    it('returns every non-overlapping occurrence with exact index ranges', async () => {
      // "alpha beta alpha\n" occupies indices 1..18; "alpha" at offsets 0 and 11.
      const docs = mockDocsWith([
        {
          paragraph: {
            elements: [
              { startIndex: 1, endIndex: 18, textRun: { content: 'alpha beta alpha\n' } },
            ],
          },
        },
      ]);

      const hits = await findElements(docs, 'doc123', { textQuery: 'alpha' });
      assert.strictEqual(hits.length, 2);
      assert.deepStrictEqual(hits[0], { type: 'text', instance: 1, startIndex: 1, endIndex: 6, text: 'alpha' });
      assert.deepStrictEqual(hits[1], { type: 'text', instance: 2, startIndex: 12, endIndex: 17, text: 'alpha' });
      for (const h of hits) assert.strictEqual(h.endIndex - h.startIndex, 'alpha'.length);
    });

    it('does not emit overlapping hits for a self-overlapping needle', async () => {
      // 'aaa' can match 'aa' at offsets 0 and 1; non-overlapping advance (from = at + len)
      // yields a single width-2 hit. from = at + 1 would yield two.
      const docs = mockDocsWith([
        {
          paragraph: {
            elements: [
              { startIndex: 1, endIndex: 4, textRun: { content: 'aaa' } },
            ],
          },
        },
      ]);

      const hits = await findElements(docs, 'doc123', { textQuery: 'aa' });
      assert.deepStrictEqual(hits, [
        { type: 'text', instance: 1, startIndex: 1, endIndex: 3, text: 'aa' },
      ]);
      assert.strictEqual(hits[0].endIndex - hits[0].startIndex, 2);
    });

    it('matches a phrase split across runs by a mid-phrase style boundary (cross-run)', async () => {
      const docs = mockDocsWith([
        {
          paragraph: {
            elements: [
              { startIndex: 1, endIndex: 3, textRun: { content: 'al' } },
              { startIndex: 3, endIndex: 7, textRun: { content: 'pha\n' } },
            ],
          },
        },
      ]);

      const hits = await findElements(docs, 'doc123', { textQuery: 'alpha' });
      assert.deepStrictEqual(hits, [
        { type: 'text', instance: 1, startIndex: 1, endIndex: 6, text: 'alpha' },
      ]);
      assert.strictEqual(hits[0].endIndex - hits[0].startIndex, 'alpha'.length);
    });

    it('does NOT span an inline object between two runs', async () => {
      const docs = mockDocsWith([
        {
          paragraph: {
            elements: [
              { startIndex: 1, endIndex: 3, textRun: { content: 'ab' } },
              { startIndex: 3, endIndex: 4, inlineObjectElement: { inlineObjectId: 'img1' } },
              { startIndex: 4, endIndex: 7, textRun: { content: 'cd\n' } },
            ],
          },
        },
      ]);
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'bc' }), []);
      const cd = await findElements(docs, 'doc123', { textQuery: 'cd' });
      assert.deepStrictEqual(cd, [{ type: 'text', instance: 1, startIndex: 4, endIndex: 6, text: 'cd' }]);
      assert.strictEqual(cd[0].endIndex - cd[0].startIndex, 'cd'.length);
    });

    it('does NOT span a numeric index gap between runs with no intervening element', async () => {
      const docs = mockDocsWith([
        {
          paragraph: {
            elements: [
              { startIndex: 1, endIndex: 3, textRun: { content: 'ab' } },
              { startIndex: 5, endIndex: 8, textRun: { content: 'cd\n' } },
            ],
          },
        },
      ]);
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'bc' }), []);
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'cd' }), [
        { type: 'text', instance: 1, startIndex: 5, endIndex: 7, text: 'cd' },
      ]);
    });

    it('does NOT match across a paragraph boundary', async () => {
      const docs = mockDocsWith([
        { paragraph: { elements: [{ startIndex: 1, endIndex: 7, textRun: { content: 'alpha\n' } }] } },
        { paragraph: { elements: [{ startIndex: 7, endIndex: 12, textRun: { content: 'beta\n' } }] } },
      ]);
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'alpha\nbeta' }), []);
    });

    it('indexes table-cell text by its run start, not by concatenated-text offset', async () => {
      const docs = mockDocsWith([
        { paragraph: { elements: [{ startIndex: 1, endIndex: 7, textRun: { content: 'alpha\n' } }] } },
        {
          startIndex: 100,
          endIndex: 120,
          table: {
            rows: 1,
            columns: 1,
            tableRows: [
              {
                tableCells: [
                  {
                    content: [
                      {
                        paragraph: {
                          elements: [
                            { startIndex: 103, endIndex: 108, textRun: { content: 'beta\n' } },
                          ],
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          },
        },
      ]);

      const hits = await findElements(docs, 'doc123', { textQuery: 'beta' });
      assert.strictEqual(hits.length, 1);
      assert.deepStrictEqual(hits[0], { type: 'text', instance: 1, startIndex: 103, endIndex: 107, text: 'beta' });
      assert.strictEqual(hits[0].endIndex - hits[0].startIndex, 'beta'.length);
    });

    it('returns an empty array when the text is absent', async () => {
      const docs = mockDocsWith([
        { paragraph: { elements: [{ startIndex: 1, endIndex: 6, textRun: { content: 'abcd\n' } }] } },
      ]);
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'zzz' }), []);
    });

    it('returns an empty array when body content is missing', async () => {
      const docs = {
        documents: {
          get: mock.fn(async () => ({ data: {} })),
        },
      };
      assert.deepStrictEqual(await findElements(docs, 'doc123', { textQuery: 'alpha' }), []);
    });
  });

  describe('elementType listing', () => {
    it('lists tables with their range and a size preview', async () => {
      const docs = mockDocsWith([
        { startIndex: 20, endIndex: 40, table: { rows: 2, columns: 3, tableRows: [] } },
      ]);
      const hits = await findElements(docs, 'doc123', { elementType: 'table' });
      assert.deepStrictEqual(hits, [
        { type: 'table', startIndex: 20, endIndex: 40, text: 'table 2x3' },
      ]);
    });

    it('lists top-level body paragraphs (newline-stripped preview) and not table-cell paragraphs', async () => {
      const docs = mockDocsWith([
        { startIndex: 1, endIndex: 17, paragraph: { elements: [{ startIndex: 1, endIndex: 17, textRun: { content: 'intro paragraph\n' } }] } },
        {
          startIndex: 100,
          endIndex: 120,
          table: {
            rows: 1,
            columns: 1,
            tableRows: [
              { tableCells: [{ content: [{ paragraph: { elements: [{ startIndex: 103, endIndex: 113, textRun: { content: 'cell para\n' } }] } }] }] },
            ],
          },
        },
      ]);
      const hits = await findElements(docs, 'doc123', { elementType: 'paragraph' });
      assert.deepStrictEqual(hits, [
        { type: 'paragraph', startIndex: 1, endIndex: 17, text: 'intro paragraph' },
      ]);
    });

    it('truncates paragraph preview to 120 characters after stripping a trailing newline', async () => {
      const long = `${'a'.repeat(130)}\n`;
      const docs = mockDocsWith([
        {
          startIndex: 1,
          endIndex: 132,
          paragraph: { elements: [{ startIndex: 1, endIndex: 132, textRun: { content: long } }] },
        },
      ]);
      const hits = await findElements(docs, 'doc123', { elementType: 'paragraph' });
      assert.strictEqual(hits.length, 1);
      assert.strictEqual(hits[0].text, 'a'.repeat(120));
    });

    it('rejects unsupported element types (list/image) regardless of textQuery', async () => {
      const docs = mockDocsWith([]);
      const unsupported = 'elementType "image" is not supported. Omit elementType and pass textQuery to locate content by text.';
      const unsupportedList = 'elementType "list" is not supported. Omit elementType and pass textQuery to locate content by text.';
      await assert.rejects(
        () => findElements(docs, 'doc123', { elementType: 'image', textQuery: 'x' }),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, unsupported);
          return true;
        }
      );
      await assert.rejects(
        () => findElements(docs, 'doc123', { elementType: 'list' }),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, unsupportedList);
          return true;
        }
      );
    });
  });

  describe('combined elementType + textQuery', () => {
    it('returns the table listing then the text matches', async () => {
      const docs = mockDocsWith([
        { paragraph: { elements: [{ startIndex: 1, endIndex: 14, textRun: { content: 'foundme here\n' } }] } },
        { startIndex: 50, endIndex: 70, table: { rows: 1, columns: 1, tableRows: [] } },
      ]);
      const hits = await findElements(docs, 'doc123', { elementType: 'table', textQuery: 'foundme' });
      assert.deepStrictEqual(hits, [
        { type: 'table', startIndex: 50, endIndex: 70, text: 'table 1x1' },
        { type: 'text', instance: 1, startIndex: 1, endIndex: 8, text: 'foundme' },
      ]);
    });
  });

  it('requires at least one of textQuery or elementType', async () => {
    const docs = mockDocsWith([]);
    await assert.rejects(
      () => findElements(docs, 'doc123', {}),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'findElement requires at least one of "textQuery" or "elementType".');
        return true;
      }
    );
  });

  it('throws UserError for 404', async () => {
    const docs = {
      documents: {
        get: mock.fn(async () => {
          const error = new Error('Not found');
          error.code = 404;
          throw error;
        }),
      },
    };
    await assert.rejects(
      () => findElements(docs, 'doc123', { textQuery: 'alpha' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Document not found (ID: doc123).');
        return true;
      }
    );
  });

  it('throws UserError for 403', async () => {
    const docs = {
      documents: {
        get: mock.fn(async () => {
          const error = new Error('Forbidden');
          error.code = 403;
          throw error;
        }),
      },
    };
    await assert.rejects(
      () => findElements(docs, 'doc123', { textQuery: 'alpha' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'Permission denied for document doc123.');
        return true;
      }
    );
  });

  it('throws Error (not UserError) for other get failures', async () => {
    const docs = {
      documents: {
        get: mock.fn(async () => {
          const error = new Error('backend');
          error.code = 500;
          throw error;
        }),
      },
    };
    await assert.rejects(
      () => findElements(docs, 'doc123', { textQuery: 'alpha' }),
      (error) => {
        assert.ok(error instanceof Error);
        assert.equal(error instanceof UserError, false);
        assert.ok(error.message.startsWith('Failed to retrieve document for findElement:'));
        assert.ok(error.message.includes('backend'));
        return true;
      }
    );
  });
});

describe('findElement tool contract', () => {
  it('keeps the catalog name and upstream execute/return strings', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    const body = toolBody(source, 'findElement');
    assert.ok(body.includes("name: 'findElement'"));
    assert.equal(body.includes('(Not Implemented)'), false);
    assert.equal(body.includes('NotImplementedError'), false);
    assert.ok(body.includes('.min(1)'));
    assert.ok(body.includes('GDocsHelpers.findElements'));
    assert.ok(body.includes('getDocsClient'));
    assert.ok(body.includes('No matching elements found (textQuery=${args.textQuery ?? \'none\'}, elementType=${args.elementType ?? \'none\'}).'));
    assert.ok(body.includes('JSON.stringify({ count: found.length, elements: found }, null, 2)'));
    assert.ok(body.includes('if (error instanceof UserError) throw error'));
    assert.ok(body.includes('Failed to find elements: ${error.message || \'Unknown error\'}'));
    assert.equal(body.includes('throw new Error'), false);
  });
});
