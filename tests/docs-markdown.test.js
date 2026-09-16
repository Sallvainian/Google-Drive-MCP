// tests/docs-markdown.test.js
import {
  convertDocsJsonToMarkdown,
  markdownContentSourceFromDocumentTab,
} from '../dist/googleDocsApiHelpers.js';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function paragraph({ content, textStyle, bullet, namedStyleType }) {
  const p = {
    elements: [
      {
        textRun: {
          content,
          ...(textStyle ? { textStyle } : {}),
        },
      },
    ],
  };
  if (bullet) p.bullet = bullet;
  if (namedStyleType) p.paragraphStyle = { namedStyleType };
  return { paragraph: p };
}

function listDoc(paragraphs, nestingLevels) {
  const doc = {
    body: { content: paragraphs },
  };
  if (nestingLevels) {
    doc.lists = {
      k1: {
        listProperties: {
          nestingLevels,
        },
      },
    };
  }
  return doc;
}

function oneByOneTable(cellText) {
  return {
    body: {
      content: [
        {
          table: {
            tableRows: [
              {
                tableCells: [
                  {
                    content: [
                      {
                        paragraph: {
                          elements: [
                            { textRun: { content: cellText } },
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
      ],
    },
  };
}

describe('convertDocsJsonToMarkdown', () => {
  it('emits 1. for a DECIMAL list item', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [paragraph({ content: 'one\n', bullet: { listId: 'k1', nestingLevel: 0 } })],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. one'));
    assert.equal(md.includes('- one'), false);
  });

  it('treats omitted nestingLevel as 0 for a DECIMAL list item', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [paragraph({ content: 'one\n', bullet: { listId: 'k1' } })],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. one'));
    assert.equal(md.includes('- one'), false);
  });

  it('indents a nested ordered item two spaces and still uses 1.', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [
        paragraph({ content: 'one\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        paragraph({ content: 'nested\n', bullet: { listId: 'k1', nestingLevel: 1 } }),
      ],
      [{ glyphType: 'DECIMAL' }, { glyphType: 'ALPHA' }],
    ));
    assert.match(md, /^  1\. nested/m);
  });

  it('keeps two-space indent when the first paragraph is a nested list item', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [paragraph({ content: 'nested\n', bullet: { listId: 'k1', nestingLevel: 1 } })],
      [{ glyphType: 'DECIMAL' }, { glyphType: 'ALPHA' }],
    ));
    assert.match(md, /^  1\. nested/m);
  });

  it('emits - for an unordered NONE glyphType', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [paragraph({ content: 'item\n', bullet: { listId: 'k1', nestingLevel: 0 } })],
      [{ glyphType: 'NONE' }],
    ));
    assert.ok(md.includes('- item'));
    assert.equal(md.includes('1. item'), false);
  });

  it('strips a trailing newline before wrapping a bold run', () => {
    const md = convertDocsJsonToMarkdown({
      body: {
        content: [
          paragraph({ content: 'hello\n', textStyle: { bold: true } }),
        ],
      },
    });
    assert.strictEqual(md, '**hello**');
    assert.equal(md.includes('**hello\n**'), false);
    assert.equal(md.includes('\n**'), false);
  });

  it('does not wrap a bold run that is only a newline', () => {
    const md = convertDocsJsonToMarkdown({
      body: {
        content: [
          {
            paragraph: {
              elements: [
                { textRun: { content: 'hello', textStyle: { bold: true } } },
                { textRun: { content: '\n', textStyle: { bold: true } } },
              ],
            },
          },
        ],
      },
    });
    assert.strictEqual(md, '**hello**');
    assert.equal(md.includes('****'), false);
  });

  it('backslash-escapes a pipe in a table cell', () => {
    const md = convertDocsJsonToMarkdown(oneByOneTable('a|b'));
    assert.ok(md.split('\n').includes('| a\\|b |'));
  });

  it('escapes a backslash in a table cell before pipes', () => {
    const md = convertDocsJsonToMarkdown(oneByOneTable('a\\b'));
    assert.ok(md.split('\n').includes('| a\\\\b |'));
  });

  it('escapes backslash then pipe when a cell contains both', () => {
    const md = convertDocsJsonToMarkdown(oneByOneTable('a\\|b'));
    assert.ok(md.split('\n').includes('| a\\\\\\|b |'));
  });

  it('does not escape a pipe in a non-list paragraph', () => {
    const md = convertDocsJsonToMarkdown({
      body: {
        content: [paragraph({ content: 'a|b\n' })],
      },
    });
    assert.strictEqual(md, 'a|b');
  });

  it('falls back to unordered when the lists map is missing', () => {
    const md = convertDocsJsonToMarkdown({
      body: {
        content: [
          paragraph({ content: 'item\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        ],
      },
    });
    assert.ok(md.includes('- item'));
    assert.equal(md.includes('1. item'), false);
  });

  it('emits 1. for a DECIMAL item even when namedStyleType is HEADING_1', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [paragraph({
        content: 'text\n',
        bullet: { listId: 'k1', nestingLevel: 0 },
        namedStyleType: 'HEADING_1',
      })],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. text'));
    assert.equal(md.includes('# text'), false);
  });

  it('emits a list marker for a whitespace-only list paragraph so the list is not split', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [
        paragraph({ content: 'one\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        paragraph({ content: '   \n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        paragraph({ content: 'two\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
      ],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. one\n1. \n1. two'));
  });

  it('does not lazy-continue a following paragraph into a list item', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [
        paragraph({ content: 'one\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        paragraph({ content: 'following\n' }),
      ],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. one\n\nfollowing'));
    assert.equal(md.includes('1. one\nfollowing'), false);
  });

  it('keeps consecutive list items in one list', () => {
    const md = convertDocsJsonToMarkdown(listDoc(
      [
        paragraph({ content: 'one\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        paragraph({ content: 'two\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
      ],
      [{ glyphType: 'DECIMAL' }],
    ));
    assert.ok(md.includes('1. one\n1. two'));
  });
});

describe('markdownContentSourceFromDocumentTab', () => {
  it('passes document lists into convertDocsJsonToMarkdown when documentTab.lists is missing', () => {
    const documentTab = {
      body: {
        content: [
          paragraph({ content: 'item\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        ],
      },
    };
    const documentLists = {
      k1: { listProperties: { nestingLevels: [{ glyphType: 'DECIMAL' }] } },
    };
    const source = markdownContentSourceFromDocumentTab(documentTab, documentLists);
    assert.strictEqual(source.lists, documentLists);
    const md = convertDocsJsonToMarkdown(source);
    assert.ok(md.includes('1. item'));
    assert.equal(md.includes('- item'), false);
  });

  it('prefers documentTab.lists when present', () => {
    const tabLists = {
      k1: { listProperties: { nestingLevels: [{ glyphType: 'DECIMAL' }] } },
    };
    const documentTab = {
      body: {
        content: [
          paragraph({ content: 'item\n', bullet: { listId: 'k1', nestingLevel: 0 } }),
        ],
      },
      lists: tabLists,
    };
    const documentLists = {
      k1: { listProperties: { nestingLevels: [{ glyphType: 'NONE' }] } },
    };
    const source = markdownContentSourceFromDocumentTab(documentTab, documentLists);
    assert.strictEqual(source.lists, tabLists);
    const md = convertDocsJsonToMarkdown(source);
    assert.ok(md.includes('1. item'));
    assert.equal(md.includes('- item'), false);
  });
});

describe('readGoogleDoc markdown wiring in server.ts', () => {
  it('calls GDocsHelpers.convertDocsJsonToMarkdown and markdownContentSourceFromDocumentTab', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.ok(source.includes('GDocsHelpers.convertDocsJsonToMarkdown'));
    assert.ok(source.includes('GDocsHelpers.markdownContentSourceFromDocumentTab(targetTab.documentTab, res.data.lists)'));
    assert.ok(source.includes('contentSource = { body: targetTab.documentTab.body }'));
    assert.equal(/function convertDocsJsonToMarkdown/.test(source), false);
    assert.equal(/function convertParagraphToMarkdown/.test(source), false);
    assert.equal(/function convertTextRunToMarkdown/.test(source), false);
    assert.equal(/function convertTableToMarkdown/.test(source), false);
    assert.equal(/function markdownContentSourceFromDocumentTab/.test(source), false);
  });
});
