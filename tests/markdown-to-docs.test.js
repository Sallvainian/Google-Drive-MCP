// tests/markdown-to-docs.test.js
import {
  convertMarkdownToRequests,
  convertMarkdownTokensToRequests,
} from '../dist/markdown-transformer/markdownToDocs.js';
import { MarkdownConversionError } from '../dist/types.js';
import assert from 'node:assert';
import { describe, it } from 'node:test';


// ============================================================
// Markdown -> Google Docs Requests
// ============================================================

describe('Markdown to Docs Conversion', () => {
  describe('Basic Text Formatting', () => {
    it('should convert bold text', () => {
      const requests = convertMarkdownToRequests('**bold text**', 1);

      const insertReq = requests.find((r) => r.insertText);
      assert.ok(insertReq);
      assert.strictEqual(insertReq.insertText.text, 'bold text');

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.bold, true);
    });

    it('should convert italic text', () => {
      const requests = convertMarkdownToRequests('*italic text*', 1);

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.italic, true);
    });

    it('should convert strikethrough text', () => {
      const requests = convertMarkdownToRequests('~~strikethrough text~~', 1);

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.strikethrough, true);
    });

    it('should convert nested bold and italic', () => {
      const requests = convertMarkdownToRequests('***bold italic***', 1);

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.bold, true);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.italic, true);
    });

    it('should style inline code as monospace', () => {
      const requests = convertMarkdownToRequests('Use `inline_code` here', 1);

      const styleReqs = requests.filter((r) => r.updateTextStyle);
      const codeStyleReq = styleReqs.find(
        (r) => r.updateTextStyle.textStyle.weightedFontFamily?.fontFamily === 'Roboto Mono'
      );
      assert.ok(codeStyleReq);
    });
  });

  describe('Links', () => {
    it('should convert basic links', () => {
      const requests = convertMarkdownToRequests('[link text](https://example.com)', 1);

      const insertReq = requests.find((r) => r.insertText);
      assert.ok(insertReq);
      assert.strictEqual(insertReq.insertText.text, 'link text');

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.textStyle.link.url, 'https://example.com');
    });
  });

  describe('Headings', () => {
    it('should convert H1', () => {
      const requests = convertMarkdownToRequests('# Heading 1', 1);

      const insertReq = requests.find((r) => r.insertText && r.insertText.text === 'Heading 1');
      assert.ok(insertReq);

      const paraReq = requests.find((r) => r.updateParagraphStyle);
      assert.ok(paraReq);
      assert.strictEqual(paraReq.updateParagraphStyle.paragraphStyle.namedStyleType, 'HEADING_1');
    });

    it('should convert H2', () => {
      const requests = convertMarkdownToRequests('## Heading 2', 1);

      const paraReq = requests.find((r) => r.updateParagraphStyle);
      assert.ok(paraReq);
      assert.strictEqual(paraReq.updateParagraphStyle.paragraphStyle.namedStyleType, 'HEADING_2');
    });

    it('should convert H3', () => {
      const requests = convertMarkdownToRequests('### Heading 3', 1);

      const paraReq = requests.find((r) => r.updateParagraphStyle);
      assert.ok(paraReq);
      assert.strictEqual(paraReq.updateParagraphStyle.paragraphStyle.namedStyleType, 'HEADING_3');
    });
  });

  describe('firstHeadingAsTitle option', () => {
    it('should style the first H1 as TITLE when enabled', () => {
      const requests = convertMarkdownToRequests(
        '# My Document Title\n\nSome body text.',
        1,
        undefined,
        {
          firstHeadingAsTitle: true,
        }
      );

      const paraReqs = requests.filter((r) => r.updateParagraphStyle);
      const titleReq = paraReqs.find(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'TITLE'
      );
      assert.ok(titleReq);

      // Should NOT have a HEADING_1
      const h1Req = paraReqs.find(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'HEADING_1'
      );
      assert.equal(h1Req, undefined);
    });

    it('should only convert the first H1 to TITLE, not subsequent H1s', () => {
      const markdown = '# Title\n\n# Second H1\n\nSome text.';
      const requests = convertMarkdownToRequests(markdown, 1, undefined, {
        firstHeadingAsTitle: true,
      });

      const paraReqs = requests.filter((r) => r.updateParagraphStyle);
      const titleReqs = paraReqs.filter(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'TITLE'
      );
      const h1Reqs = paraReqs.filter(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'HEADING_1'
      );

      assert.strictEqual(titleReqs.length, 1);
      assert.strictEqual(h1Reqs.length, 1);
    });

    it('should leave H1 as HEADING_1 when option is disabled (default)', () => {
      const requests = convertMarkdownToRequests('# Heading 1', 1);

      const paraReq = requests.find((r) => r.updateParagraphStyle);
      assert.ok(paraReq);
      assert.strictEqual(paraReq.updateParagraphStyle.paragraphStyle.namedStyleType, 'HEADING_1');
    });

    it('should not affect H2+ headings when enabled', () => {
      const markdown = '## Section\n\n### Subsection';
      const requests = convertMarkdownToRequests(markdown, 1, undefined, {
        firstHeadingAsTitle: true,
      });

      const paraReqs = requests.filter((r) => r.updateParagraphStyle);
      const titleReqs = paraReqs.filter(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'TITLE'
      );
      assert.strictEqual(titleReqs.length, 0);

      const h2 = paraReqs.find(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'HEADING_2'
      );
      const h3 = paraReqs.find(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'HEADING_3'
      );
      assert.ok(h2);
      assert.ok(h3);
    });

    it('should handle a full document with title, headings, and lists', () => {
      const markdown = [
        '# Project Plan',
        '',
        '## Overview',
        '',
        'This is the overview.',
        '',
        '## Tasks',
        '',
        '- Task 1',
        '- Task 2',
      ].join('\n');

      const requests = convertMarkdownToRequests(markdown, 1, undefined, {
        firstHeadingAsTitle: true,
      });

      const paraReqs = requests.filter((r) => r.updateParagraphStyle);
      const titleReqs = paraReqs.filter(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'TITLE'
      );
      const h2Reqs = paraReqs.filter(
        (r) => r.updateParagraphStyle.paragraphStyle.namedStyleType === 'HEADING_2'
      );

      assert.strictEqual(titleReqs.length, 1);
      assert.strictEqual(h2Reqs.length, 2);
    });
  });

  describe('Lists', () => {
    it('should convert bullet lists', () => {
      const requests = convertMarkdownToRequests('- Item 1\n- Item 2\n- Item 3', 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      assert.strictEqual(bulletReqs.length, 1);
      assert.strictEqual(bulletReqs[0].createParagraphBullets.bulletPreset, 'BULLET_DISC_CIRCLE_SQUARE');
    });

    it('should convert numbered lists', () => {
      const requests = convertMarkdownToRequests('1. Item 1\n2. Item 2\n3. Item 3', 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      assert.strictEqual(bulletReqs.length, 1);
      assert.strictEqual(
        bulletReqs[0].createParagraphBullets.bulletPreset,
        'NUMBERED_DECIMAL_ALPHA_ROMAN'
      );
    });

    it('should preserve nested list levels with leading tabs', () => {
      const requests = convertMarkdownToRequests('- Parent\n  - Child', 1);

      const insertReqs = requests.filter((r) => r.insertText);
      assert.strictEqual(insertReqs.some((r) => r.insertText.text.includes('Parent')), true);
      assert.strictEqual(insertReqs.some((r) => r.insertText.text === '\t'), true);
      assert.strictEqual(insertReqs.some((r) => r.insertText.text.includes('Child')), true);
    });

    it('should insert multiple tabs for deeply nested lists (3 levels)', () => {
      const markdown = '- Level 0\n  - Level 1\n    - Level 2';
      const requests = convertMarkdownToRequests(markdown, 1);

      const insertReqs = requests.filter((r) => r.insertText);
      // Level 0 has no tab, Level 1 has 1 tab, Level 2 has 2 tabs
      assert.strictEqual(insertReqs.some((r) => r.insertText.text === '\t\t'), true);
      assert.strictEqual(insertReqs.some((r) => r.insertText.text.includes('Level 2')), true);
    });

    it('should use ordered preset for nested ordered list inside bullets', () => {
      const markdown = '- Bullet parent\n  1. Ordered child 1\n  2. Ordered child 2';
      const requests = convertMarkdownToRequests(markdown, 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      const presets = bulletReqs.map((r) => r.createParagraphBullets.bulletPreset);
      assert.ok(presets.includes('BULLET_DISC_CIRCLE_SQUARE'));
      assert.ok(presets.includes('NUMBERED_DECIMAL_ALPHA_ROMAN'));
    });

    it('should use bullet preset for nested bullets inside ordered list', () => {
      const markdown = '1. Ordered parent\n  - Bullet child 1\n  - Bullet child 2';
      const requests = convertMarkdownToRequests(markdown, 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      const presets = bulletReqs.map((r) => r.createParagraphBullets.bulletPreset);
      assert.ok(presets.includes('NUMBERED_DECIMAL_ALPHA_ROMAN'));
      assert.ok(presets.includes('BULLET_DISC_CIRCLE_SQUARE'));
    });

    it('should produce separate bullet requests for mixed nested list types', () => {
      const markdown = '- Parent\n  1. Child\n- Parent 2';
      const requests = convertMarkdownToRequests(markdown, 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      // Bullet and ordered are different presets so they cannot merge
      assert.ok(bulletReqs.length >= 2);
    });

    it('should merge sibling items of the same type even around nested sub-lists', () => {
      // Both "Parent 1" and "Parent 2" are BULLET_DISC_CIRCLE_SQUARE at level 0.
      // The ordered sub-list between them is a different preset.
      const markdown = '- Parent 1\n  1. Ordered child\n- Parent 2';
      const requests = convertMarkdownToRequests(markdown, 1);

      const allText = requests
        .filter((r) => r.insertText)
        .map((r) => r.insertText.text)
        .join('');
      assert.ok(allText.includes('Parent 1'));
      assert.ok(allText.includes('Ordered child'));
      assert.ok(allText.includes('Parent 2'));
    });

    it('should convert markdown task lists to checkbox bullets', () => {
      const requests = convertMarkdownToRequests('- [x] done\n- [ ] todo', 1);

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      assert.strictEqual(bulletReqs.length, 1);
      assert.strictEqual(bulletReqs[0].createParagraphBullets.bulletPreset, 'BULLET_CHECKBOX');

      const allInsertedText = requests
        .filter((r) => r.insertText)
        .map((r) => r.insertText.text)
        .join('');
      assert.equal(allInsertedText.includes('[x]'), false);
      assert.equal(allInsertedText.includes('[ ]'), false);
    });

    it('should not let list bullet ranges bleed into following headings', () => {
      const requests = convertMarkdownToRequests('- Parent\n  1. Child\n\n## Next Heading', 1);

      const headingReq = requests.find(
        (r) => r.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'HEADING_2'
      );
      assert.ok(headingReq);
      const headingStart = headingReq.updateParagraphStyle.range.startIndex;

      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      const overlappingBullet = bulletReqs.find((r) => {
        const { startIndex, endIndex } = r.createParagraphBullets.range;
        return headingStart >= startIndex && headingStart < endIndex;
      });
      assert.equal(overlappingBullet, undefined);
    });

    it('should not merge separate bullet lists with content between them', () => {
      const markdown = [
        '**Part 1: The Question**',
        '- Item A',
        '- Item B',
        '',
        '**Part 2: The Results**',
        '- Item C',
        '- Item D',
      ].join('\n');

      const requests = convertMarkdownToRequests(markdown, 1);
      const bulletReqs = requests.filter((r) => r.createParagraphBullets);

      // Should produce two separate bullet ranges, not one merged range
      assert.strictEqual(bulletReqs.length, 2);

      // The paragraph "Part 2: The Results" must not fall inside any bullet range
      const insertReqs = requests.filter((r) => r.insertText);
      let part2Index;
      for (const r of insertReqs) {
        if (r.insertText.text.includes('Part 2')) {
          part2Index = r.insertText.location.index;
          break;
        }
      }
      assert.ok(part2Index);

      for (const b of bulletReqs) {
        const { startIndex, endIndex } = b.createParagraphBullets.range;
        const inside = part2Index >= startIndex && part2Index < endIndex;
        assert.strictEqual(inside, false);
      }
    });

    it('should keep adjacent items in the same list merged', () => {
      const requests = convertMarkdownToRequests('- A\n- B\n- C', 1);
      const bulletReqs = requests.filter((r) => r.createParagraphBullets);
      assert.strictEqual(bulletReqs.length, 1);
    });
  });

  describe('Code Blocks', () => {
    it('should insert a 1x1 table for fenced code blocks', () => {
      const requests = convertMarkdownToRequests('```js\nconst x = 1;\nconsole.log(x);\n```', 1);

      // Should have an insertTable request
      const tableReqs = requests.filter((r) => r.insertTable);
      assert.strictEqual(tableReqs.length, 1);
      assert.strictEqual(tableReqs[0].insertTable.rows, 1);
      assert.strictEqual(tableReqs[0].insertTable.columns, 1);
    });

    it('should insert code text into the table cell', () => {
      const requests = convertMarkdownToRequests('```\nhello world\n```', 1);

      const insertReqs = requests.filter((r) => r.insertText);
      assert.strictEqual(insertReqs.some((r) => r.insertText.text.includes('hello world')), true);
    });

    it('should style code block text as monospace', () => {
      const requests = convertMarkdownToRequests('```\nconst x = 1;\nconsole.log(x);\n```', 1);

      const styleReqs = requests.filter((r) => r.updateTextStyle);
      const monospaceReqs = styleReqs.filter(
        (r) => r.updateTextStyle.textStyle.weightedFontFamily?.fontFamily === 'Roboto Mono'
      );
      assert.strictEqual(monospaceReqs.length, 1);
    });

    it('should style the table cell with background color', () => {
      const requests = convertMarkdownToRequests('```\ncode\n```', 1);

      const cellStyleReqs = requests.filter((r) => r.updateTableCellStyle);
      assert.strictEqual(cellStyleReqs.length, 1);

      const cellStyle = cellStyleReqs[0].updateTableCellStyle.tableCellStyle;
      assert.ok(cellStyle.backgroundColor);
      assert.ok(cellStyle.paddingTop);
      assert.ok(cellStyle.paddingBottom);
      assert.ok(cellStyle.paddingLeft);
      assert.ok(cellStyle.paddingRight);
    });

    it('should reference the actual table start (insertTable index + 1) in updateTableCellStyle', () => {
      // insertTable auto-inserts a preceding newline at T, so the table element
      // starts at T+1. The updateTableCellStyle must reference T+1, not T.
      const requests = convertMarkdownToRequests('```\ncode\n```', 1);

      const tableReq = requests.find((r) => r.insertTable);
      const cellStyleReq = requests.find((r) => r.updateTableCellStyle);

      assert.ok(tableReq);
      assert.ok(cellStyleReq);

      const insertTableIndex = tableReq.insertTable.location.index;
      const tableStartLocationIndex =
        cellStyleReq.updateTableCellStyle.tableRange.tableCellLocation.tableStartLocation
          .index;

      // The actual table start is insertTable target + 1 (preceding newline shifts it)
      assert.strictEqual(tableStartLocationIndex, insertTableIndex + 1);
    });

    it('should insert code text at correct offset from table start', () => {
      const requests = convertMarkdownToRequests('```\nhello\n```', 1);

      const tableReq = requests.find((r) => r.insertTable);
      const codeInsertReq = requests.find((r) => r.insertText && r.insertText.text === 'hello');

      assert.ok(tableReq);
      assert.ok(codeInsertReq);

      const tableIndex = tableReq.insertTable.location.index;
      const textIndex = codeInsertReq.insertText.location.index;

      // Cell content should be at table start + 4 (CELL_CONTENT_OFFSET)
      assert.strictEqual(textIndex, tableIndex + 4);
    });

    it('should handle multi-line code blocks', () => {
      const requests = convertMarkdownToRequests('```\nline1\nline2\nline3\n```', 1);

      const insertReqs = requests.filter((r) => r.insertText);
      const codeContent = insertReqs.find((r) => r.insertText.text === 'line1\nline2\nline3');
      assert.ok(codeContent);
    });

    it('should handle empty code blocks', () => {
      const requests = convertMarkdownToRequests('```\n```', 1);

      const tableReqs = requests.filter((r) => r.insertTable);
      assert.strictEqual(tableReqs.length, 1);

      // No code text insertion (empty block)
      const codeInsertReqs = requests.filter((r) => r.insertText && r.insertText.text !== '\n');
      assert.strictEqual(codeInsertReqs.length, 0);
    });

    it('should handle multiple code blocks in sequence', () => {
      const markdown = '```\ncode1\n```\n\n```\ncode2\n```';
      const requests = convertMarkdownToRequests(markdown, 1);

      const tableReqs = requests.filter((r) => r.insertTable);
      assert.strictEqual(tableReqs.length, 2);

      const cellStyleReqs = requests.filter((r) => r.updateTableCellStyle);
      assert.strictEqual(cellStyleReqs.length, 2);
    });

    it('should include tabId in table, text, and cell style requests when provided', () => {
      const requests = convertMarkdownToRequests('```\ncode\n```', 1, 'tab-code');

      const tableReq = requests.find((r) => r.insertTable);
      assert.strictEqual(tableReq.insertTable.location.tabId, 'tab-code');

      const codeInsertReq = requests.find((r) => r.insertText && r.insertText.text === 'code');
      assert.strictEqual(codeInsertReq.insertText.location.tabId, 'tab-code');

      const cellStyleReq = requests.find((r) => r.updateTableCellStyle);
      assert.strictEqual(
        cellStyleReq.updateTableCellStyle.tableRange.tableCellLocation.tableStartLocation.tabId,
        'tab-code'
      );
    });

    it('should not affect inline code styling', () => {
      const requests = convertMarkdownToRequests('Use `inline_code` here', 1);

      // Inline code should NOT create a table
      const tableReqs = requests.filter((r) => r.insertTable);
      assert.strictEqual(tableReqs.length, 0);

      // Inline code should still use text styling (monospace + green + background)
      const styleReqs = requests.filter((r) => r.updateTextStyle);
      const codeStyleReq = styleReqs.find(
        (r) => r.updateTextStyle.textStyle.weightedFontFamily?.fontFamily === 'Roboto Mono'
      );
      assert.ok(codeStyleReq);
    });

    it('should correctly track indices after a code block for following content', () => {
      const markdown = '```\ncode\n```\n\nFollowing text.';
      const requests = convertMarkdownToRequests(markdown, 1);

      // The following text should have valid insert locations
      const followingInsert = requests.find(
        (r) => r.insertText && r.insertText.text.includes('Following text')
      );
      assert.ok(followingInsert);
      assert.ok(followingInsert.insertText.location.index > 1);
    });
  });

  describe('Tables', () => {
    it('should convert a markdown table into Docs table requests', () => {
      const markdown = [
        '| Task ID | Task Name |',
        '| --- | --- |',
        '| SHIN-1 | Mapping check |',
      ].join('\n');
      const requests = convertMarkdownToRequests(markdown, 1);

      const tableReq = requests.find((r) => r.insertTable);
      assert.ok(tableReq);
      assert.strictEqual(tableReq.insertTable.rows, 2);
      assert.strictEqual(tableReq.insertTable.columns, 2);

      const insertedTexts = requests.filter((r) => r.insertText).map((r) => r.insertText.text);
      assert.ok(insertedTexts.includes('Task ID'));
      assert.ok(insertedTexts.includes('Task Name'));
      assert.ok(insertedTexts.includes('SHIN-1'));
      assert.ok(insertedTexts.includes('Mapping check'));

      const boldReqs = requests.filter((r) => r.updateTextStyle?.textStyle?.bold);
      assert.ok(boldReqs.length > 0);
    });
  });

  describe('Mixed Content', () => {
    it('should convert document with multiple elements', () => {
      const markdown = `# Title

This is **bold** and *italic* text with a [link](https://example.com).

- List item 1
- List item 2

## Heading 2

More content.`;

      const requests = convertMarkdownToRequests(markdown, 1);

      assert.strictEqual(requests.some((r) => r.insertText), true);
      assert.strictEqual(requests.some((r) => r.updateTextStyle), true);
      assert.strictEqual(requests.some((r) => r.updateParagraphStyle), true);
      assert.strictEqual(requests.some((r) => r.createParagraphBullets), true);

      assert.ok(
        requests.find((r) => r.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'HEADING_1')
      );

      assert.ok(
        requests.find((r) => r.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'HEADING_2')
      );
    });
  });

  describe('Index Tracking', () => {
    it('should use correct start index', () => {
      const requests = convertMarkdownToRequests('Test text', 100);

      const insertReq = requests.find((r) => r.insertText);
      assert.ok(insertReq);
      assert.strictEqual(insertReq.insertText.location.index, 100);
    });

    it('should track indices for sequential inserts', () => {
      const requests = convertMarkdownToRequests('First paragraph.\n\nSecond paragraph.', 1);

      const insertReqs = requests.filter((r) => r.insertText);
      assert.ok(insertReqs.length > 0);

      for (const req of insertReqs) {
        assert.ok(req.insertText.location);
        assert.strictEqual(typeof req.insertText.location.index, 'number');
      }
    });
  });

  describe('Tab Support', () => {
    it('should include tabId in requests when provided', () => {
      const requests = convertMarkdownToRequests('**bold text**', 1, 'tab123');

      const insertReq = requests.find((r) => r.insertText);
      assert.ok(insertReq);
      assert.strictEqual(insertReq.insertText.location.tabId, 'tab123');

      const styleReq = requests.find((r) => r.updateTextStyle);
      assert.ok(styleReq);
      assert.strictEqual(styleReq.updateTextStyle.range.tabId, 'tab123');
    });
  });

  describe('Paragraph Spacing', () => {
    it('should apply spaceBelow to normal text paragraphs', () => {
      const requests = convertMarkdownToRequests('First paragraph.\n\nSecond paragraph.', 1);

      const spacingReqs = requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType &&
          !r.updateParagraphStyle?.paragraphStyle?.borderBottom
      );
      assert.strictEqual(spacingReqs.length, 2);

      for (const req of spacingReqs) {
        assert.strictEqual(req.updateParagraphStyle.paragraphStyle.spaceBelow.magnitude, 8);
        assert.strictEqual(req.updateParagraphStyle.paragraphStyle.spaceBelow.unit, 'PT');
        assert.strictEqual(req.updateParagraphStyle.fields, 'spaceBelow');
      }
    });

    it('should only apply spaceBelow to the last item of a list, not every item', () => {
      const requests = convertMarkdownToRequests('- Item 1\n- Item 2\n- Item 3', 1);

      const spacingReqs = requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType &&
          !r.updateParagraphStyle?.paragraphStyle?.borderBottom
      );
      // Only 1 spacing request: the trailing spacing on the last list item
      assert.strictEqual(spacingReqs.length, 1);
    });

    it('should not apply spaceBelow to headings (they have named styles)', () => {
      const requests = convertMarkdownToRequests('# Heading\n\n## Subheading', 1);

      const spacingReqs = requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType &&
          !r.updateParagraphStyle?.paragraphStyle?.borderBottom
      );
      assert.strictEqual(spacingReqs.length, 0);
    });

    it('should apply spaceBelow to normal paragraphs and last list items in mixed content', () => {
      const markdown = '# Title\n\nA paragraph.\n\n- List item\n\nAnother paragraph.';
      const requests = convertMarkdownToRequests(markdown, 1);

      const spacingReqs = requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType &&
          !r.updateParagraphStyle?.paragraphStyle?.borderBottom
      );
      // "A paragraph." + "Another paragraph." + last list item trailing spacing = 3
      assert.strictEqual(spacingReqs.length, 3);
    });

    it('should include tabId in spacing requests when provided', () => {
      const requests = convertMarkdownToRequests('A paragraph.', 1, 'tab-xyz');

      const spacingReqs = requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType
      );
      assert.strictEqual(spacingReqs.length, 1);
      assert.strictEqual(spacingReqs[0].updateParagraphStyle.range.tabId, 'tab-xyz');
    });
  });

  describe('List Trailing Spacing', () => {
    // Helper to find spacing requests that target list items (not normal paragraphs or headings).
    // We identify them by checking they don't overlap with normalParagraph spacing ranges or
    // heading styles. Instead we just verify the total spaceBelow count vs paragraph-only count.
    function getListSpacingReqs(requests) {
      // All spaceBelow requests that are NOT heading styles and NOT border styles
      return requests.filter(
        (r) =>
          r.updateParagraphStyle?.paragraphStyle?.spaceBelow &&
          !r.updateParagraphStyle?.paragraphStyle?.namedStyleType &&
          !r.updateParagraphStyle?.paragraphStyle?.borderBottom
      );
    }

    it('should apply spaceBelow to the last item of a bullet list', () => {
      const requests = convertMarkdownToRequests('- Item 1\n- Item 2\n- Item 3', 1);

      const spacingReqs = getListSpacingReqs(requests);
      // 1 request for the last list item (no normal paragraphs here)
      assert.strictEqual(spacingReqs.length, 1);
      assert.strictEqual(spacingReqs[0].updateParagraphStyle.paragraphStyle.spaceBelow.magnitude, 8);
    });

    it('should apply spaceBelow to the last item of an ordered list', () => {
      const requests = convertMarkdownToRequests('1. First\n2. Second\n3. Third', 1);

      const spacingReqs = getListSpacingReqs(requests);
      assert.strictEqual(spacingReqs.length, 1);
      assert.strictEqual(spacingReqs[0].updateParagraphStyle.paragraphStyle.spaceBelow.magnitude, 8);
    });

    it('should apply spaceBelow after each separate list in the document', () => {
      const markdown = '- A\n- B\n\nSome text.\n\n1. One\n2. Two';
      const requests = convertMarkdownToRequests(markdown, 1);

      const spacingReqs = getListSpacingReqs(requests);
      // 2 list-trailing spacing + 1 normal paragraph spacing = 3 total
      assert.strictEqual(spacingReqs.length, 3);
    });

    it('should create spacing between a list and the following paragraph', () => {
      const markdown = '- Item 1\n- Item 2\n\nFollowing paragraph.';
      const requests = convertMarkdownToRequests(markdown, 1);

      const spacingReqs = getListSpacingReqs(requests);
      // 1 for last list item + 1 for the following paragraph = 2
      assert.strictEqual(spacingReqs.length, 2);
    });

    it('should handle nested lists and apply spacing after the top-level list', () => {
      const markdown = '- Parent\n  - Child 1\n  - Child 2\n\nAfter the list.';
      const requests = convertMarkdownToRequests(markdown, 1);

      const spacingReqs = getListSpacingReqs(requests);
      // 1 for last item of the top-level list + 1 for the following paragraph = 2
      assert.strictEqual(spacingReqs.length, 2);
    });

    it('should include tabId in list spacing requests when provided', () => {
      const requests = convertMarkdownToRequests('- Item 1\n- Item 2', 1, 'tab-list');

      const spacingReqs = getListSpacingReqs(requests);
      assert.strictEqual(spacingReqs.length, 1);
      assert.strictEqual(spacingReqs[0].updateParagraphStyle.range.tabId, 'tab-list');
    });
  });

  describe('Edge Cases', () => {
    it('should handle empty markdown', () => {
      assert.strictEqual(convertMarkdownToRequests('', 1).length, 0);
    });

    it('should handle whitespace-only markdown', () => {
      assert.strictEqual(convertMarkdownToRequests('   \n\n   ', 1).length, 0);
    });

    it('should handle plain text without formatting', () => {
      const requests = convertMarkdownToRequests('Just plain text', 1);

      const insertReq = requests.find((r) => r.insertText);
      assert.ok(insertReq);
      assert.strictEqual(insertReq.insertText.text, 'Just plain text');

      const styleReqs = requests.filter((r) => r.updateTextStyle);
      assert.strictEqual(styleReqs.length, 0);
    });
  });

  describe('Horizontal Rules', () => {
    it('should produce a border-bottom paragraph style for ---', () => {
      const requests = convertMarkdownToRequests('Above\n\n---\n\nBelow', 1);

      const hrReqs = requests.filter((r) => r.updateParagraphStyle?.paragraphStyle?.borderBottom);
      assert.strictEqual(hrReqs.length, 1);

      const border = hrReqs[0].updateParagraphStyle.paragraphStyle.borderBottom;
      assert.strictEqual(border.dashStyle, 'SOLID');
      assert.strictEqual(border.width.magnitude, 1);
      assert.strictEqual(border.width.unit, 'PT');
    });

    it('should handle multiple horizontal rules', () => {
      const requests = convertMarkdownToRequests(
        '# Title\n\n---\n\n## S1\n\nText.\n\n---\n\n## S2',
        1
      );

      const hrReqs = requests.filter((r) => r.updateParagraphStyle?.paragraphStyle?.borderBottom);
      assert.strictEqual(hrReqs.length, 2);
    });

    it('should not drop surrounding content', () => {
      const requests = convertMarkdownToRequests('Above\n\n---\n\nBelow', 1);

      const allText = requests
        .filter((r) => r.insertText)
        .map((r) => r.insertText.text)
        .join('');

      assert.ok(allText.includes('Above'));
      assert.ok(allText.includes('Below'));
    });

    it('should place the HR paragraph between surrounding content', () => {
      const requests = convertMarkdownToRequests('Above\n\n---\n\nBelow', 1);

      const hrReqs = requests.filter((r) => r.updateParagraphStyle?.paragraphStyle?.borderBottom);
      assert.strictEqual(hrReqs.length, 1);

      const hrStart = hrReqs[0].updateParagraphStyle.range.startIndex;
      const hrEnd = hrReqs[0].updateParagraphStyle.range.endIndex;

      const aboveInsert = requests.find(
        (r) => r.insertText && r.insertText.text.includes('Above')
      );
      const belowInsert = requests.find(
        (r) => r.insertText && r.insertText.text.includes('Below')
      );

      assert.ok(aboveInsert.insertText.location.index < hrStart);
      assert.ok(belowInsert.insertText.location.index >= hrEnd);
    });

    it('should include tabId on HR border requests when provided', () => {
      const requests = convertMarkdownToRequests('---', 1, 'tab-abc');

      const hrReqs = requests.filter((r) => r.updateParagraphStyle?.paragraphStyle?.borderBottom);
      assert.ok(hrReqs.length > 0);
      assert.strictEqual(hrReqs[0].updateParagraphStyle.range.tabId, 'tab-abc');
    });

    it('should work in a realistic document with headings, lists, and rules', () => {
      const markdown = `# Project Plan

---

## Goals

- **Speed:** Ship faster
- **Quality:** Fewer bugs

## Timeline

1. Planning
2. Execution
3. Review

---

*Last updated: 2026*`;

      const requests = convertMarkdownToRequests(markdown, 1);

      // HRs
      const hrReqs = requests.filter((r) => r.updateParagraphStyle?.paragraphStyle?.borderBottom);
      assert.strictEqual(hrReqs.length, 2);

      // Headings
      const h1Reqs = requests.filter(
        (r) => r.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'HEADING_1'
      );
      const h2Reqs = requests.filter(
        (r) => r.updateParagraphStyle?.paragraphStyle?.namedStyleType === 'HEADING_2'
      );
      assert.strictEqual(h1Reqs.length, 1);
      assert.strictEqual(h2Reqs.length, 2);

      // Bullet lists (merged into one range)
      const bulletReqs = requests.filter(
        (r) => r.createParagraphBullets?.bulletPreset === 'BULLET_DISC_CIRCLE_SQUARE'
      );
      assert.strictEqual(bulletReqs.length, 1);

      // Numbered list (merged into one range)
      const numberedReqs = requests.filter(
        (r) => r.createParagraphBullets?.bulletPreset === 'NUMBERED_DECIMAL_ALPHA_ROMAN'
      );
      assert.strictEqual(numberedReqs.length, 1);

      // Bold
      const boldReqs = requests.filter((r) => r.updateTextStyle?.textStyle?.bold === true);
      assert.ok(boldReqs.length >= 2);

      // Italic
      const italicReqs = requests.filter((r) => r.updateTextStyle?.textStyle?.italic === true);
      assert.ok(italicReqs.length >= 1);

      // All text present
      const allText = requests
        .filter((r) => r.insertText)
        .map((r) => r.insertText.text)
        .join('');
      assert.ok(allText.includes('Project Plan'));
      assert.ok(allText.includes('Ship faster'));
      assert.ok(allText.includes('Execution'));
      assert.ok(allText.includes('Last updated: 2026'));
    });
  });

  describe('Images omitted', () => {
    it('omits image tokens and does not insert the image URL', () => {
      const requests = convertMarkdownToRequests('![alt](https://example.com/img.png)', 1);
      assert.equal(requests.some((r) => r.insertInlineImage), false);
      assert.equal(JSON.stringify(requests).includes('https://example.com/img.png'), false);
    });
  });

  describe('List outside context', () => {
    it('throws MarkdownConversionError when a list item is outside a list', () => {
      assert.throws(
        () => convertMarkdownTokensToRequests([{ type: 'list_item_open' }]),
        (error) => {
          assert.ok(error instanceof MarkdownConversionError);
          assert.strictEqual(error.message, 'List item found outside of list context');
          return true;
        },
      );
    });
  });
});
