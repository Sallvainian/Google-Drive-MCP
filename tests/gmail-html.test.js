// tests/gmail-html.test.js
import { htmlToText } from '../dist/googleGmailApiHelpers.js';
import assert from 'node:assert';
import { describe, it } from 'node:test';

describe('htmlToText', () => {
  it('keeps paragraph text and an http(s) link as label (url)', () => {
    const out = htmlToText('<p>See <a href="https://example.com/order">your order</a>.</p>');
    assert.ok(out.includes('your order (https://example.com/order)'));
    assert.equal(out.includes('<p>'), false);
  });

  it('strips nested script wrappers that a single replace would reintroduce', () => {
    const out = htmlToText(
      '<scrip<script>is removed</script>t>alert(123)</script>visible'
    );
    assert.equal(out.toLowerCase().includes('<script'), false);
    assert.ok(out.includes('visible'));
  });

  it('does not resurrect tags from HTML entities', () => {
    const out = htmlToText('safe&lt;script&gt;alert(1)&lt;/script&gt;text');
    assert.equal(out.toLowerCase().includes('<script'), false);
    assert.ok(out.includes('safe'));
    assert.ok(out.includes('text'));
  });

  it('decodes a less-than entity into plaintext', () => {
    assert.equal(htmlToText('a &lt; b'), 'a < b');
  });
});
