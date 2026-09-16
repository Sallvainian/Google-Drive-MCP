// tests/gmail-headers.test.js
import {
  createSimpleEmail,
  createEmailWithAttachments,
  encodeEmailHeader,
  extractEmailAddresses,
  buildReplyRecipients,
  validateEmail,
} from '../dist/googleGmailApiHelpers.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const existingAttachment = fileURLToPath(import.meta.url);
const CRLF_MESSAGE = 'Header values must not contain CR or LF.';
const NO_RECIPIENT_MESSAGE = 'Cannot reply: no valid recipient address found in the original message.';

function assertCrlfUserError(error) {
  assert.ok(error instanceof UserError);
  assert.strictEqual(error.message, CRLF_MESSAGE);
  return true;
}

function assertCannotReplyUserError(error) {
  assert.ok(error instanceof UserError);
  assert.strictEqual(error.message, NO_RECIPIENT_MESSAGE);
  return true;
}

describe('gmail header sanitization and reply recipients', () => {
  describe('createSimpleEmail', () => {
    it('writes a happy-path Subject and a single To', () => {
      const mime = createSimpleEmail({
        to: ['alice@example.com'],
        subject: 'Report',
        body: 'Hi',
      });
      assert.ok(mime.includes('Subject: Report'));
      assert.ok(mime.includes('To: alice@example.com'));
      assert.strictEqual((mime.match(/^To:/gm) || []).length, 1);
    });

    it('encodes a non-ASCII subject as RFC 2047 without introducing CR/LF', () => {
      const mime = createSimpleEmail({
        to: ['alice@example.com'],
        subject: 'Café',
        body: 'Hi',
      });
      const encoded = Buffer.from('Café', 'utf-8').toString('base64');
      assert.ok(mime.includes(`Subject: =?UTF-8?B?${encoded}?=`));
      const subjectLine = mime.split('\r\n').find((line) => line.startsWith('Subject: '));
      assert.ok(subjectLine);
      assert.equal(subjectLine.includes('\r'), false);
      assert.equal(subjectLine.includes('\n'), false);
    });

    it('throws UserError on CR/LF in subject and does not return MIME', () => {
      let mime;
      assert.throws(() => {
        mime = createSimpleEmail({
          to: ['alice@example.com'],
          subject: 'Report\r\nBcc: exfil@attacker.com',
          body: 'Hi',
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on LF-only in subject and does not return MIME', () => {
      let mime;
      assert.throws(() => {
        mime = createSimpleEmail({
          to: ['alice@example.com'],
          subject: 'Report\nBcc: exfil@attacker.com',
          body: 'Hi',
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on CR-only in subject and does not return MIME', () => {
      let mime;
      assert.throws(() => {
        mime = createSimpleEmail({
          to: ['alice@example.com'],
          subject: 'Report\rBcc: exfil@attacker.com',
          body: 'Hi',
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on CR/LF in inReplyTo and does not return MIME', () => {
      let mime;
      assert.throws(() => {
        mime = createSimpleEmail({
          to: ['alice@example.com'],
          subject: 'Report',
          body: 'Hi',
          inReplyTo: '<id@x>\r\nBcc: exfil@attacker.com',
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on CR/LF in recipient and does not return MIME', () => {
      let mime;
      assert.throws(() => {
        mime = createSimpleEmail({
          to: ['alice@example.com\r\nBcc: exfil@attacker.com'],
          subject: 'Report',
          body: 'Hi',
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });
  });

  describe('createEmailWithAttachments', () => {
    it('throws UserError on CR/LF in subject with a real attachment file', async () => {
      let mime;
      await assert.rejects(async () => {
        mime = await createEmailWithAttachments({
          to: ['alice@example.com'],
          subject: 'Report\r\nBcc: exfil@attacker.com',
          body: 'Hi',
          attachments: [existingAttachment],
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on CR/LF in inReplyTo with a real attachment file', async () => {
      let mime;
      await assert.rejects(async () => {
        mime = await createEmailWithAttachments({
          to: ['alice@example.com'],
          subject: 'Report',
          body: 'Hi',
          inReplyTo: '<id@x>\r\nBcc: exfil@attacker.com',
          attachments: [existingAttachment],
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });

    it('throws UserError on CR/LF in recipient with a real attachment file', async () => {
      let mime;
      await assert.rejects(async () => {
        mime = await createEmailWithAttachments({
          to: ['alice@example.com\r\nBcc: exfil@attacker.com'],
          subject: 'Report',
          body: 'Hi',
          attachments: [existingAttachment],
        });
      }, assertCrlfUserError);
      assert.equal(mime, undefined);
    });
  });

  describe('encodeEmailHeader', () => {
    it('throws UserError on CR/LF in a filename value and does not write it', () => {
      let encoded;
      assert.throws(() => {
        encoded = encodeEmailHeader('file\r\nBcc: x.txt');
      }, assertCrlfUserError);
      assert.equal(encoded, undefined);
    });
  });

  describe('extractEmailAddresses', () => {
    it('extracts the addr-spec from a display-name From', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('Alice <alice@example.com>'),
        ['alice@example.com']
      );
    });

    it('pairs < with the next > not the last > in the mailbox', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('Alice <alice@example.com> [score>5]'),
        ['alice@example.com']
      );
    });

    it('drops addr-specs that fail validateEmail and keeps survivors', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('Foo <not-an-email>, bar@x.com'),
        ['bar@x.com']
      );
    });

    it('extracts only the addr-spec from a quoted comma display name', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('"Smith, Alice" <alice@example.com>'),
        ['alice@example.com']
      );
    });

    it('keeps every surviving addr-spec from a two-mailbox From', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('"Bob" <bob@x.com>, evil@attacker.com'),
        ['bob@x.com', 'evil@attacker.com']
      );
    });

    it('extracts a bare From addr-spec', () => {
      assert.deepStrictEqual(
        extractEmailAddresses('alice@example.com'),
        ['alice@example.com']
      );
    });

    it('returns an empty array and does not throw for missing or invalid mailboxes', () => {
      assert.deepStrictEqual(extractEmailAddresses(undefined), []);
      assert.deepStrictEqual(extractEmailAddresses(''), []);
      assert.deepStrictEqual(extractEmailAddresses('Not a mailbox'), []);
    });
  });

  describe('buildReplyRecipients', () => {
    it('throws UserError when From is missing', () => {
      assert.throws(
        () => buildReplyRecipients(undefined, undefined, false),
        assertCannotReplyUserError
      );
    });

    it('throws UserError when From is empty', () => {
      assert.throws(
        () => buildReplyRecipients('', undefined, false),
        assertCannotReplyUserError
      );
    });

    it('throws UserError when From is not a mailbox', () => {
      assert.throws(
        () => buildReplyRecipients('Not a mailbox', undefined, false),
        assertCannotReplyUserError
      );
    });

    it('returns the From addr-spec when replyAll is false', () => {
      assert.deepStrictEqual(
        buildReplyRecipients('Alice <alice@example.com>', undefined, false),
        ['alice@example.com']
      );
    });

    it('ignores the To header when replyAll is false', () => {
      assert.deepStrictEqual(
        buildReplyRecipients('Alice <alice@example.com>', 'bob@x.com', false),
        ['alice@example.com']
      );
    });

    it('concatenates From and extracted To when replyAll is true', () => {
      assert.deepStrictEqual(
        buildReplyRecipients('bob@x.com', '"Smith, Alice" <alice@example.com>', true),
        ['bob@x.com', 'alice@example.com']
      );
    });
  });

  describe('validateEmail', () => {
    it('accepts an addr-spec and rejects a display-name mailbox', () => {
      assert.strictEqual(validateEmail('alice@example.com'), true);
      assert.strictEqual(validateEmail('Alice <alice@example.com>'), false);
    });
  });

  describe('reply_to_email recipient assignment', () => {
    it('calls buildReplyRecipients and does not splice raw From or comma-split To', () => {
      const source = readFileSync(
        join(dirname(fileURLToPath(import.meta.url)), '..', 'src', 'server.ts'),
        'utf8'
      );
      const start = source.indexOf("name: 'reply_to_email'");
      const end = source.indexOf("name: 'forward_email'");
      assert.ok(start !== -1 && end > start);
      const block = source.slice(start, end);
      assert.ok(block.includes("buildReplyRecipients(headers['from'], headers['to'], args.replyAll)"));
      assert.equal(block.includes("[headers['from']"), false);
      assert.equal(block.includes("headers['to']?.split(',')"), false);
    });
  });
});
