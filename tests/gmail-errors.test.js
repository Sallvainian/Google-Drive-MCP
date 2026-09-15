// tests/gmail-errors.test.js
import { listLabels, createLabel, getLabel } from '../dist/gmailLabelManager.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

describe('Gmail label error classification', () => {
  describe('listLabels', () => {
    it('should throw UserError for 500 response', async () => {
      const mockGmail = {
        users: {
          labels: {
            list: mock.fn(async () => {
              throw { code: 500, message: 'backend' };
            }),
          },
        },
      };

      await assert.rejects(
        async () => await listLabels(mockGmail),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, 'Gmail API Error listing labels: backend');
          return true;
        }
      );
    });
  });

  describe('createLabel', () => {
    it('should throw UserError when create returns no id', async () => {
      const mockGmail = {
        users: {
          labels: {
            list: mock.fn(async () => ({ data: { labels: [] } })),
            create: mock.fn(async () => ({ data: {} })),
          },
        },
      };

      await assert.rejects(
        async () => await createLabel(mockGmail, { name: 'Inbox-Triage' }),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, 'Invalid label data received from create operation.');
          return true;
        }
      );
    });
  });

  describe('getLabel', () => {
    it('should throw UserError for 500 response (ternary arm)', async () => {
      const mockGmail = {
        users: {
          labels: {
            get: mock.fn(async () => {
              throw { code: 500, message: 'backend' };
            }),
          },
        },
      };

      await assert.rejects(
        async () => await getLabel(mockGmail, 'Label_1'),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, 'Gmail API Error: backend');
          return true;
        }
      );
    });

    it('should throw UserError for 404 response', async () => {
      const mockGmail = {
        users: {
          labels: {
            get: mock.fn(async () => {
              throw { code: 404, message: 'Not found' };
            }),
          },
        },
      };

      await assert.rejects(
        async () => await getLabel(mockGmail, 'Label_1'),
        (error) => {
          assert.ok(error instanceof UserError);
          assert.strictEqual(error.message, 'Label not found (ID: Label_1).');
          return true;
        }
      );
    });
  });
});
