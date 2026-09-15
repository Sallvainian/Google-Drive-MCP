// tests/gmail-filter-criteria.test.js
import { createFilter } from '../dist/gmailFilterManager.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import { describe, it, mock } from 'node:test';

function stubGmail() {
  return {
    users: {
      settings: {
        filters: {
          create: mock.fn(async () => ({
            data: {
              id: 'f1',
              criteria: { from: 'alice@example.com' },
              action: { addLabelIds: ['STARRED'] },
            },
          })),
        },
      },
    },
  };
}

describe('createFilter positive matcher', () => {
  it('should create a filter with from', async () => {
    const gmail = stubGmail();
    const result = await createFilter(
      gmail,
      { from: 'alice@example.com' },
      { addLabelIds: ['STARRED'] }
    );

    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 1);
    assert.strictEqual(result.id, 'f1');
  });

  it('should create a filter with hasAttachment true', async () => {
    const gmail = stubGmail();
    await createFilter(
      gmail,
      { hasAttachment: true },
      { addLabelIds: ['STARRED'] }
    );

    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 1);
  });

  it('should create a filter with size 0', async () => {
    const gmail = stubGmail();
    await createFilter(
      gmail,
      { size: 0 },
      { addLabelIds: ['STARRED'] }
    );

    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 1);
  });

  it('should create a filter with a matcher plus excludeChats false', async () => {
    const gmail = stubGmail();
    await createFilter(
      gmail,
      { from: 'alice@example.com', excludeChats: false },
      { removeLabelIds: ['INBOX'] }
    );

    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 1);
  });

  it('should throw UserError for hasAttachment false plus archive', async () => {
    const gmail = stubGmail();

    await assert.rejects(
      async () => await createFilter(
        gmail,
        { hasAttachment: false },
        { removeLabelIds: ['INBOX'] }
      ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'At least one filter criteria must be specified.');
        return true;
      }
    );
    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 0);
  });

  it('should throw UserError for excludeChats false plus archive', async () => {
    const gmail = stubGmail();

    await assert.rejects(
      async () => await createFilter(
        gmail,
        { excludeChats: false },
        { removeLabelIds: ['INBOX'] }
      ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'At least one filter criteria must be specified.');
        return true;
      }
    );
    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 0);
  });

  it('should throw UserError for empty string from', async () => {
    const gmail = stubGmail();

    await assert.rejects(
      async () => await createFilter(
        gmail,
        { from: '' },
        { addLabelIds: ['STARRED'] }
      ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'At least one filter criteria must be specified.');
        return true;
      }
    );
    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 0);
  });

  it('should throw UserError for negatedQuery only', async () => {
    const gmail = stubGmail();

    await assert.rejects(
      async () => await createFilter(
        gmail,
        { negatedQuery: 'unsubscribe' },
        { addLabelIds: ['STARRED'] }
      ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'At least one filter criteria must be specified.');
        return true;
      }
    );
    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 0);
  });

  it('should throw UserError for excludeChats true only', async () => {
    const gmail = stubGmail();

    await assert.rejects(
      async () => await createFilter(
        gmail,
        { excludeChats: true },
        { addLabelIds: ['STARRED'] }
      ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(error.message, 'At least one filter criteria must be specified.');
        return true;
      }
    );
    assert.strictEqual(gmail.users.settings.filters.create.mock.calls.length, 0);
  });
});
