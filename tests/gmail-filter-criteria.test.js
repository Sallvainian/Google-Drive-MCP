// tests/gmail-filter-criteria.test.js
import { createFilter, getFilter, listFilters } from '../dist/gmailFilterManager.js';
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

  it('should keep size 0 and false booleans from createFilter response mapping', async () => {
    const gmail = {
      users: {
        settings: {
          filters: {
            create: mock.fn(async () => ({
              data: {
                id: 'f1',
                criteria: {
                  from: 'alice@example.com',
                  hasAttachment: false,
                  excludeChats: false,
                  size: 0,
                },
                action: { addLabelIds: ['STARRED'] },
              },
            })),
          },
        },
      },
    };

    const result = await createFilter(
      gmail,
      { from: 'alice@example.com', hasAttachment: false, excludeChats: false, size: 0 },
      { addLabelIds: ['STARRED'] }
    );

    assert.strictEqual(result.criteria.size, 0);
    assert.strictEqual(result.criteria.hasAttachment, false);
    assert.strictEqual(result.criteria.excludeChats, false);
  });
});

describe('filter criteria 0/false mapping on get and list', () => {
  const criteria = {
    from: 'alice@example.com',
    hasAttachment: false,
    excludeChats: false,
    size: 0,
  };

  it('should keep size 0 and false booleans from getFilter', async () => {
    const gmail = {
      users: {
        settings: {
          filters: {
            get: mock.fn(async () => ({
              data: {
                id: 'f1',
                criteria,
                action: { addLabelIds: ['STARRED'] },
              },
            })),
          },
        },
      },
    };

    const result = await getFilter(gmail, 'f1');
    assert.strictEqual(result.criteria.size, 0);
    assert.strictEqual(result.criteria.hasAttachment, false);
    assert.strictEqual(result.criteria.excludeChats, false);
  });

  it('should keep size 0 and false booleans from listFilters', async () => {
    const gmail = {
      users: {
        settings: {
          filters: {
            list: mock.fn(async () => ({
              data: {
                filter: [
                  {
                    id: 'f1',
                    criteria,
                    action: { addLabelIds: ['STARRED'] },
                  },
                ],
              },
            })),
          },
        },
      },
    };

    const result = await listFilters(gmail);
    assert.strictEqual(result.length, 1);
    assert.strictEqual(result[0].criteria.size, 0);
    assert.strictEqual(result[0].criteria.hasAttachment, false);
    assert.strictEqual(result[0].criteria.excludeChats, false);
  });
});

describe('createFilter positive matcher leftovers', () => {
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
