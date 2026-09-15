// tests/auth.test.js
import { accountMatchesRequired, authorize } from '../dist/auth.js';
import assert from 'node:assert';
import { afterEach, describe, it } from 'node:test';

const originalRequiredAccountEmail = process.env.REQUIRED_ACCOUNT_EMAIL;

afterEach(() => {
  if (originalRequiredAccountEmail === undefined) {
    delete process.env.REQUIRED_ACCOUNT_EMAIL;
  } else {
    process.env.REQUIRED_ACCOUNT_EMAIL = originalRequiredAccountEmail;
  }
});

describe('accountMatchesRequired', () => {
  it('should export accountMatchesRequired as a function and keep authorize importable', () => {
    assert.strictEqual(typeof accountMatchesRequired, 'function');
    assert.strictEqual(typeof authorize, 'function');
  });

  it('should match required email case-insensitively with mixed casing on both operands', () => {
    process.env.REQUIRED_ACCOUNT_EMAIL = 'Work@Example.com';
    assert.strictEqual(accountMatchesRequired('work@example.com'), true);
    process.env.REQUIRED_ACCOUNT_EMAIL = 'work@example.com';
    assert.strictEqual(accountMatchesRequired('Work@Example.com'), true);
  });

  it('should fail closed on the (unknown) sentinel when REQUIRED_ACCOUNT_EMAIL is set', () => {
    process.env.REQUIRED_ACCOUNT_EMAIL = 'work@example.com';
    assert.strictEqual(accountMatchesRequired('(unknown)'), false);
  });

  it('should return true for any email when REQUIRED_ACCOUNT_EMAIL is unset', () => {
    delete process.env.REQUIRED_ACCOUNT_EMAIL;
    assert.strictEqual(accountMatchesRequired('anyone@example.com'), true);
  });
});
