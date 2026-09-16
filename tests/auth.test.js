// tests/auth.test.js
//
// TOKEN_PATH and CREDENTIALS_PATH are bound BEFORE ../dist/auth.js is imported, and the
// env snapshot is taken after that binding so afterEach restores to the sandbox rather
// than to the operator's environment. auth.ts resolves both into module constants at
// import time (src/auth.ts:17-18), so a file that sets them afterwards leaves
// saveCredentials() and loadSavedCredentialsIfExist() pointed at the real token.json.
import { OAuth2Client } from 'google-auth-library';
import assert from 'node:assert';
import http from 'node:http';
import { mkdtempSync } from 'node:fs';
import { readFile, rm, stat, unlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { after, afterEach, describe, it } from 'node:test';

const sandbox = mkdtempSync(path.join(tmpdir(), 'gd-mcp-auth-'));
const SANDBOX_TOKEN_PATH = path.join(sandbox, 'token.json');
process.env.TOKEN_PATH = SANDBOX_TOKEN_PATH;
process.env.CREDENTIALS_PATH = path.join(sandbox, 'credentials.json');
process.env.GOOGLE_CLIENT_ID = 'test-client-id';
process.env.GOOGLE_CLIENT_SECRET = 'test-client-secret';
delete process.env.GOOGLE_REFRESH_TOKEN;
delete process.env.SERVICE_ACCOUNT_PATH;

const { accountMatchesRequired, authorize, listenForOAuthCode, saveCredentials } = await import(
  '../dist/auth.js'
);

const originalEnv = {
  REQUIRED_ACCOUNT_EMAIL: process.env.REQUIRED_ACCOUNT_EMAIL,
  TOKEN_PATH: process.env.TOKEN_PATH,
  CREDENTIALS_PATH: process.env.CREDENTIALS_PATH,
  GOOGLE_CLIENT_ID: process.env.GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET,
  GOOGLE_REDIRECT_URI: process.env.GOOGLE_REDIRECT_URI,
  GOOGLE_REFRESH_TOKEN: process.env.GOOGLE_REFRESH_TOKEN,
  SERVICE_ACCOUNT_PATH: process.env.SERVICE_ACCOUNT_PATH,
};

const AUTH_TIMEOUT_MS = 5 * 60 * 1000;
let activeCallback = null;

function restoreEnvVar(key, original) {
  if (original === undefined) {
    delete process.env[key];
  } else {
    process.env[key] = original;
  }
}

async function isPending(promise) {
  const sentinel = Symbol('pending');
  return (await Promise.race([promise, Promise.resolve(sentinel)])) === sentinel;
}

function httpGet(port, requestPath, hostname = '127.0.0.1', timeoutMs = 3000) {
  return new Promise((resolve, reject) => {
    const req = http.get({ hostname, port, path: requestPath }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        clearTimeout(timer);
        resolve({
          status: res.statusCode,
          body: Buffer.concat(chunks).toString(),
        });
      });
    });
    const timer = setTimeout(() => {
      req.destroy();
      reject(new Error(`httpGet timeout ${hostname}:${port}${requestPath}`));
    }, timeoutMs);
    req.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
  });
}

function countActiveTimeoutsWithDelay(ms) {
  const probe = setTimeout(() => {}, ms);
  const seen = new Set();
  let count = 0;
  const visit = (node) => {
    while (node && !seen.has(node)) {
      seen.add(node);
      if (
        node !== probe &&
        node.constructor?.name === 'Timeout' &&
        node._idleTimeout === ms &&
        node._destroyed === false
      ) {
        count += 1;
      }
      node = node._idleNext;
    }
  };
  visit(probe._idleNext);
  visit(probe._idlePrev);
  clearTimeout(probe);
  return count;
}

function mockClient(refreshToken) {
  return {
    credentials: {
      refresh_token: refreshToken,
      access_token: 'ya29.access-token',
      expiry_date: Date.now() + 3600_000,
      token_type: 'Bearer',
      scope: 'https://www.googleapis.com/auth/drive',
    },
  };
}

afterEach(async () => {
  if (activeCallback) {
    activeCallback.close();
    await activeCallback.code.catch(() => {});
    activeCallback = null;
  }
  for (const key of Object.keys(originalEnv)) {
    restoreEnvVar(key, originalEnv[key]);
  }
  // Leave the sandbox token absent so each saveCredentials case exercises a fresh create.
  await unlink(SANDBOX_TOKEN_PATH).catch(() => {});
});

after(async () => {
  await rm(sandbox, { recursive: true, force: true });
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

describe('saveCredentials', () => {
  it('should create TOKEN_PATH with mode 0600 when the file does not yet exist', async () => {
    await assert.rejects(stat(SANDBOX_TOKEN_PATH), 'the token file must not exist yet');

    await saveCredentials(mockClient('refresh-token-for-mode-test'));
    const info = await stat(SANDBOX_TOKEN_PATH);
    assert.strictEqual(info.mode & 0o777, 0o600);
  });

  it('should print the env-var hint without the live refresh token', async () => {
    const refreshToken = 'ya29.distinctive-refresh-token-NOT-FOR-LOGS';
    const loggedArgs = [];
    const origError = console.error;
    console.error = (...args) => {
      loggedArgs.push(args);
    };
    try {
      await saveCredentials(mockClient(refreshToken));
    } finally {
      console.error = origError;
    }

    const combined = loggedArgs
      .flat()
      .map((arg) => String(arg))
      .join('\n');
    assert.ok(
      combined.includes('set GOOGLE_REFRESH_TOKEN to the refresh_token field in'),
      'stderr should include the env-var hint'
    );
    assert.ok(
      !/GOOGLE_REFRESH_TOKEN\s*=/.test(combined),
      'the hint must not print an assignment whose value is a placeholder'
    );
    for (const args of loggedArgs) {
      for (const arg of args) {
        assert.ok(
          !String(arg).includes(refreshToken),
          'no console.error argument may contain the live refresh token'
        );
      }
    }
  });

  it('should write the same payload keys an existing token.json uses', async () => {
    await saveCredentials(mockClient('refresh-token-existing-shape'));
    const parsed = JSON.parse(await readFile(SANDBOX_TOKEN_PATH, 'utf8'));
    assert.deepStrictEqual(Object.keys(parsed).sort(), [
      'access_token',
      'client_id',
      'client_secret',
      'expiry_date',
      'refresh_token',
      'scope',
      'token_type',
      'type',
    ]);
  });
});

describe('OAuth callback state and timeout', () => {
  it('should resolve with the code on a matching-state callback', async () => {
    const expectedState = 'expected-state-matching';
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    const res = await httpGet(port, `/?code=legit-code&state=${encodeURIComponent(expectedState)}`);
    assert.strictEqual(res.status, 200);
    assert.strictEqual(await waiter.code, 'legit-code');
    activeCallback = null;
  });

  it('should reject a mismatched-state callback without exchanging the code and still accept a later match', async () => {
    const expectedState = 'expected-state-mismatch';
    let getTokenCalls = 0;
    const origGetToken = OAuth2Client.prototype.getToken;
    OAuth2Client.prototype.getToken = async function mockedGetToken(...args) {
      getTokenCalls += 1;
      return origGetToken.apply(this, args);
    };
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    try {
      const port = await waiter.listening;
      const bad = await httpGet(port, '/?code=attacker-code&state=not-the-expected');
      assert.ok(bad.status >= 400 && bad.status < 500);
      assert.strictEqual(await isPending(waiter.code), true);
      assert.strictEqual(getTokenCalls, 0);

      const good = await httpGet(port, `/?code=legit-code&state=${encodeURIComponent(expectedState)}`);
      assert.strictEqual(good.status, 200);
      assert.strictEqual(await waiter.code, 'legit-code');
      assert.strictEqual(getTokenCalls, 0);
      activeCallback = null;
    } finally {
      OAuth2Client.prototype.getToken = origGetToken;
    }
  });

  it('should reject a missing-state callback without exchanging the code and still accept a later match', async () => {
    const expectedState = 'expected-state-missing';
    let getTokenCalls = 0;
    const origGetToken = OAuth2Client.prototype.getToken;
    OAuth2Client.prototype.getToken = async function mockedGetToken(...args) {
      getTokenCalls += 1;
      return origGetToken.apply(this, args);
    };
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    try {
      const port = await waiter.listening;
      const missing = await httpGet(port, '/?code=attacker-code');
      assert.ok(missing.status >= 400 && missing.status < 500);
      assert.strictEqual(await isPending(waiter.code), true);
      assert.strictEqual(getTokenCalls, 0);

      const good = await httpGet(port, `/?code=legit-code&state=${encodeURIComponent(expectedState)}`);
      assert.strictEqual(good.status, 200);
      assert.strictEqual(await waiter.code, 'legit-code');
      assert.strictEqual(getTokenCalls, 0);
      activeCallback = null;
    } finally {
      OAuth2Client.prototype.getToken = origGetToken;
    }
  });

  it('should warn on stderr when a code arrives with a mismatched state', async () => {
    // The mismatch branch is the only one that returns without settling, so this warning is
    // the operator's sole signal during the five minutes before the timer fires.
    const expectedState = 'expected-state-stderr-warning';
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    const port = await waiter.listening;

    const loggedArgs = [];
    const origError = console.error;
    console.error = (...args) => {
      loggedArgs.push(args);
    };
    let res;
    try {
      res = await httpGet(port, '/?code=attacker-code&state=not-the-expected');
    } finally {
      console.error = origError;
    }

    assert.ok(res.status >= 400 && res.status < 500);
    const combined = loggedArgs
      .flat()
      .map((arg) => String(arg))
      .join('\n');
    assert.match(combined, /state parameter did not match/i);
    assert.strictEqual(await isPending(waiter.code), true);
  });

  it('should drop the 5-minute auth timer from active resources after a matching-state callback', async () => {
    const expectedState = 'expected-state-timeout-cleared';
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    process.getActiveResourcesInfo();
    assert.ok(
      countActiveTimeoutsWithDelay(AUTH_TIMEOUT_MS) >= 1,
      '5-minute auth timer should be among active resources while waiting'
    );

    const res = await httpGet(port, `/?code=legit-code&state=${encodeURIComponent(expectedState)}`);
    assert.strictEqual(res.status, 200);
    assert.strictEqual(await waiter.code, 'legit-code');
    activeCallback = null;

    process.getActiveResourcesInfo();
    assert.strictEqual(
      countActiveTimeoutsWithDelay(AUTH_TIMEOUT_MS),
      0,
      '5-minute auth timer must not remain among the process active resources'
    );
  });

  it('should reject the code promise when the 5-minute timer actually fires', async (t) => {
    t.mock.timers.enable({ apis: ['setTimeout'] });
    try {
      const waiter = listenForOAuthCode('expected-state-timeout-fires', 0);
      activeCallback = waiter;
      await waiter.listening;
      const rejected = waiter.code.then(
        () => {
          throw new Error('expected waiter.code to reject when the timer fires');
        },
        (err) => err
      );
      t.mock.timers.tick(AUTH_TIMEOUT_MS);
      const err = await rejected;
      assert.match(err.message, /Authentication timed out after 5 minutes/);
      activeCallback = null;
    } finally {
      t.mock.timers.reset();
    }
  });

  it('should drop the 5-minute auth timer from active resources after error=access_denied', async () => {
    const waiter = listenForOAuthCode('expected-state-error-param', 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    const timerWasArmed = countActiveTimeoutsWithDelay(AUTH_TIMEOUT_MS) >= 1;
    const codeRejected = waiter.code.then(
      () => {
        throw new Error('expected waiter.code to reject');
      },
      (err) => err
    );
    const res = await httpGet(port, '/?error=access_denied');
    assert.ok(res.status >= 400 && res.status < 500);
    const err = await codeRejected;
    assert.match(err.message, /Authorization error: access_denied/);
    activeCallback = null;
    assert.ok(
      timerWasArmed,
      '5-minute auth timer should have been among active resources while waiting'
    );

    process.getActiveResourcesInfo();
    assert.strictEqual(
      countActiveTimeoutsWithDelay(AUTH_TIMEOUT_MS),
      0,
      '5-minute auth timer must not remain among the process active resources'
    );
  });

  it('should HTML-escape the untrusted error query in the failure page', async () => {
    const waiter = listenForOAuthCode('expected-state-html-escape', 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    const payload = '<script>alert(1)</script>&"\'';
    const codeRejected = waiter.code.then(
      () => {
        throw new Error('expected waiter.code to reject');
      },
      (err) => err
    );
    const res = await httpGet(port, `/?error=${encodeURIComponent(payload)}`);
    assert.ok(res.status >= 400 && res.status < 500);
    assert.equal(res.body.includes('<script>'), false);
    assert.ok(res.body.includes('&lt;script&gt;alert(1)&lt;/script&gt;&amp;&quot;&#39;'));
    const err = await codeRejected;
    assert.match(err.message, /Authorization error:/);
    assert.ok(err.message.includes(payload));
    activeCallback = null;
  });

  it('should ignore /favicon.ico even when it carries ?error=', async () => {
    const expectedState = 'expected-state-favicon';
    const waiter = listenForOAuthCode(expectedState, 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    const favicon = await httpGet(port, '/favicon.ico?error=access_denied');
    assert.strictEqual(favicon.status, 404);
    assert.strictEqual(await isPending(waiter.code), true);
    const good = await httpGet(port, `/?code=legit-code&state=${encodeURIComponent(expectedState)}`);
    assert.strictEqual(good.status, 200);
    assert.strictEqual(await waiter.code, 'legit-code');
    activeCallback = null;
  });

  it('should not settle the flow for ?error= on a non-callback path', async () => {
    const expectedState = 'expected-state-callback-path';
    const waiter = listenForOAuthCode(expectedState, 0, { callbackPath: '/oauth2callback' });
    activeCallback = waiter;
    const port = await waiter.listening;
    const stray = await httpGet(port, '/?error=access_denied');
    assert.strictEqual(stray.status, 404);
    assert.strictEqual(await isPending(waiter.code), true);
    const favicon = await httpGet(port, '/favicon.ico?error=access_denied');
    assert.strictEqual(favicon.status, 404);
    assert.strictEqual(await isPending(waiter.code), true);
    const good = await httpGet(
      port,
      `/oauth2callback?code=legit-code&state=${encodeURIComponent(expectedState)}`
    );
    assert.strictEqual(good.status, 200);
    assert.strictEqual(await waiter.code, 'legit-code');
    activeCallback = null;
  });

  it('should bind 127.0.0.1 so ::1 does not receive the callback', async () => {
    const waiter = listenForOAuthCode('expected-state-loopback-bind', 0);
    activeCallback = waiter;
    const port = await waiter.listening;
    await assert.rejects(
      () => httpGet(port, '/', '::1', 500),
      (err) => {
        assert.ok(err instanceof Error);
        assert.match(err.message, /ECONNREFUSED|httpGet timeout/);
        return true;
      }
    );
    const res = await httpGet(port, '/', '127.0.0.1');
    assert.strictEqual(res.status, 400);
    assert.ok(res.body.includes('No authorization code received'));
    assert.strictEqual(await isPending(waiter.code), true);
  });
});

describe('authorize() with an existing token.json', () => {
  it('should parse and use the saved token and leave an unrotated payload untouched', async () => {
    // The "Existing token still loads" matrix row. It also pins that the read path resolves
    // to the pre-import TOKEN_PATH binding: if it did not, this would hit the real token.json
    // and make a live Google call.
    const saved = {
      type: 'authorized_user',
      client_id: 'test-client-id',
      client_secret: 'test-client-secret',
      refresh_token: 'saved-refresh-token',
      access_token: 'saved-access-token',
      expiry_date: Date.now() + 3600_000,
      token_type: 'Bearer',
      scope: 'https://www.googleapis.com/auth/drive',
    };
    await writeFile(SANDBOX_TOKEN_PATH, JSON.stringify(saved), { mode: 0o600 });

    let refreshCalls = 0;
    const origRefresh = OAuth2Client.prototype.refreshAccessToken;
    OAuth2Client.prototype.refreshAccessToken = async function mockedRefresh() {
      refreshCalls += 1;
      // Same refresh_token back, so the rotation branch must not rewrite the file.
      return { credentials: { ...saved } };
    };
    try {
      const client = await authorize();
      assert.strictEqual(refreshCalls, 1, 'authorize() must verify the saved token');
      assert.strictEqual(client.credentials.refresh_token, 'saved-refresh-token');
      const onDisk = JSON.parse(await readFile(SANDBOX_TOKEN_PATH, 'utf8'));
      assert.deepStrictEqual(onDisk, saved, 'an unrotated token.json must be left as it was');
    } finally {
      OAuth2Client.prototype.refreshAccessToken = origRefresh;
    }
  });
});
