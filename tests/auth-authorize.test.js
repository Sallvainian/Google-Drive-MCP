// tests/auth-authorize.test.js
//
// Drives the real authorize() so the state handed to generateAuthUrl is tied to the
// callback listener and to getToken.
//
// TOKEN_PATH and CREDENTIALS_PATH are bound BEFORE ../dist/auth.js is imported on
// purpose: auth.ts resolves them into module constants at import time, so a test that
// sets them afterwards leaves loadSavedCredentialsIfExist() reading the operator's real
// token.json. On a checkout that has one, authorize() would return through the saved
// credentials branch (never reaching the code under test), refreshAccessToken() would
// make a live Google call, and a rotated refresh token would be written into this
// throwaway directory instead of the operator's file.
import assert from 'node:assert';
import http from 'node:http';
import { createRequire } from 'node:module';
import { mkdtempSync } from 'node:fs';
import { rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { after, describe, it } from 'node:test';
import { OAuth2Client } from 'google-auth-library';

const require = createRequire(import.meta.url);
const childProcess = require('child_process');

const sandbox = mkdtempSync(path.join(tmpdir(), 'gd-mcp-authorize-'));
process.env.TOKEN_PATH = path.join(sandbox, 'token.json');
process.env.CREDENTIALS_PATH = path.join(sandbox, 'credentials.json');
process.env.GOOGLE_CLIENT_ID = 'test-client-id';
process.env.GOOGLE_CLIENT_SECRET = 'test-client-secret';
delete process.env.GOOGLE_REFRESH_TOKEN;
delete process.env.SERVICE_ACCOUNT_PATH;
delete process.env.REQUIRED_ACCOUNT_EMAIL;

const { authorize } = await import('../dist/auth.js');

// authenticate() hard-codes this port for the installed-app flow.
const CALLBACK_PORT = 3000;
// What our own listener answers a bare GET / with.
const LISTENER_PROBE_BODY = 'No authorization code received';

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

// Returns the hostname that reaches THIS flow's callback listener.
//
// "Does anything answer on the port?" is not a sufficient readiness probe. Port 3000 is a
// common dev-server port, authenticate() calls server.listen(port) without a host, and
// localhost may resolve to either family — so a foreign occupant can answer the probe and
// every assertion below would then run against that service instead of the OAuth
// listener. Identify the responder before trusting it.
async function waitForOwnListener(port, timeoutMs = 8000) {
  const start = Date.now();
  const hosts = ['127.0.0.1', '::1', 'localhost'];
  let foreign = null;
  for (;;) {
    for (const hostname of hosts) {
      try {
        const res = await httpGet(port, '/', hostname, 250);
        if (res.status === 400 && res.body.includes(LISTENER_PROBE_BODY)) {
          return hostname;
        }
        foreign = `${hostname} answered HTTP ${res.status}`;
      } catch {
        // nothing listening on this host yet
      }
    }
    if (Date.now() - start > timeoutMs) {
      throw new Error(
        `no OAuth callback listener on port ${port} within ${timeoutMs}ms` +
          (foreign ? ` — a different service is holding it (${foreign})` : '')
      );
    }
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
}

describe('authorize() OAuth wiring', () => {
  after(async () => {
    await rm(sandbox, { recursive: true, force: true });
  });

  it(
    'should carry the generateAuthUrl state into the callback and exchange only a matching code',
    { timeout: 20000 },
    async () => {
      let capturedState;
      const getTokenCodes = [];
      const execCommands = [];
      const origGenerateAuthUrl = OAuth2Client.prototype.generateAuthUrl;
      const origGetToken = OAuth2Client.prototype.getToken;
      const origExec = childProcess.exec;

      OAuth2Client.prototype.generateAuthUrl = function (opts) {
        capturedState = opts.state;
        return origGenerateAuthUrl.call(this, opts);
      };
      OAuth2Client.prototype.getToken = async function (code) {
        getTokenCodes.push(code);
        return {
          tokens: {
            refresh_token: 'test-refresh-token',
            access_token: 'test-access-token',
            expiry_date: Date.now() + 3600_000,
            token_type: 'Bearer',
            scope: 'https://www.googleapis.com/auth/drive',
          },
        };
      };
      childProcess.exec = (command, callback) => {
        execCommands.push(command);
        if (typeof callback === 'function') callback(null, '', '');
        return { unref() {}, ref() {}, kill() {} };
      };

      let authorizeSettled = false;
      const authorizePromise = authorize().finally(() => {
        authorizeSettled = true;
      });
      authorizePromise.catch(() => {});
      try {
        const hostname = await waitForOwnListener(CALLBACK_PORT);
        // Not just "some string": the whole point of state is that a third party cannot
        // predict it, so pin the shape randomBytes(32).toString('hex') produces. A constant
        // would satisfy a length check and every other assertion in this file.
        assert.match(
          capturedState ?? '',
          /^[0-9a-f]{64}$/,
          'generateAuthUrl must receive a 32-byte random per-flow state'
        );

        const bad = await httpGet(
          CALLBACK_PORT,
          '/?code=attacker-code&state=not-the-expected',
          hostname
        );
        assert.ok(
          bad.status >= 400 && bad.status < 500,
          `mismatched state must be rejected, got HTTP ${bad.status}`
        );
        assert.strictEqual(getTokenCodes.length, 0, 'a mismatched-state code must not be exchanged');

        const good = await httpGet(
          CALLBACK_PORT,
          `/?code=legit-code&state=${encodeURIComponent(capturedState)}`,
          hostname
        );
        assert.strictEqual(good.status, 200);
        await authorizePromise;
        assert.deepStrictEqual(getTokenCodes, ['legit-code']);

        // The browser hand-off hangs off the listener's `listening` promise; without this
        // nothing notices if that wiring is dropped and the operator is left staring at
        // "Opening browser for authentication..." with no browser.
        assert.strictEqual(execCommands.length, 1, 'authorize() must launch the browser once');
        assert.ok(
          execCommands[0].includes(capturedState),
          "browser command must carry this flow's authorize URL"
        );
      } finally {
        OAuth2Client.prototype.generateAuthUrl = origGenerateAuthUrl;
        OAuth2Client.prototype.getToken = origGetToken;
        childProcess.exec = origExec;
        if (!authorizeSettled) {
          await httpGet(CALLBACK_PORT, '/?error=access_denied', '127.0.0.1', 500).catch(() => {});
          await httpGet(CALLBACK_PORT, '/?error=access_denied', '::1', 500).catch(() => {});
          await authorizePromise.catch(() => {});
        }
      }
    }
  );
});
