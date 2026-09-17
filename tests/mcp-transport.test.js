// tests/mcp-transport.test.js
import {
  createBearerAuthenticate,
  resolveHttpStreamBind,
  startConfiguredTransport,
} from '../dist/mcpTransport.js';
import {
  assertCreateFromTemplateReplacementsSchema,
  listToolsOverHttpStream,
  listToolsOverStdio,
  runStartServerCatch,
} from './live-fastmcp-helpers.js';
import { UserError } from 'fastmcp';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath } from 'node:url';

const srcDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');

function mockStartServer() {
  const calls = [];
  return {
    calls,
    server: {
      start: async (config) => {
        calls.push(config);
      },
    },
  };
}

describe('resolveHttpStreamBind', () => {
  it('Default loopback', () => {
    assert.deepStrictEqual(resolveHttpStreamBind({}), {
      host: '127.0.0.1',
      port: 8787,
      token: undefined,
    });
  });

  it('Custom port', () => {
    const bind = resolveHttpStreamBind({ MCP_PORT: '5555' });
    assert.strictEqual(bind.host, '127.0.0.1');
    assert.strictEqual(bind.port, 5555);
  });

  it('localhost', () => {
    const bind = resolveHttpStreamBind({ MCP_HOST: 'localhost' });
    assert.strictEqual(bind.host, 'localhost');
    assert.strictEqual(bind.token, undefined);
  });

  it('IPv6 loopback', () => {
    assert.strictEqual(resolveHttpStreamBind({ MCP_HOST: '::1' }).host, '::1');
  });

  it('Empty port', () => {
    assert.strictEqual(resolveHttpStreamBind({ MCP_PORT: '' }).port, 8787);
  });

  it('Non-loopback with token', () => {
    assert.throws(
      () =>
        resolveHttpStreamBind({
          MCP_HOST: '0.0.0.0',
          MCP_HTTP_TOKEN: 'secret',
        }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_HOST must be a loopback address (127.0.0.1, ::1, or localhost).',
        );
        return true;
      },
    );
  });

  it('Non-loopback without token', () => {
    assert.throws(
      () => resolveHttpStreamBind({ MCP_HOST: '0.0.0.0' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_HOST must be a loopback address (127.0.0.1, ::1, or localhost).',
        );
        return true;
      },
    );
  });

  it('Non-numeric port', () => {
    assert.throws(
      () => resolveHttpStreamBind({ MCP_PORT: 'abc' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_PORT must be an integer between 1 and 65535.',
        );
        return true;
      },
    );
  });

  it('Port 0', () => {
    assert.throws(
      () => resolveHttpStreamBind({ MCP_PORT: '0' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_PORT must be an integer between 1 and 65535.',
        );
        return true;
      },
    );
  });

  it('Port 65536', () => {
    assert.throws(
      () => resolveHttpStreamBind({ MCP_PORT: '65536' }),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_PORT must be an integer between 1 and 65535.',
        );
        return true;
      },
    );
  });
});

describe('createBearerAuthenticate', () => {
  const authenticate = createBearerAuthenticate('secret');

  it('Bearer match', async () => {
    const result = await authenticate({
      headers: { authorization: 'Bearer secret' },
    });
    assert.ok(result);
    assert.strictEqual(result.authenticated, true);
  });

  it('Bearer missing', async () => {
    await assert.rejects(
      () => authenticate({ headers: {} }),
      (error) => {
        assert.ok(error instanceof Response);
        assert.strictEqual(error.status, 401);
        return true;
      },
    );
  });

  it('Bearer wrong', async () => {
    await assert.rejects(
      () => authenticate({ headers: { authorization: 'Bearer other' } }),
      (error) => {
        assert.ok(error instanceof Response);
        assert.strictEqual(error.status, 401);
        return true;
      },
    );
  });

  it('stdio authenticate(undefined) stays authenticated', async () => {
    assert.deepStrictEqual(await authenticate(undefined), { authenticated: true });
  });
});

describe('startConfiguredTransport', () => {
  it('Start success http', { timeout: 180000 }, async () => {
    const listed = await listToolsOverHttpStream();
    const names = listed.tools.map((tool) => tool.name);
    assert.strictEqual(names.length, 111);
    assert.strictEqual(names[0], 'readGoogleDoc');
    assert.strictEqual(names[names.length - 1], 'send_draft');
    assert.deepStrictEqual(Object.keys(listed.listResults[0].result), ['tools']);
    const create = listed.tools.find((tool) => tool.name === 'createFromTemplate');
    assert.notEqual(create, undefined);
    assertCreateFromTemplateReplacementsSchema(create.inputSchema.properties.replacements);
    assert.ok(
      listed.stderr.includes(
        `MCP Server running on httpStream transport, port ${listed.port}. Endpoint: /mcp`,
      ),
    );
  });

  it('Start failure', async () => {
    const listenError = new Error('listen EADDRINUSE');
    const calls = [];
    const logs = [];
    const server = {
      start: async (config) => {
        calls.push(config);
        throw listenError;
      },
    };
    await assert.rejects(
      () =>
        startConfiguredTransport(server, { MCP_TRANSPORT: 'httpStream' }, (message) => {
          logs.push(message);
        }),
      (error) => error === listenError,
    );
    assert.strictEqual(calls.length, 1);
    assert.deepStrictEqual(logs, []);
  });

  it('Start success stdio', { timeout: 180000 }, async () => {
    const listed = await listToolsOverStdio();
    const names = listed.tools.map((tool) => tool.name);
    assert.strictEqual(names.length, 111);
    assert.strictEqual(names[0], 'readGoogleDoc');
    assert.strictEqual(names[names.length - 1], 'send_draft');
    assert.deepStrictEqual(Object.keys(listed.listResults[0].result), ['tools']);
    const create = listed.tools.find((tool) => tool.name === 'createFromTemplate');
    assert.notEqual(create, undefined);
    assertCreateFromTemplateReplacementsSchema(create.inputSchema.properties.replacements);
  });

  it('Start failure stdio', async () => {
    const listenError = new Error('listen EADDRINUSE');
    const calls = [];
    const logs = [];
    const server = {
      start: async (config) => {
        calls.push(config);
        throw listenError;
      },
    };
    await assert.rejects(
      () =>
        startConfiguredTransport(server, {}, (message) => {
          logs.push(message);
        }),
      (error) => error === listenError,
    );
    assert.strictEqual(calls.length, 1);
    assert.deepStrictEqual(logs, []);
  });

  it('httpStream custom port reaches start', async () => {
    const { calls, server } = mockStartServer();
    const logs = [];
    await startConfiguredTransport(
      server,
      { MCP_TRANSPORT: 'httpStream', MCP_PORT: '5555' },
      (message) => {
        logs.push(message);
      },
    );
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].httpStream.port, 5555);
    assert.deepStrictEqual(logs, [
      'MCP Server running on httpStream transport, port 5555. Endpoint: /mcp',
    ]);
  });

  it('httpStream localhost reaches start', async () => {
    const { calls, server } = mockStartServer();
    await startConfiguredTransport(
      server,
      { MCP_TRANSPORT: 'httpStream', MCP_HOST: 'localhost' },
      () => {},
    );
    assert.strictEqual(calls.length, 1);
    assert.strictEqual(calls[0].httpStream.host, 'localhost');
  });

  it('httpStream non-loopback with token does not start', async () => {
    const { calls, server } = mockStartServer();
    await assert.rejects(
      () =>
        startConfiguredTransport(
          server,
          {
            MCP_TRANSPORT: 'httpStream',
            MCP_HOST: '0.0.0.0',
            MCP_HTTP_TOKEN: 'secret',
          },
          () => {},
        ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_HOST must be a loopback address (127.0.0.1, ::1, or localhost).',
        );
        assert.strictEqual(calls.length, 0);
        return true;
      },
    );
  });

  it('httpStream non-loopback without token does not start', async () => {
    const { calls, server } = mockStartServer();
    await assert.rejects(
      () =>
        startConfiguredTransport(
          server,
          { MCP_TRANSPORT: 'httpStream', MCP_HOST: '0.0.0.0' },
          () => {},
        ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(
          error.message,
          'MCP_HOST must be a loopback address (127.0.0.1, ::1, or localhost).',
        );
        assert.strictEqual(calls.length, 0);
        return true;
      },
    );
  });

  it('httpStream non-numeric port does not start', async () => {
    const { calls, server } = mockStartServer();
    await assert.rejects(
      () =>
        startConfiguredTransport(
          server,
          { MCP_TRANSPORT: 'httpStream', MCP_PORT: 'abc' },
          () => {},
        ),
      (error) => {
        assert.ok(error instanceof UserError);
        assert.strictEqual(calls.length, 0);
        return true;
      },
    );
  });
});

describe('server.ts process handlers', () => {
  it('keeps a single unhandledRejection and uncaughtException handler', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.strictEqual(
      (source.match(/process\.on\('unhandledRejection'/g) || []).length,
      1,
    );
    assert.strictEqual(
      (source.match(/process\.on\('uncaughtException'/g) || []).length,
      1,
    );
  });

  it('awaits startConfiguredTransport and spreads bearer authenticate', () => {
    const source = readFileSync(join(srcDir, 'server.ts'), 'utf8');
    assert.ok(source.includes('await McpTransport.startConfiguredTransport'));
    assert.equal(source.includes('server.start('), false);
    assert.match(
      source,
      /\.\.\.\(process\.env\.MCP_HTTP_TOKEN\s*\?\s*\{\s*authenticate:\s*McpTransport\.createBearerAuthenticate\(process\.env\.MCP_HTTP_TOKEN\)\s*\}\s*:\s*\{\}\)/,
    );
  });

  it('startServer catch exits 1 with FATAL on invalid MCP_PORT', { timeout: 180000 }, async () => {
    const result = await runStartServerCatch();
    assert.strictEqual(result.code, 1);
    assert.ok(
      result.stderr.includes('FATAL: Server failed to start:'),
      result.stderr,
    );
    assert.ok(
      result.stderr.includes('MCP_PORT must be an integer between 1 and 65535.'),
      result.stderr,
    );
  });
});
