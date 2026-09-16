// tests/live-fastmcp-helpers.js
//
// Spawn dist/server.js with the CI stub-authorize loader and complete
// initialize + tools/list over stdio NDJSON or httpStream POST /mcp.
import assert from 'node:assert';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');
const scriptsDir = join(repoRoot, 'scripts');
const serverPath = join(repoRoot, 'dist', 'server.js');
const registerPath = join(scriptsDir, 'stub-authorize-register.mjs');

const DEFAULT_TIMEOUT_MS = Number(process.env.CAPTURE_TIMEOUT_MS || 120000);

export const CREATE_FROM_TEMPLATE_REPLACEMENTS_DESCRIPTION =
  'Key-value pairs for text replacements in the template (e.g., {"{{NAME}}": "John Doe", "{{DATE}}": "2024-01-01"}).';

// Zod 3 JSON Schema omits propertyNames; Zod 4 / FastMCP nested zod 4 emits
// propertyNames: { type: 'string' }. Both mean string-to-string.
export function assertCreateFromTemplateReplacementsSchema(schema) {
  assert.equal(schema.description, CREATE_FROM_TEMPLATE_REPLACEMENTS_DESCRIPTION);
  assert.ok(Array.isArray(schema.anyOf));
  const objectBranch = schema.anyOf.find((branch) => branch.type === 'object');
  const nullBranch = schema.anyOf.find((branch) => branch.type === 'null');
  assert.notEqual(objectBranch, undefined);
  assert.deepStrictEqual(nullBranch, { type: 'null' });
  assert.deepStrictEqual(objectBranch.additionalProperties, { type: 'string' });
  assert.equal(objectBranch.type, 'object');
  if (objectBranch.propertyNames !== undefined) {
    assert.deepStrictEqual(objectBranch.propertyNames, { type: 'string' });
  }
  assert.deepStrictEqual(
    Object.keys(objectBranch).toSorted(),
    objectBranch.propertyNames === undefined
      ? ['additionalProperties', 'type']
      : ['additionalProperties', 'propertyNames', 'type'],
  );
}

export function createStubAuthorizeEnv(overrides = {}) {
  const env = { ...process.env, ...overrides };
  delete env.GOOGLE_REFRESH_TOKEN;
  delete env.SERVICE_ACCOUNT_PATH;
  delete env.MCP_HTTP_TOKEN;
  env.TOKEN_PATH = join(scriptsDir, '.no-token.json');
  env.CREDENTIALS_PATH = join(scriptsDir, '.no-credentials.json');
  return env;
}

export function spawnLiveServer(env) {
  return spawn(
    process.execPath,
    ['--import', pathToFileURL(registerPath).href, serverPath],
    {
      cwd: repoRoot,
      env,
      stdio: ['pipe', 'pipe', 'pipe'],
    },
  );
}

export function killLiveServer(child) {
  return new Promise((resolve) => {
    let settled = false;
    let killTimer;
    const finish = () => {
      if (settled) return;
      settled = true;
      clearTimeout(killTimer);
      resolve();
    };
    child.once('exit', finish);
    if (child.exitCode !== null || child.signalCode !== null) {
      finish();
      return;
    }
    if (child.stdin) child.stdin.end();
    child.kill('SIGTERM');
    killTimer = setTimeout(() => child.kill('SIGKILL'), 1000);
  });
}

export function allocateLoopbackPort(host = '127.0.0.1') {
  return new Promise((resolve, reject) => {
    const server = createServer();
    server.listen(0, host, () => {
      const address = server.address();
      const port = typeof address === 'object' && address ? address.port : 0;
      server.close((err) => {
        if (err) {
          reject(err);
          return;
        }
        resolve(port);
      });
    });
    server.on('error', reject);
  });
}

function send(child, message) {
  child.stdin.write(`${JSON.stringify(message)}\n`);
}

function waitForChange(state, tryRead, label, timeoutMs) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (fn, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      state.emitter.delete(onChange);
      fn(value);
    };
    const onChange = () => {
      tryRead(finish, resolve, reject);
    };
    const timer = setTimeout(() => {
      finish(
        reject,
        new Error(`timed out waiting for ${label}\n${state.stderr}`),
      );
    }, timeoutMs);
    state.emitter.add(onChange);
    tryRead(finish, resolve, reject);
  });
}

function waitForMessage(state, predicate, label, timeoutMs) {
  return waitForChange(
    state,
    (finish, resolve, reject) => {
      if (state.error) {
        finish(reject, state.error);
        return;
      }
      while (state.queue.length > 0) {
        const message = state.queue.shift();
        if (predicate(message)) {
          finish(resolve, message);
          return;
        }
      }
    },
    label,
    timeoutMs,
  );
}

function waitForStderrIncludes(state, snippet, label, timeoutMs) {
  return waitForChange(
    state,
    (finish, resolve, reject) => {
      if (state.error) {
        finish(reject, state.error);
        return;
      }
      if (state.stderr.includes(snippet)) {
        finish(resolve, state.stderr);
      }
    },
    label,
    timeoutMs,
  );
}

function attachChild(child, { parseStdoutJsonRpc }) {
  const stderrChunks = [];
  const state = {
    buffer: '',
    queue: [],
    emitter: new Set(),
    error: null,
    stderr: '',
  };

  const notify = () => {
    for (const listener of state.emitter) listener();
  };

  child.stderr.on('data', (chunk) => {
    stderrChunks.push(chunk);
    state.stderr = Buffer.concat(stderrChunks).toString('utf8');
    notify();
  });

  if (parseStdoutJsonRpc) {
    child.stdout.on('data', (chunk) => {
      state.buffer += chunk.toString('utf8');
      let newline;
      while ((newline = state.buffer.indexOf('\n')) !== -1) {
        const line = state.buffer.slice(0, newline).replace(/\r$/, '');
        state.buffer = state.buffer.slice(newline + 1);
        if (!line.trim()) continue;
        try {
          state.queue.push(JSON.parse(line));
        } catch (err) {
          state.error = new Error(`invalid JSON-RPC line: ${line}\n${err}`);
        }
        notify();
      }
    });
  } else {
    child.stdout.on('data', () => {});
  }

  return { state, stderrChunks };
}

function makeChildExit(child, stderrChunks, isFinished, state) {
  return new Promise((resolve, reject) => {
    child.on('error', (err) => {
      if (!isFinished()) reject(err);
    });
    child.on('exit', (code, signal) => {
      if (isFinished()) return;
      const stderr = Buffer.concat(stderrChunks).toString('utf8');
      const err = new Error(
        `dist/server.js exited ${code ?? signal} before tools/list completed\n${stderr}`,
      );
      state.error = err;
      for (const listener of state.emitter) listener();
      reject(err);
    });
  });
}

async function collectToolsList(requestPage, childExit) {
  const tools = [];
  const listResults = [];
  let cursor;
  let listId = 2;
  do {
    const id = listId++;
    const response = await Promise.race([requestPage(id, cursor), childExit]);
    if (response.error) {
      throw new Error(`tools/list failed: ${JSON.stringify(response.error)}`);
    }
    if (!response.result || !Array.isArray(response.result.tools)) {
      throw new Error(`invalid JSON-RPC tools/list result: ${JSON.stringify(response)}`);
    }
    listResults.push(response);
    tools.push(...response.result.tools);
    cursor = response.result.nextCursor;
  } while (cursor);
  return { tools, listResults };
}

const INIT_PARAMS = {
  protocolVersion: '2024-11-05',
  capabilities: {},
  clientInfo: { name: 'live-fastmcp-suite', version: '0.0.0' },
};

async function handshakeStdio(child, state, childExit, timeoutMs) {
  send(child, {
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: INIT_PARAMS,
  });

  const initialized = await Promise.race([
    waitForMessage(
      state,
      (msg) => msg.id === 1 && (msg.result || msg.error),
      'initialize result',
      timeoutMs,
    ),
    childExit,
  ]);
  if (initialized.error) {
    throw new Error(`initialize failed: ${JSON.stringify(initialized.error)}`);
  }

  send(child, {
    jsonrpc: '2.0',
    method: 'notifications/initialized',
  });

  const listed = await collectToolsList(async (id, cursor) => {
    send(child, {
      jsonrpc: '2.0',
      id,
      method: 'tools/list',
      params: cursor ? { cursor } : {},
    });
    return waitForMessage(
      state,
      (msg) => msg.id === id && (msg.result || msg.error),
      `tools/list id=${id}`,
      timeoutMs,
    );
  }, childExit);

  return { initialize: initialized, ...listed };
}

function extractSseDataBlocks(buffer) {
  const blocks = [];
  let rest = buffer;
  for (;;) {
    const lf = rest.indexOf('\n\n');
    const crlf = rest.indexOf('\r\n\r\n');
    let sep = -1;
    let sepLen = 0;
    if (lf !== -1 && (crlf === -1 || lf < crlf)) {
      sep = lf;
      sepLen = 2;
    } else if (crlf !== -1) {
      sep = crlf;
      sepLen = 4;
    }
    if (sep === -1) break;
    const raw = rest.slice(0, sep);
    rest = rest.slice(sep + sepLen);
    const datas = [];
    for (const line of raw.split(/\r?\n/)) {
      if (line.startsWith('data:')) {
        datas.push(line.slice(5).trimStart());
      }
    }
    if (datas.length) blocks.push(datas.join('\n'));
  }
  return { blocks, rest };
}

async function readJsonRpcFromHttpResponse(response, id, timeoutMs) {
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    const body = await response.json();
    if (body && body.id === id && (body.result || body.error)) {
      return body;
    }
    throw new Error(`invalid JSON-RPC JSON body: ${JSON.stringify(body)}`);
  }
  if (!contentType.includes('text/event-stream')) {
    const text = await response.text().catch(() => '');
    throw new Error(
      `unexpected Content-Type ${contentType || '(none)'} from POST /mcp: ${text}`,
    );
  }
  if (!response.body) {
    throw new Error('SSE response has no body');
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';
  const deadline = Date.now() + timeoutMs;
  const parseBlocks = (blocks) => {
    for (const data of blocks) {
      let message;
      try {
        message = JSON.parse(data);
      } catch (err) {
        throw new Error(`invalid JSON-RPC SSE data: ${data}\n${err}`);
      }
      if (message.id === id && (message.result || message.error)) {
        return message;
      }
    }
    return null;
  };
  try {
    for (;;) {
      if (Date.now() > deadline) {
        throw new Error(`timed out waiting for SSE JSON-RPC id=${id}`);
      }
      const { value, done } = await reader.read();
      if (done) {
        buffer += decoder.decode();
        if (buffer.trim()) {
          const suffix = buffer.endsWith('\n\n') || buffer.endsWith('\r\n\r\n') ? '' : '\n\n';
          const { blocks } = extractSseDataBlocks(buffer + suffix);
          const parsed = parseBlocks(blocks);
          if (parsed) return parsed;
        }
        break;
      }
      buffer += decoder.decode(value, { stream: true });
      const { blocks, rest } = extractSseDataBlocks(buffer);
      buffer = rest;
      const parsed = parseBlocks(blocks);
      if (parsed) return parsed;
    }
  } finally {
    try {
      await reader.cancel();
    } catch {
      // already closed
    }
  }
  throw new Error(`SSE stream ended before JSON-RPC id=${id}`);
}

async function postMcp(url, message, sessionId, timeoutMs) {
  const headers = {
    Accept: 'application/json, text/event-stream',
    'Content-Type': 'application/json',
  };
  if (sessionId) {
    headers['Mcp-Session-Id'] = sessionId;
  }

  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(message),
      signal: ac.signal,
    });
    if (response.status < 200 || response.status >= 300) {
      const text = await response.text().catch(() => '');
      throw new Error(`non-2xx ${response.status} from POST /mcp: ${text}`);
    }
    const nextSession = response.headers.get('mcp-session-id') || sessionId;
    if (message.id !== undefined && !nextSession) {
      throw new Error('missing session');
    }
    if (message.id === undefined) {
      if (response.body) {
        await response.body.cancel();
      }
      return { sessionId: nextSession, message: null, status: response.status };
    }
    const rpc = await readJsonRpcFromHttpResponse(response, message.id, timeoutMs);
    return { sessionId: nextSession, message: rpc, status: response.status };
  } catch (err) {
    if (err && err.name === 'AbortError') {
      throw new Error(`handshake timeout POSTing ${message.method} to ${url}`);
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

async function handshakeHttp(url, timeoutMs, childExit) {
  const initPosted = postMcp(
    url,
    {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: INIT_PARAMS,
    },
    undefined,
    timeoutMs,
  );
  const initRes = await Promise.race([initPosted, childExit]);
  if (!initRes.sessionId) {
    throw new Error('missing session');
  }
  if (initRes.message.error) {
    throw new Error(`initialize failed: ${JSON.stringify(initRes.message.error)}`);
  }

  await Promise.race([
    postMcp(
      url,
      { jsonrpc: '2.0', method: 'notifications/initialized' },
      initRes.sessionId,
      timeoutMs,
    ),
    childExit,
  ]);

  const listed = await collectToolsList(
    async (id, cursor) => {
      const res = await postMcp(
        url,
        {
          jsonrpc: '2.0',
          id,
          method: 'tools/list',
          params: cursor ? { cursor } : {},
        },
        initRes.sessionId,
        timeoutMs,
      );
      return res.message;
    },
    childExit,
  );

  return {
    initialize: initRes.message,
    sessionId: initRes.sessionId,
    ...listed,
  };
}

export async function listToolsOverStdio({ timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const env = createStubAuthorizeEnv({ MCP_TRANSPORT: 'stdio' });
  const child = spawnLiveServer(env);
  const { state, stderrChunks } = attachChild(child, { parseStdoutJsonRpc: true });
  let finished = false;
  const childExit = makeChildExit(child, stderrChunks, () => finished, state);
  childExit.catch(() => {});
  try {
    const listed = await handshakeStdio(child, state, childExit, timeoutMs);
    return {
      ...listed,
      stderr: state.stderr,
    };
  } finally {
    finished = true;
    await killLiveServer(child);
  }
}

export async function listToolsOverHttpStream({
  timeoutMs = DEFAULT_TIMEOUT_MS,
  host = '127.0.0.1',
} = {}) {
  const port = await allocateLoopbackPort(host);
  const env = createStubAuthorizeEnv({
    MCP_TRANSPORT: 'httpStream',
    MCP_HOST: host,
    MCP_PORT: String(port),
  });
  const child = spawnLiveServer(env);
  const { state, stderrChunks } = attachChild(child, { parseStdoutJsonRpc: false });
  let finished = false;
  const childExit = makeChildExit(child, stderrChunks, () => finished, state);
  childExit.catch(() => {});
  const banner = `MCP Server running on httpStream transport, port ${port}. Endpoint: /mcp`;
  try {
    await Promise.race([
      waitForStderrIncludes(state, banner, 'httpStream listen banner', timeoutMs),
      childExit,
    ]);
    const url = `http://${host}:${port}/mcp`;
    const listed = await handshakeHttp(url, timeoutMs, childExit);
    return {
      ...listed,
      stderr: state.stderr,
      port,
      host,
      banner,
    };
  } finally {
    finished = true;
    await killLiveServer(child);
  }
}

export async function runStartServerCatch({ timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const env = createStubAuthorizeEnv({
    MCP_TRANSPORT: 'httpStream',
    MCP_PORT: 'abc',
  });
  const child = spawnLiveServer(env);
  const stderrChunks = [];
  child.stderr.on('data', (chunk) => stderrChunks.push(chunk));
  child.stdout.on('data', () => {});
  try {
    return await new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`startServer catch child did not exit within ${timeoutMs}ms`));
      }, timeoutMs);
      child.on('error', (err) => {
        clearTimeout(timer);
        reject(err);
      });
      child.on('exit', (code, signal) => {
        clearTimeout(timer);
        resolve({
          code,
          signal,
          stderr: Buffer.concat(stderrChunks).toString('utf8'),
        });
      });
    });
  } finally {
    await killLiveServer(child);
  }
}
