#!/usr/bin/env node
import { spawn } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, '..');
const serverPath = join(repoRoot, 'dist', 'server.js');
const registerPath = join(here, 'stub-authorize-register.mjs');
const outPath = process.argv[2];
const timeoutMs = Number(process.env.CAPTURE_TIMEOUT_MS || 120000);

function send(child, message) {
  child.stdin.write(`${JSON.stringify(message)}\n`);
}

function waitForMessage(state, predicate, label) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (fn, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      state.emitter.delete(onChange);
      fn(value);
    };
    const tryRead = () => {
      if (state.error) {
        finish(reject, state.error);
        return true;
      }
      while (state.queue.length > 0) {
        const message = state.queue.shift();
        if (predicate(message)) {
          finish(resolve, message);
          return true;
        }
      }
      return false;
    };
    const onChange = () => {
      tryRead();
    };
    const timer = setTimeout(() => {
      finish(reject, new Error(`timed out waiting for ${label}`));
    }, timeoutMs);
    state.emitter.add(onChange);
    tryRead();
  });
}

async function capture() {
  const env = { ...process.env };
  delete env.GOOGLE_REFRESH_TOKEN;
  delete env.SERVICE_ACCOUNT_PATH;
  env.TOKEN_PATH = join(here, '.no-token.json');
  env.CREDENTIALS_PATH = join(here, '.no-credentials.json');
  env.MCP_TRANSPORT = 'stdio';

  const child = spawn(
    process.execPath,
    ['--import', pathToFileURL(registerPath).href, serverPath],
    {
      cwd: repoRoot,
      env,
      stdio: ['pipe', 'pipe', 'pipe'],
    }
  );

  const stderrChunks = [];
  const state = {
    buffer: '',
    queue: [],
    emitter: new Set(),
    error: null,
  };

  const notify = () => {
    for (const listener of state.emitter) listener();
  };

  child.stderr.on('data', (chunk) => {
    stderrChunks.push(chunk);
    process.stderr.write(chunk);
  });

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

  let finished = false;
  const childExit = new Promise((resolve, reject) => {
    child.on('error', (err) => {
      if (!finished) reject(err);
    });
    child.on('exit', (code, signal) => {
      if (finished) return;
      const stderr = Buffer.concat(stderrChunks).toString('utf8');
      const err = new Error(
        `dist/server.js exited ${code ?? signal} before tools/list completed\n${stderr}`
      );
      state.error = err;
      notify();
      reject(err);
    });
  });

  try {
    send(child, {
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        capabilities: {},
        clientInfo: { name: 'token-cost-capture', version: '0.0.0' },
      },
    });

    const initialized = await Promise.race([
      waitForMessage(
        state,
        (msg) => msg.id === 1 && (msg.result || msg.error),
        'initialize result'
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

    const tools = [];
    let cursor;
    let listId = 2;
    do {
      const id = listId++;
      send(child, {
        jsonrpc: '2.0',
        id,
        method: 'tools/list',
        params: cursor ? { cursor } : {},
      });
      const response = await Promise.race([
        waitForMessage(
          state,
          (msg) => msg.id === id && (msg.result || msg.error),
          `tools/list id=${id}`
        ),
        childExit,
      ]);
      if (response.error) {
        throw new Error(`tools/list failed: ${JSON.stringify(response.error)}`);
      }
      tools.push(...response.result.tools);
      cursor = response.result.nextCursor;
    } while (cursor);

    finished = true;
    const payload = `${JSON.stringify({ tools }, null, 2)}\n`;
    if (outPath) {
      writeFileSync(outPath, payload);
    } else {
      process.stdout.write(payload);
    }
  } finally {
    finished = true;
    if (child.stdin) child.stdin.end();
    child.kill('SIGTERM');
    setTimeout(() => child.kill('SIGKILL'), 1000).unref();
  }
}

capture()
  .then(() => {
    process.exit(0);
  })
  .catch((err) => {
    console.error(err.message || err);
    process.exit(1);
  });
