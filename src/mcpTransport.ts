// src/mcpTransport.ts
import { UserError } from 'fastmcp';

function isLoopbackHost(host: string): boolean {
  return host === '127.0.0.1' || host === '::1' || host === 'localhost';
}

export function resolveHttpStreamBind(env: NodeJS.Dict<string>): {
  host: string;
  port: number;
  token: string | undefined;
} {
  const host = env.MCP_HOST || '127.0.0.1';
  const port = Number(env.MCP_PORT || 8787);
  const token = env.MCP_HTTP_TOKEN ? env.MCP_HTTP_TOKEN : undefined;

  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new UserError('MCP_PORT must be an integer between 1 and 65535.');
  }

  if (!isLoopbackHost(host) && !token) {
    throw new UserError(
      'MCP_HOST is not loopback; set MCP_HTTP_TOKEN to bind a non-loopback interface.',
    );
  }

  return { host, port, token };
}

export function createBearerAuthenticate(expectedToken: string) {
  return async (request?: { headers?: { authorization?: string | string[] } }) => {
    if (!request || !request.headers) {
      return { authenticated: true };
    }
    if (request.headers.authorization === `Bearer ${expectedToken}`) {
      return { authenticated: true };
    }
    throw new Response(null, { status: 401, statusText: 'Unauthorized' });
  };
}

export async function startConfiguredTransport(
  server: { start(config: unknown): Promise<void> },
  env: NodeJS.Dict<string>,
  log: (message: string) => void,
): Promise<void> {
  const useHttp = env.MCP_TRANSPORT === 'httpStream';
  if (useHttp) {
    const bind = resolveHttpStreamBind(env);
    const configToUse = {
      transportType: 'httpStream' as const,
      httpStream: { port: bind.port, host: bind.host },
    };
    await server.start(configToUse);
    log(`MCP Server running on httpStream transport, port ${bind.port}. Endpoint: /mcp`);
    return;
  }

  const configToUse = { transportType: 'stdio' as const };
  await server.start(configToUse);
  log('MCP Server running on stdio transport. Awaiting client connection...');
}
