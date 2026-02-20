# Development Guide

**Generated**: 2026-02-19

## Prerequisites

- **Node.js** 18+ (includes npm)
- **Git** for version control
- **Google Cloud Project** with Docs, Sheets, Slides, Drive, and Gmail APIs enabled
- **OAuth credentials** (`credentials.json`) from Google Cloud Console

## Quick Start

```bash
# Clone and install
git clone <repo-url> && cd Google-Drive-MCP
npm install

# Place credentials.json in project root (from Google Cloud Console)

# Build TypeScript
npm run build

# First run — triggers OAuth browser flow
node ./dist/server.js
# Visit the URL printed in terminal, authorize, token.json is created

# Subsequent runs (MCP clients call this automatically)
node index.js
```

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `CREDENTIALS_PATH` | No | Custom path to credentials.json (default: project root) |
| `TOKEN_PATH` | No | Custom path to token.json (default: project root) |
| `GOOGLE_CLIENT_ID` | No | OAuth client ID (alternative to credentials.json) |
| `GOOGLE_CLIENT_SECRET` | No | OAuth client secret (alternative to credentials.json) |
| `GOOGLE_REDIRECT_URI` | No | Custom redirect URI (default: `http://localhost:3000/`) |
| `GOOGLE_REFRESH_TOKEN` | No | Refresh token (alternative to token.json) |
| `SERVICE_ACCOUNT_PATH` | No | Path to service account key file (skips OAuth) |
| `GOOGLE_IMPERSONATE_USER` | No | Email to impersonate with service account |

## Project Commands

| Command | Description |
|---------|-------------|
| `npm run build` | Compile TypeScript to `dist/` |
| `npm test` | Run tests with Node.js test runner |
| `node index.js` | Start the MCP server |

## Build Process

```bash
# TypeScript compilation
npm run build
# Equivalent to: tsc
# Compiles src/**/*.ts → dist/**/*.js
# Config: tsconfig.json (ES2022, NodeNext modules, strict)
```

## Testing

Tests use Node.js built-in test runner (`node --test`):

```bash
npm test
# Runs: node --test tests/
```

Test files:
- `tests/types.test.js` — Zod schema validation
- `tests/slides.test.js` — Slides helper functions
- `tests/helpers.test.js` — Docs/Sheets helpers

## Adding a New Tool

1. **Define parameter schema** in `src/types.ts` (if reusable) or inline in `server.ts`
2. **Add helper functions** in the appropriate `src/*Helpers.ts` module
3. **Register the tool** in `src/server.ts`:

```typescript
server.addTool({
  name: 'myNewTool',
  description: 'What this tool does',
  parameters: z.object({
    // Zod schema for parameters
  }),
  execute: async (args, { log }) => {
    const client = await getDocsClient(); // or getSheetsClient(), etc.
    // Implementation
    return 'result string';
  },
});
```

4. **Build and test**: `npm run build && npm test`

## MCP Client Configuration

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "google-docs": {
      "command": "node",
      "args": ["/absolute/path/to/Google-Drive-MCP/index.js"]
    }
  }
}
```

### VS Code

See `vscode.md` for VS Code MCP extension setup.

## File Organization Conventions

- **One helper module per Google API** — keeps related logic together
- **Zod schemas in types.ts** when reused across tools, inline when tool-specific
- **Error handling**: Use `UserError` from FastMCP for client-facing errors, plain `Error` for internal issues
- **Logging**: Use `log.info()` / `log.error()` from tool context, `console.error()` in helpers (stderr, not visible to MCP clients)

## Authentication Notes

- First run requires interactive browser OAuth flow
- `token.json` persists the refresh token — treat as sensitive
- Google may rotate refresh tokens; the server auto-detects and saves rotated tokens
- Service account mode (`SERVICE_ACCOUNT_PATH`) bypasses OAuth entirely
- After adding Gmail scope to an existing setup, delete `token.json` and re-authenticate
