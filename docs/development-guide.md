# Development Guide

**Generated:** 2026-05-05 | **Scan Level:** Exhaustive

---

## Prerequisites

- **Node.js 18+** (ES2022 compatible, ESM)
- **npm** (lockfile committed; `package-lock.json`)
- **Google Cloud Project** with these APIs enabled:
  - Google Docs API
  - Google Drive API
  - Google Sheets API
  - Google Slides API
  - Gmail API
- **OAuth credentials** (`credentials.json`) OR **Service Account key** (`service-account.json`)

## Quick Start

```bash
# Clone and install
git clone <repo-url>
cd Google-Drive-MCP
npm install

# Place credentials
# Option A: OAuth (interactive)
cp /path/to/credentials.json ./credentials.json

# Option B: Service Account (headless)
export SERVICE_ACCOUNT_PATH=/path/to/service-account.json

# Build
npm run build

# First run (triggers OAuth consent if using OAuth)
node ./dist/server.js
```

## Commands

| Command | Description |
|---------|-------------|
| `npm run build` | Compile TypeScript to `dist/` via `tsc` |
| `npm test` | Run tests via `node --test tests/` (requires prior `npm run build`) |
| `node ./dist/server.js` | Start the MCP server (stdio mode) |
| `node index.js` | Same as above (entry point wrapper) |

**Always build before testing:** Tests import from `dist/`, not `src/`.

## Environment Variables

### Authentication

| Variable | Required | Description |
|----------|----------|-------------|
| `GOOGLE_CLIENT_ID` | No* | OAuth client ID (alternative to credentials.json) |
| `GOOGLE_CLIENT_SECRET` | No* | OAuth client secret (alternative to credentials.json) |
| `GOOGLE_REFRESH_TOKEN` | No | Pre-set refresh token (skips interactive OAuth on first run; falls through to `token.json` if stale) |
| `GOOGLE_REDIRECT_URI` | No | OAuth redirect URI (default: `http://localhost:3000/`) |
| `SERVICE_ACCOUNT_PATH` | No | Path to service account JSON key file — when set, takes precedence over OAuth |
| `GOOGLE_IMPERSONATE_USER` | No | Email to impersonate via domain-wide delegation (Service Account only) |
| `REQUIRED_ACCOUNT_EMAIL` | No | Pin authentication to a specific Google account; throws if mismatch |

\* Either set `GOOGLE_CLIENT_ID` + `GOOGLE_CLIENT_SECRET` env vars, OR place `credentials.json` in the project root.

### Path Overrides

| Variable | Default | Description |
|----------|---------|-------------|
| `TOKEN_PATH` | `./token.json` | Custom path for OAuth token persistence |
| `CREDENTIALS_PATH` | `./credentials.json` | Custom path for OAuth credentials file |

### direnv Configuration

The project includes `.envrc` for automatic env var loading via `direnv`:
```bash
export GOOGLE_CLIENT_ID="..."
export GOOGLE_CLIENT_SECRET="..."
export GOOGLE_PROJECT_ID="..."
```

## MCP Client Configuration

### Claude Code / Claude Desktop

```json
{
  "mcpServers": {
    "google-drive-mcp": {
      "command": "node",
      "args": ["/path/to/Google-Drive-MCP/index.js"],
      "env": {
        "REQUIRED_ACCOUNT_EMAIL": "your-email@gmail.com"
      }
    }
  }
}
```

### Multi-account setup

Run separate instances pinned to different accounts:

```json
{
  "mcpServers": {
    "Drive-MCP-Personal": {
      "command": "node",
      "args": ["/path/to/Google-Drive-MCP/index.js"],
      "env": {
        "TOKEN_PATH": "/path/to/token-personal.json",
        "REQUIRED_ACCOUNT_EMAIL": "personal@gmail.com"
      }
    },
    "Drive-MCP-Work": {
      "command": "node",
      "args": ["/path/to/Google-Drive-MCP/index.js"],
      "env": {
        "TOKEN_PATH": "/path/to/token-work.json",
        "REQUIRED_ACCOUNT_EMAIL": "work@company.com"
      }
    }
  }
}
```

### VS Code (MCP Extension)

See `vscode.md` in the project root for VS Code-specific setup.

## Adding a New Tool

1. **Define a Zod schema** (if new parameters needed) in `src/types.ts`:
   ```typescript
   export const MyNewParameter = z.object({
     myField: z.string().describe('Description shown to the LLM'),
   });
   export type MyNewArgs = z.infer<typeof MyNewParameter>;
   ```

2. **Add a helper function** in the appropriate `src/google*ApiHelpers.ts`:
   ```typescript
   export async function myNewOperation(docs: Docs, ...args): Promise<Result> {
     // Implementation using Google API client
     // Wrap API errors with UserError for client-facing messages
   }
   ```

3. **Register the tool** in `src/server.ts`:
   ```typescript
   server.addTool({
     name: 'myNewTool',
     description: 'What this tool does',
     parameters: MyNewParameter,
     execute: async (args, { log }) => {
       const docs = await getDocsClient(); // lazy client getter
       const result = await GDocsHelpers.myNewOperation(docs, args.myField);
       return JSON.stringify(result); // tools return strings
     },
   });
   ```

4. **Build and test**:
   ```bash
   npm run build
   npm test
   ```

### Tool registration patterns

- Use `UserError` (from `fastmcp`) for client-facing errors (wrong document ID, permission denied, validation failures the LLM should see)
- Use plain `Error` for internal/fatal issues (init failures, programmer errors)
- Always call the appropriate `getXxxClient()` helper at the top of the handler — these throw `UserError` if the client failed to initialize
- Tools must return a `string` (JSON-serialized for complex data, plain text for simple results)
- For large responses, summarize and provide key data points rather than dumping raw API responses

### Mapping Google API errors

The codebase maps API error codes consistently:

| API code | Maps to |
|----------|---------|
| 404 | `UserError("... not found")` |
| 403 | `UserError("Permission denied ...")` |
| 400 | `UserError("Invalid request: ..." + parsed details)` |
| Other | `Error("Google API Error (code): message")` |

## Project Structure Conventions

- **Source**: All TypeScript source in `src/`
- **Output**: Compiled JS in `dist/` (gitignored)
- **Tests**: JavaScript test files in `tests/` (import from `dist/`, not `src/`)
- **Docs**: Generated documentation in `docs/`
- **Entry point**: `index.js` at root (thin wrapper that imports `dist/server.js`)
- **Module system**: ESM (`"type": "module"` in `package.json`)
- **Imports**: Use `.js` extensions in all relative imports — even when importing `.ts` files (NodeNext resolution)
- **Logging**: `console.error()` only — `console.log()` corrupts the MCP stdio protocol
- **No linter**: No ESLint or Prettier configured. Match surrounding code style when editing `server.ts`.

## Common Development Tasks

### Rebuilding after changes
```bash
npm run build
```

### Running a specific test file
```bash
node --test tests/helpers.test.js
```

### Testing with a real Google account
```bash
# Ensure credentials/token are set up, then start the server in stdio mode:
node ./dist/server.js
# Connect via an MCP client to exercise tools
```

### Debugging auth issues
```bash
# Re-authenticate from scratch
rm token.json
node ./dist/server.js   # Browser will open for OAuth consent
```

If you've added new OAuth scopes (e.g., adding Gmail to a Docs-only setup), the existing `token.json` doesn't have the new scopes — you must delete it and re-authenticate.

### Force a specific account during OAuth
Set `REQUIRED_ACCOUNT_EMAIL` — this seeds `login_hint` on the consent URL (pre-selects the right account in Google's picker) AND validates after consent (rejects wrong-account auth before persisting the token).

### Checking which account a token belongs to
```bash
# The server logs the authenticated email at startup:
node ./dist/server.js 2>&1 | head -10
# Look for: "Account verified: <email>"
```

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| `invalid_grant` on startup | Stale refresh token | Server falls through env var → file → interactive. Delete `token.json` if file is also stale. |
| "Account mismatch" thrown | Wrong account during consent OR wrong `REQUIRED_ACCOUNT_EMAIL` | Re-run consent and pick the required account in the picker |
| MCP client disconnects immediately | `console.log()` somewhere in the code | Search for `console.log`, replace with `console.error` |
| Port 3000 in use | Another process holds it | Stop the other process; OAuth flow needs port 3000 specifically |
| Tools return "client not initialized" | Auth failed at startup | Check stderr for the original auth error |
| `findElement`/`fixListFormatting` errors | These are stubs/experimental | Don't rely on them; `findElement` throws `NotImplementedError` |
| Drive `files.export` fails for large file | 10MB API limit | `downloadFile` falls back to web URL automatically; expect a URL string, not file content |

## Process Resilience

The server installs two process-level handlers:

- **`uncaughtException`**: exits cleanly on `EPIPE`/`EBADF` (broken stdio = parent died); other errors are logged but the server stays alive.
- **`unhandledRejection`**: rejection-storm detector — if ≥10 rejections occur within 60s, exits with code 1 so a supervisor can restart with fresh state. Otherwise logs with counter context (`[N/10 in 60s]`).

Don't add competing handlers; they'd interfere with this safety net.

## Style Notes

- `server.ts` uses inconsistent indentation (some sections 0-indent, some 4-space) — match the surrounding section when editing.
- Helpers define local type aliases for the API client: `type Docs = docs_v1.Docs`, `type Gmail = gmail_v1.Gmail`. Continue this pattern.
- Reusable Zod schemas live in `types.ts`. Tool-specific one-off schemas can be defined inline in `server.ts`.
- Schema names use PascalCase with descriptive suffix: `DocumentIdParameter`, `TextStyleParameters`, `ApplyTextStyleToolParameters`.
- Type aliases follow: `export type TextStyleArgs = z.infer<typeof TextStyleParameters>`.
