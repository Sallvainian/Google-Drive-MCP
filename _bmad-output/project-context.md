---
project_name: 'Google-Drive-MCP'
user_name: 'Sallvain'
date: '2026-02-19'
sections_completed: ['technology_stack', 'language_rules', 'framework_rules', 'testing_rules', 'quality_rules', 'workflow_rules', 'critical_rules']
status: 'complete'
rule_count: 38
optimized_for_llm: true
---

# Project Context for AI Agents

_Critical rules and patterns for implementing code in this project. Focus on unobvious details that agents might otherwise miss._

---

## Technology Stack & Versions

| Technology | Version | Notes |
|---|---|---|
| Node.js | 18+ | ESM (`"type": "module"`), built-in test runner |
| TypeScript | 5.8.3 | `strict: true`, ES2022 target, NodeNext modules |
| FastMCP | ^3.24.0 | MCP server framework |
| googleapis | ^148.0.0 | Docs v1, Sheets v4, Slides v1, Drive v3, Gmail v1 |
| google-auth-library | ^9.15.1 | OAuth2 + Service Account JWT |
| zod | ^3.24.2 | Runtime parameter validation |
| Package manager | npm | Lock file: package-lock.json |
| Build | `tsc` | `src/` → `dist/`, entry: `index.js` → `dist/server.js` |
| Test | `node --test tests/` | Node built-in runner, JS test files import from `dist/` |

---

## Critical Implementation Rules

### Language-Specific Rules (TypeScript + ESM)

- **Always use `.js` extensions in imports** — even when importing `.ts` files. NodeNext module resolution requires it: `import { authorize } from './auth.js'` not `'./auth'`
- **ESM only** — no `require()`, no `module.exports`. Use `import`/`export`. The project has `"type": "module"` in package.json
- **`strict: true`** is enforced — no implicit `any`, strict null checks, etc.
- **Use `unknown` over `any`** — when type is uncertain, prefer `unknown` and narrow. The codebase uses `error: any` in catch blocks for Google API errors (legacy pattern), but new code should use `unknown`
- **Path resolution in ESM** — use `fileURLToPath(import.meta.url)` and `path.dirname()` instead of `__dirname`/`__filename` (see `auth.ts` for the pattern)
- **`console.error()` for all logging** — MCP protocol uses stdout for communication. ALL server logging MUST go to stderr via `console.error()`. Never use `console.log()`

### Framework-Specific Rules (FastMCP + Google APIs)

- **Tool definition pattern** — all tools use `server.addTool()` with Zod schemas:
  ```typescript
  server.addTool({
    name: 'toolName',
    description: 'What it does',
    parameters: z.object({ /* Zod schema */ }),
    execute: async (args, { log }) => {
      const client = await getDocsClient(); // lazy init
      // implementation
      return 'string result';
    },
  });
  ```
- **Lazy client initialization** — Google API clients initialize on first tool call, not at server startup. Use getter functions: `getDocsClient()`, `getDriveClient()`, `getSheetsClient()`, `getSlidesClient()`, `getGmailClient()`
- **Error handling: `UserError` vs `Error`** — use `UserError` (from FastMCP) for client-facing errors the LLM should see. Use plain `Error` for internal/fatal issues. Map Google API error codes: 404→"not found", 403→"permission denied", 400→parse details
- **Zod schema organization** — reusable schemas go in `src/types.ts` (e.g., `DocumentIdParameter`, `TextStyleParameters`). Tool-specific one-off schemas are defined inline in `server.ts`
- **Helper module pattern** — one helper module per Google API. Helpers accept the typed API client as first parameter. All batch operations go through wrapper functions (e.g., `executeBatchUpdate`)
- **Tool return values** — tools must return a `string`. Format complex data as readable text or JSON strings
- **Google Docs indices are 1-based** — document content indices start at 1, not 0. Slides element positions are in points (72 points = 1 inch)

### Testing Rules

- **Test runner:** Node.js built-in (`node --test`). NOT Jest, NOT Vitest
- **Test files are JavaScript** (`.test.js`), NOT TypeScript. They import from compiled `dist/` directory
- **Import pattern in tests:** `import { fn } from '../dist/types.js'` — always from `dist/`, always with `.js` extension
- **Must build before testing:** `npm run build && npm test`. Tests run against compiled output
- **Test structure:** Use `describe()` and `it()` from `node:test`, `assert` from `node:assert`
- **No test framework dependencies** — no jest, mocha, chai, vitest in devDependencies
- **Test location:** `tests/` directory at project root (not `__tests__/`, not `src/**/*.test.ts`)

### Code Quality & Style Rules

- **File naming:** camelCase for all source files (`googleDocsApiHelpers.ts`, `gmailLabelManager.ts`)
- **No ESLint/Prettier configured** — no project-level linting. Follow existing code style
- **Inconsistent indentation** — `server.ts` and helpers use mixed indentation (some sections 0-indent, some 4-space). Match surrounding code when editing a section
- **Type aliases for API clients** — each helper defines a local alias: `type Docs = docs_v1.Docs`, `type Gmail = gmail_v1.Gmail`
- **Constants at module top** — batch limits, regex patterns, etc. defined as module-level constants
- **Zod schema naming** — PascalCase with descriptive suffix: `DocumentIdParameter`, `TextStyleParameters`, `ApplyTextStyleToolParameters`
- **Export types alongside schemas** — `export type TextStyleArgs = z.infer<typeof TextStyleParameters>`

### Development Workflow Rules

- **Build before run:** Always `npm run build` before testing or running the server
- **Entry point chain:** `index.js` → `dist/server.js` (compiled from `src/server.ts`)
- **Credentials never committed** — `credentials.json`, `token.json`, `.env*` are gitignored
- **After adding new OAuth scopes:** delete `token.json` and re-authenticate (applies when adding new Google API access)
- **server.ts is massive (211KB, ~5300 lines)** — all 99 tool definitions live here. When adding tools, add to the appropriate section (tools are grouped by API: Docs, Sheets, Slides, Drive, Gmail)

### Critical Don't-Miss Rules

- **NEVER use `console.log()`** — stdout is the MCP protocol transport. Any `console.log()` will corrupt the protocol stream and crash the client connection. Use `console.error()` or `log.info()`/`log.error()` from the tool context
- **NEVER commit `credentials.json` or `token.json`** — contains OAuth secrets. These are gitignored
- **Google API batch update ordering** — for Docs, requests in a batch are applied in reverse index order. When making multiple changes, process from end of document to beginning, or indices will shift
- **Token rotation handling** — Google may rotate refresh tokens at any time. The `installTokenRefreshListener()` in `auth.ts` handles this. Don't bypass or duplicate this logic
- **`findElement` is NOT IMPLEMENTED** — tool exists but throws `NotImplementedError`. Don't try to use or extend it without full implementation
- **`fixListFormatting` is EXPERIMENTAL** — may not work reliably. Document limitations when using
- **Gmail first-use gotcha** — after adding Gmail scope to an existing OAuth setup, existing `token.json` must be deleted and user must re-authenticate to get the new scope
- **Process-level error handlers exist** — `uncaughtException` and `unhandledRejection` handlers prevent server crashes. Don't add competing handlers
- **Dual auth support** — `authorize()` returns `OAuth2Client | JWT`. The auth method is determined by `SERVICE_ACCOUNT_PATH` env var presence. Both types work as `auth` parameter for Google API clients

---

## Usage Guidelines

**For AI Agents:**

- Read this file before implementing any code
- Follow ALL rules exactly as documented
- When in doubt, prefer the more restrictive option
- Update this file if new patterns emerge

**For Humans:**

- Keep this file lean and focused on agent needs
- Update when technology stack changes
- Review quarterly for outdated rules
- Remove rules that become obvious over time

Last Updated: 2026-02-19
