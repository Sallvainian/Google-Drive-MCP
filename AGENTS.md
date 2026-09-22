<!-- bmad:context -->
<!-- Verified 2026-09-17 against 11d9158. Managed by bmad-project-context; edits inside
     this block are replaced on refresh. Keep anything you want preserved outside the markers. -->

## Google-Drive-MCP

FastMCP server over five Google Workspace APIs (Docs, Sheets, Slides, Drive, Gmail).
TypeScript ESM on npm, no framework beyond FastMCP. Generated reference docs live
in `docs/` — they are regenerated wholesale, not hand-edited.

## Policy

- Never commit `credentials.json` or `token.json` — OAuth secrets, gitignored at `.gitignore:6-7`.
- `.claude/skills/` is gitignored (vendor BMAD install). `.claude/` config, hooks,
  and personalities stay tracked. `.grok/`, `.codex/`, `.cursor/` and `_bmad/` stay
  ignored. `_bmad-output/` is gitignored in full; bmad-loop copies it into
  worktrees via `scm.worktree_seed`. Keep it local: specs index unfixed
  findings by file and line, and this repo is public.

## Where things are

- All tools register in `src/server.ts` via `server.addTool`, grouped by API (Docs,
  Sheets, Slides, Drive, Gmail) plus opt-in `sheets-advanced` and `docs-chips`.
  Grouping is `currentToolGroup` around `addTool` (`src/toolGroups.ts`). Add a new
  tool in its API section; do not introduce a `src/tools/` tree.
- Helper modules: `googleDocsApiHelpers.ts`, `googleSheetsApiHelpers.ts`,
  `googleSlidesApiHelpers.ts`, `googleDriveApiHelpers.ts`, `googleGmailApiHelpers.ts`,
  plus `gmailLabelManager.ts`, `gmailFilterManager.ts`, `markdown-transformer/`,
  and `mcpTransport.ts`. Helpers that call an API take the typed client as first parameter.
- Auth, token handling and account pinning: `src/auth.ts`.
- Reusable Zod schemas: `src/types.ts`; one-off tool schemas stay inline in `server.ts`.

## Running and verifying

- Tests import from `dist/`, which is gitignored. `npm test` runs a `pretest` hook
  (`tsc`) first, so a fresh clone compiles `dist/` before the suite imports it.
  `npm run build` is still `tsc` for a standalone compile.
- `npm test` is `node --test tests/*.test.js`. Invoking that glob directly skips
  pretest, so `dist/` must already exist. Do not run `node --test tests/` (the
  directory) — Node 24 resolves it as a module and fails with MODULE_NOT_FOUND.
- No linter or formatter is configured. Do not add one uninvited; there is no `lint` script.
- `tsconfig.json` covers `src/**/*` only — `tests/` is never typechecked.
- CI: `.github/workflows/ci.yml` on `pull_request` runs `npm ci`, `npm run build`,
  `npm test` (`npm test` runs `tsc` again via pretest), then the token-cost gate
  (`scripts/capture-tools-list.mjs` as the live ListTools input,
  `scripts/measure-token-cost.py` plus `scripts/check-token-budget.js`) under
  Node 24. The two Claude workflows still only invoke
  `anthropics/claude-code-action@v1` after `npm ci`; they are not the gate.
- `package.json` declares `"engines": { "node": ">=22" }`, matching README 22+. There is
  no `.nvmrc`. ci.yml, claude.yml, and claude-code-review.yml all pin Node 24.
- Locally, `python3 scripts/measure-token-cost.py` needs tiktoken. If system Python
  lacks it, use `uv run --with tiktoken python scripts/measure-token-cost.py`. CI
  installs tiktoken with pip.

## Conventions that differ from defaults

- Relative imports carry `.js` even when the source is `.ts` — NodeNext requires it:
  `import { authorize } from './auth.js'`.
- Every tool's `execute` returns a `string`. Format structured data as readable text or
  JSON; never return an object.
- Throw `UserError` (FastMCP) for anything the model should see; plain `Error` only for
  internal faults.
- Google clients initialize lazily via `getDocsClient()`, `getDriveClient()` and siblings —
  never at module load. `MCP_TOOL_GROUPS` is parsed at module load: unset or empty is
  the default five groups (docs, drive, sheets, slides, gmail), not `all`; unknown
  names log `FATAL: Invalid MCP_TOOL_GROUPS` and `process.exit(1)` before auth.
- `authorize()` returns `OAuth2Client | JWT`. `SERVICE_ACCOUNT_PATH` selects the JWT
  path; otherwise OAuth. Both types are passed as `auth` to the Google clients.
- Index bases differ per API: Docs content indices are 1-based, Sheets `GridRange` is
  0-based, Slides positions are in points (72pt = 1 inch). When a tool parameter
  states a different base (Docs table row/col are 0-based), match that parameter.
- Indentation in `src/server.ts` is mixed (column 0, two-space, four-space). Match the
  surrounding block; do not reformat a file you are editing.

## Known pitfalls

- Never add `console.log()` in `src/` — stdout carries the MCP protocol on the default
  stdio transport (`src/mcpTransport.ts:59-61`). Use `console.error()`. 26 calls already
  violate this in `googleDocsApiHelpers.ts` and `googleSlidesApiHelpers.ts` and the
  client tolerates them today; treat them as debt, not as the local style.
- Batch updates over 50 requests are split and executed sequentially, and document
  indices shift between batches (`googleDocsApiHelpers.ts:10`, `:73-75`). Build
  requests against the document state that batch will actually see, not the original.
  Inside a single Docs batch, order index-shifting edits from the end of the document
  to the beginning (`markdown-transformer/markdownToDocs.ts:1048-1049`); the API
  applies requests in array order.
- After adding an OAuth scope, delete `token.json` and re-authenticate — the old token
  lacks the new scope and fails confusingly.
- Do not restructure `src/auth.ts` to hard-fail on a bad `GOOGLE_REFRESH_TOKEN`. It falls
  through to `token.json` only for the recoverable errors listed at `auth.ts:77-83`
  (`invalid_grant`, `invalid_client`, expired, revoked) and throws on anything else
  (`:135`); an account mismatch on the env-var client skips `token.json` and goes
  straight to interactive OAuth (`:124-126`). All three paths are deliberate.
- Do not bypass `enforceRequiredAccount()`. When `REQUIRED_ACCOUNT_EMAIL` is set it pins
  the instance to one account and seeds `login_hint` (`auth.ts:360-362`); it runs at two
  sites, and `auth.ts:397-400` fires before the token is written, which is what stops a
  wrong-account login from ever reaching `token.json`. It is what keeps the Personal
  and Work instances apart.
- Do not add competing `uncaughtException`/`unhandledRejection` handlers, or duplicate
  `installTokenRefreshListener()` (registered at `auth.ts:117,153,396`). The handlers
  exit deliberately in three cases — dead stdio (EPIPE/EBADF) and a rejection storm of
  10 within 60s (`server.ts:117-161`). Those exits are not bugs.
- `downloadFile` retries through `exportViaWebUrl()` on any 403 from `files.export`
  (`server.ts:4077-4080`). That covers Drive's ~10MB export ceiling, but it is a
  status-code branch, not a size check, so a genuine permission-denied takes the same
  path. Do not add export paths that skip the fallback.
- Two helpers still throw `NotImplementedError`, not slow paths to optimize:
  `detectAndFormatLists` (behind `fixListFormatting`, so that tool always fails), and
  the dead `addCommentHelper`. The `addComment` tool itself works — it calls the Drive
  API directly. `findElement` is implemented (`findElements` in `googleDocsApiHelpers.ts`).
- Omitted `tabId` still reads and writes `document.body`, not the first document tab,
  even where some tool descriptions say "first tab". Do not implement a first-tab
  fallback as if the description were the contract. `replaceAllText` without `tabId`
  is whole-document on purpose.
- Programmatically created comments land in "All Comments" but are not visibly anchored
  in the Docs UI.

<!-- /bmad:context -->
