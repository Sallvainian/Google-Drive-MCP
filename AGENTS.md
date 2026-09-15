<!-- bmad:context -->
<!-- Verified 2026-09-14 against 53337ef. Managed by bmad-project-context; edits inside
     this block are replaced on refresh. Keep anything you want preserved outside the markers. -->

## Google-Drive-MCP

FastMCP server exposing 108 tools over five Google Workspace APIs (Docs, Sheets, Slides,
Drive, Gmail). TypeScript ESM on npm, no framework beyond FastMCP. Generated reference
docs live in `docs/` — they are regenerated wholesale, not hand-edited.

## Policy

- Never commit `credentials.json` or `token.json` — OAuth secrets, gitignored at `.gitignore:6-7`.
- `.claude/` is tracked deliberately; `.grok/`, `.codex/`, `.cursor/` and `_bmad/` stay
  ignored. This flipped across six commits — do not re-litigate it. `_bmad-output/` is
  ignored. `_bmad-output/specs/` is un-ignored so a bmad-loop story re-drive can read its
  COMMITTED spec from a fresh worktree, but those specs are NOT published: they index
  unfixed findings by file and line, and this repo is public. Keep them local.

## Where things are

- All 108 tool definitions: `src/server.ts` (6015 lines), grouped by API in the order
  Docs, Sheets, Slides, Drive, Gmail. Add a new tool to its API's section.
- One helper module per API — `googleDocsApiHelpers.ts`, `googleSheetsApiHelpers.ts`,
  `googleSlidesApiHelpers.ts`, `googleGmailApiHelpers.ts`, plus `gmailLabelManager.ts`
  and `gmailFilterManager.ts`. Helpers take the typed API client as first parameter.
- Auth, token handling and account pinning: `src/auth.ts`.
- Reusable Zod schemas: `src/types.ts`; one-off tool schemas stay inline in `server.ts`.

## Running and verifying

- Build first — `npm run build`. Tests import from `dist/`, which is gitignored, so a
  fresh clone has nothing to run against.
- `npm test` runs the suite — the script is `node --test tests/*.test.js`. It used to be
  `node --test tests/`, which Node 24 resolves as a module and fails with
  MODULE_NOT_FOUND; the script itself was fixed, so no workaround is needed.
- That suite is 45 tests and all 45 pass. The `tests/helpers.test.js` `fields` expectation
  that used to fail was corrected to match what `googleDocsApiHelpers.ts:79` requests.
- No linter or formatter is configured. Do not add one uninvited; there is no `lint` script.
- `tsconfig.json` covers `src/**/*` only — `tests/` is never typechecked.
- CI runs no build and no tests; both workflows only invoke `anthropics/claude-code-action@v1`.
  Verify locally, nothing will catch it for you.
- No Node version is declared (no `engines`, no `.nvmrc`). README says 18+; both CI
  workflows pin Node 24.

## Conventions that differ from defaults

- Relative imports carry `.js` even when the source is `.ts` — NodeNext requires it:
  `import { authorize } from './auth.js'`.
- Every tool's `execute` returns a `string`. Format structured data as readable text or
  JSON; never return an object.
- Throw `UserError` (FastMCP) for anything the model should see; plain `Error` only for
  internal faults.
- Google clients initialize lazily via `getDocsClient()`, `getDriveClient()` and siblings —
  never at module load.
- Index bases differ per API: Docs 1-based, Sheets 0-based, Slides positions in points
  (72pt = 1 inch).
- Indentation is genuinely inconsistent — `src/server.ts` splits 925 lines at column 0,
  946 at two-space, 996 at four-space. Match the surrounding block; do not reformat a file
  you are editing.

## Known pitfalls

- Never add `console.log()` in `src/` — stdout carries the MCP protocol on the default
  stdio transport (`server.ts:5975-5981`). Use `console.error()`. 26 calls already violate
  this in `googleDocsApiHelpers.ts` and `googleSlidesApiHelpers.ts` and the client
  tolerates them today; treat them as debt, not as the local style.
- Batch updates over 50 requests are split and executed sequentially, and document indices
  shift between batches (`googleDocsApiHelpers.ts:18-19`). Build requests against the
  document state that batch will actually see, not the original.
- After adding an OAuth scope, delete `token.json` and re-authenticate — the old token
  lacks the new scope and fails confusingly.
- Do not restructure `src/auth.ts` to hard-fail on a bad `GOOGLE_REFRESH_TOKEN`. It falls
  through to `token.json` only for the recoverable errors listed at `auth.ts:76-83`
  (`invalid_grant`, `invalid_client`, expired, revoked) and throws on anything else
  (`:134`); an account mismatch instead skips `token.json` and goes straight to interactive
  OAuth (`:125`). All three paths are deliberate.
- Do not bypass `enforceRequiredAccount()`. When `REQUIRED_ACCOUNT_EMAIL` is set it pins
  the instance to one account and seeds `login_hint`; it runs at two sites, and `auth.ts:314`
  fires before the token is written, which is what stops a wrong-account login from ever
  reaching `token.json`. It is what keeps the Personal and Work instances apart.
- Do not add competing `uncaughtException`/`unhandledRejection` handlers, or duplicate
  `installTokenRefreshListener()` (registered at `auth.ts:116,152,312`). The handlers exit
  deliberately in three cases — dead stdio (EPIPE/EBADF) and a rejection storm of 10 within
  60s (`server.ts:152`). Those exits are not bugs.
- `downloadFile` retries through `exportViaWebUrl()` on any 403 from `files.export`
  (`server.ts:2653`). That covers Drive's ~10MB export ceiling, but it is a status-code
  branch, not a size check, so a genuine permission-denied takes the same path. Do not add
  export paths that skip the fallback.
- Three helpers are stubs that throw `NotImplementedError`, not slow paths to optimize:
  `findElement`, `detectAndFormatLists` (behind `fixListFormatting`, so that tool always
  fails), and the dead `addCommentHelper`. The `addComment` tool itself works — it calls
  the Drive API directly.
- Programmatically created comments land in "All Comments" but are not visibly anchored
  in the Docs UI.

<!-- /bmad:context -->
