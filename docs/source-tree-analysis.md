# Source Tree Analysis

**Generated**: 2026-02-19

## Directory Structure

```
Google-Drive-MCP/
├── index.js                          # Entry point — imports dist/server.js
├── package.json                      # npm manifest (fastmcp, googleapis, zod)
├── tsconfig.json                     # TypeScript config (ES2022, NodeNext)
├── credentials.json                  # Google OAuth client secrets (gitignored)
├── token.json                        # OAuth refresh token (gitignored)
├── .envrc                            # direnv environment variables
├── .gitignore                        # Excludes credentials, dist, node_modules
│
├── src/                              # TypeScript source (9 files)
│   ├── server.ts                     # ★ MAIN: FastMCP server, all 99 tool definitions
│   ├── auth.ts                       # Authentication: OAuth2 + Service Account + token refresh
│   ├── types.ts                      # Zod schemas, type exports, hex color utils
│   ├── googleDocsApiHelpers.ts       # Docs: batch update, text search, style builders, tabs, images
│   ├── googleSheetsApiHelpers.ts     # Sheets: A1 notation, range ops, cell formatting
│   ├── googleSlidesApiHelpers.ts     # Slides: EMU conversion, element builders, batch update
│   ├── googleGmailApiHelpers.ts      # Gmail: MIME construction, message parsing, all operations
│   ├── gmailLabelManager.ts          # Gmail labels: CRUD, resolution, system labels
│   ├── gmailFilterManager.ts         # Gmail filters: CRUD, template-based creation
│   └── backup/                       # Backup copies (gitignored)
│       ├── auth.ts.bak
│       └── server.ts.bak
│
├── dist/                             # Compiled JS output (gitignored)
│   └── server.js                     # Compiled entry point
│
├── tests/                            # Tests (Node.js built-in test runner)
│   ├── types.test.js                 # Zod schema validation
│   ├── slides.test.js                # Slides helper unit tests
│   └── helpers.test.js               # Docs/Sheets helper tests
│
├── pages/                            # GitHub Pages source (OAuth consent screen)
│   ├── index.html                    # Landing page
│   ├── privacy.html                  # Privacy policy
│   ├── terms.html                    # Terms of service
│   └── pages.md                      # Pages config notes
│
├── docs/                             # Project documentation (this directory)
│   ├── index.html                    # GitHub Pages landing (OAuth verification)
│   ├── privacy.html                  # Privacy policy page
│   ├── terms.html                    # Terms of service page
│   ├── google25e86b582c5bc5ae.html   # Google site verification
│   └── *.md                          # Generated documentation files
│
├── assets/                           # Static assets
│   └── google.docs.mcp.1.gif         # Demo animation for README
│
├── .github/
│   └── workflows/
│       ├── claude.yml                # Claude Code PR review action
│       └── claude-code-review.yml    # Additional review workflow
│
├── README.md                         # Setup guide, feature list, usage
├── CLAUDE.md                         # AI assistant quick reference
├── SAMPLE_TASKS.md                   # 15 example workflows
├── vscode.md                         # VS Code MCP extension guide
├── LICENSE                           # ISC license
└── export-docs.mjs                   # One-off utility (gitignored)
```

## Critical Files

| File | Purpose | Why Critical |
|------|---------|-------------|
| `src/server.ts` | All 99 tool definitions | Core of the entire project — every MCP tool is registered here |
| `src/auth.ts` | Authentication routing | Controls all Google API access — OAuth2, Service Account, token refresh |
| `src/types.ts` | Zod schemas + types | Every tool parameter is validated through schemas defined here |
| `index.js` | Node entry point | What MCP clients actually execute to start the server |
| `credentials.json` | OAuth client secrets | Required for authentication (not in git) |
| `token.json` | Saved OAuth token | Persists user authorization (not in git) |

## Module Dependency Graph

```
index.js
  └── dist/server.js (compiled from src/server.ts)
        ├── src/auth.ts
        │     └── googleapis, google-auth-library, fs, http
        ├── src/types.ts
        │     └── zod, googleapis (docs_v1 types only)
        ├── src/googleDocsApiHelpers.ts
        │     └── googleapis, fastmcp (UserError), types.ts
        ├── src/googleSheetsApiHelpers.ts
        │     └── googleapis, fastmcp (UserError)
        ├── src/googleSlidesApiHelpers.ts
        │     └── googleapis, fastmcp (UserError)
        ├── src/googleGmailApiHelpers.ts
        │     └── googleapis, fastmcp (UserError), fs, path
        ├── src/gmailLabelManager.ts
        │     └── googleapis, fastmcp (UserError)
        └── src/gmailFilterManager.ts
              └── googleapis, fastmcp (UserError), types.ts
```

## Code Metrics

| File | Approximate LOC | Exports |
|------|----------------|---------|
| server.ts | ~3,800 | 0 (side-effect: registers tools & starts server) |
| auth.ts | ~310 | 1 (`authorize`) |
| types.ts | ~387 | ~40 (schemas + types + utils) |
| googleDocsApiHelpers.ts | ~936 | ~15 functions |
| googleSheetsApiHelpers.ts | ~427 | ~10 functions |
| googleSlidesApiHelpers.ts | ~595 | ~15 functions |
| googleGmailApiHelpers.ts | ~802 | ~20 functions |
| gmailLabelManager.ts | ~299 | ~10 functions |
| gmailFilterManager.ts | ~321 | ~8 functions |
| **Total** | **~7,877** | |
