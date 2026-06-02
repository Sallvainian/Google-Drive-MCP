# Google-Drive-MCP -- Documentation Index

**Type:** Monolith Library (MCP Server)
**Language:** TypeScript (ES2022, NodeNext, strict)
**Framework:** FastMCP 3.35.0
**Tools:** 108 across 5 Google APIs
**Source:** 9 files, 10,279 LOC
**Generated:** 2026-05-05

---

## Quick Reference

| Attribute | Value |
|-----------|-------|
| **Entry Point** | `index.js` → `dist/server.js` |
| **Source** | `src/` (9 TypeScript files) |
| **Package Manager** | npm |
| **Build** | `npm run build` (tsc) |
| **Test** | `npm test` (`node --test tests/`) |
| **Auth** | OAuth 2.0 or Service Account (JWT) |
| **Account Lock** | `REQUIRED_ACCOUNT_EMAIL` env var |
| **Server Name** | Ultimate Google Docs & Sheets MCP Server |

### API Coverage

| Google API | Tools | Key Capabilities |
|------------|-------|-------------------|
| Docs v1 | 28 (incl. 6 comments + 4 formatted-doc tools) | Read/write/format documents, tables, images, comments, tabs, sections |
| Sheets v4 | 8 | Range read/write/append/clear, spreadsheet creation, cell formatting |
| Slides v1 | 17 | Presentations, slides, shapes, images, tables, speaker notes |
| Drive v3 | 21 | Files, folders, search, upload, download, copy, templates, permissions |
| Gmail v1 | 34 | Send/read/search, labels, filters, threads, drafts, attachments, batch ops |

---

## Generated Documentation

- [Project Overview](./project-overview.md) — Executive summary, tech stack, tool categories, API integrations, auth methods
- [Architecture](./architecture.md) — System design, module breakdown, data flow, auth flow, testing strategy
- [Source Tree Analysis](./source-tree-analysis.md) — Directory structure, critical files, dependency graph, code metrics
- [Development Guide](./development-guide.md) — Setup, commands, env vars, adding tools, MCP client config, error patterns

## Project Documentation (in repo root)

- [README](../README.md) — Full setup instructions, Google Cloud project setup, usage guide
- [claude.md](../claude.md) — AI assistant quick reference (tool categories, parameter patterns, source files)
- [SAMPLE_TASKS.md](../SAMPLE_TASKS.md) — 15 example workflows demonstrating tool usage
- [VS Code Guide](../vscode.md) — VS Code MCP extension integration setup

## GitHub Pages (OAuth Consent)

- [Landing Page](./index.html) — App landing page for Google OAuth verification
- [Privacy Policy](./privacy.html) — Required for Google OAuth consent screen
- [Terms of Service](./terms.html) — Required for Google OAuth consent screen

---

## Getting Started

1. Clone the repo and run `npm install`
2. Place `credentials.json` from Google Cloud Console in the project root
3. Run `npm run build` to compile TypeScript
4. Run `node ./dist/server.js` for first-time OAuth authorization
5. Configure your MCP client to use `node /path/to/index.js`

See the [Development Guide](./development-guide.md) for detailed instructions.

## AI-Assisted Development

When working with this codebase using AI tools:
- Start with `claude.md` in the repo root for tool reference
- Use [Architecture](./architecture.md) for system design and module responsibilities
- Use [Source Tree Analysis](./source-tree-analysis.md) for file locations and dependency graph
- Consult [SAMPLE_TASKS.md](../SAMPLE_TASKS.md) for example workflows showing tool composition
- See `_bmad-output/project-context.md` for the canonical 41-rule implementation guide for AI agents
