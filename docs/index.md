# Google-Drive-MCP — Documentation Index

**Type:** Monolith Library (MCP Server)
**Language:** TypeScript
**Framework:** FastMCP 3.24.0
**Tools:** 99 across 5 Google APIs
**Generated:** 2026-02-19

---

## Quick Reference

| Attribute | Value |
|-----------|-------|
| **Entry Point** | `index.js` → `dist/server.js` |
| **Source** | `src/` (9 TypeScript files) |
| **Package Manager** | npm |
| **Build** | `npm run build` (tsc) |
| **Test** | `npm test` (node --test) |
| **Auth** | OAuth 2.0 or Service Account |

### API Coverage

| Google API | Tools | Key Capabilities |
|------------|-------|-------------------|
| Docs v1 | 22 | Read/write/format documents, tables, images, comments, tabs |
| Sheets v4 | 8 | Range read/write/append/clear, spreadsheet creation |
| Slides v1 | 17 | Presentations, slides, shapes, images, tables, notes |
| Drive v3 | 18 | Files, folders, search, upload, download, templates |
| Gmail v1 | 34 | Send/read/search, labels, filters, threads, drafts, attachments |

---

## Generated Documentation

- [Project Overview](./project-overview.md) — Executive summary, tech stack, tool categories, API integrations
- [Architecture](./architecture.md) — System design, module breakdown, tool catalog, auth flow, data flow
- [Source Tree Analysis](./source-tree-analysis.md) — Directory structure, critical files, dependency graph, code metrics
- [Development Guide](./development-guide.md) — Setup, commands, env vars, adding tools, MCP client config

## Project Documentation (in repo root)

- [README](../README.md) — Full setup instructions, Google Cloud project setup, usage guide
- [CLAUDE.md](../CLAUDE.md) — AI assistant quick reference (tool categories, parameter patterns, source files)
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
