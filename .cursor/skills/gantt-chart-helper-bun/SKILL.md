---
name: gantt-chart-helper-bun
description: >-
  The gantt-chart-helper repository uses Bun as its package manager and script
  runner. Use when installing dependencies, adding packages, running dev/build/lint
  scripts, or invoking local CLIs for this project. Prefer bun over npm, npx, pnpm,
  or yarn unless the user explicitly asks otherwise.
---

# gantt-chart-helper: use Bun

## Commands

- Install: `bun install`
- Add dependency: `bun add <pkg>` (dev: `bun add -d <pkg>`)
- Run scripts from `package.json`: `bun run <script>` (e.g. `bun run dev`, `bun run build`, `bun run lint`)
- One-off / CLIs: `bunx <command>` (equivalent to npx)

## Do not

- Default to `npm`, `npx`, `pnpm`, or `yarn` for this repo in examples or terminal commands.

## Lockfile

- Prefer the existing Bun lockfile; do not introduce a different package manager’s lockfile without an explicit user request.
