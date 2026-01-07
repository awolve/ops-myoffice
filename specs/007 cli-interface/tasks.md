# CLI Interface - Tasks

## Overview

Estimated scope: **Medium** (2-3 days of work)

## Tasks

### Phase 1: Foundation

- [ ] 1. **Project setup and dependencies**
  - [ ] 1.1 Add Commander.js dependency
    - Files: `package.json`
    - Run: `npm install commander`
  - [ ] 1.2 Update package.json with new bin entry and package name
    - Files: `package.json`
    - Add: `"myoffice": "dist/cli.js"` to bin
    - Consider: rename package or use scoped name

### Phase 2: Extract Shared Handler

- [ ] 2. **Create shared tool dispatch layer**
  - [ ] 2.1 Create core handler module
    - Files: `src/core/handler.ts` (new)
    - Extract switch statement from `src/index.ts`
    - Define `ToolResult` interface
    - Export `executeCommand()` function
    - Requirements: R13, R14, R15
  - [ ] 2.2 Refactor MCP server to use handler
    - Files: `src/index.ts`
    - Import and call `executeCommand()`
    - Keep MCP response formatting
    - Test: MCP server still works via `npm run dev`

### Phase 3: CLI Entry Point

- [ ] 3. **Build CLI structure**
  - [ ] 3.1 Create CLI entry point
    - Files: `src/cli.ts` (new)
    - Setup Commander.js program
    - Add `--version` and `--help` flags
    - Add global `--json` flag
    - Requirements: R5, R8
  - [ ] 3.2 Add login command
    - Files: `src/cli.ts`
    - Integrate existing device code flow
    - Handle success/failure output
    - Requirements: R2, R3
  - [ ] 3.3 Add status command
    - Files: `src/cli.ts`
    - Call `auth_status` tool
    - Format output
    - Requirements: R1

### Phase 4: Command Categories

- [ ] 4. **Implement mail commands**
  - [ ] 4.1 Create mail command group
    - Files: `src/cli/commands/mail.ts` (new)
    - Subcommands: list, read, send, reply, search, delete, mark
    - Map flags to tool parameters
    - Requirements: R7, R14
  - [ ] 4.2 Register mail commands with main program
    - Files: `src/cli.ts`

- [ ] 5. **Implement calendar commands**
  - [ ] 5.1 Create calendar command group
    - Files: `src/cli/commands/calendar.ts` (new)
    - Subcommands: list, get, create, update, delete
    - Requirements: R7, R14

- [ ] 6. **Implement tasks commands**
  - [ ] 6.1 Create tasks command group
    - Files: `src/cli/commands/tasks.ts` (new)
    - Subcommands: lists, list, create, update, complete, delete
    - Requirements: R7, R14

- [ ] 7. **Implement files commands**
  - [ ] 7.1 Create files command group (OneDrive)
    - Files: `src/cli/commands/files.ts` (new)
    - Subcommands: list, get, search, read, mkdir, shared
    - Requirements: R7, R14

- [ ] 8. **Implement sharepoint commands**
  - [ ] 8.1 Create sharepoint command group
    - Files: `src/cli/commands/sharepoint.ts` (new)
    - Subcommands: sites, site, drives, files, file, read, search
    - Requirements: R7, R14

- [ ] 9. **Implement contacts commands**
  - [ ] 9.1 Create contacts command group
    - Files: `src/cli/commands/contacts.ts` (new)
    - Subcommands: list, search, get
    - Requirements: R7, R14

### Phase 5: Output Formatting

- [ ] 10. **Create output formatter**
  - [ ] 10.1 Implement formatter module
    - Files: `src/cli/formatter.ts` (new)
    - JSON output mode (simple stringify)
    - Human-readable tables for lists
    - Human-readable key-value for single items
    - Requirements: R9, R10
  - [ ] 10.2 Apply formatter to all commands
    - Files: `src/cli/commands/*.ts`
    - Use `--json` flag to switch modes

### Phase 6: Error Handling

- [ ] 11. **Implement error handling**
  - [ ] 11.1 Add auth check wrapper
    - Files: `src/cli.ts` or `src/cli/middleware.ts` (new)
    - Check auth before running commands (except login/status)
    - Show helpful error message with login instructions
    - Requirements: R1, R11, R12
  - [ ] 11.2 Handle API errors gracefully
    - Files: `src/cli/commands/*.ts`
    - Catch errors, format for human/JSON
    - Set exit code 1 on error
    - Requirements: R11, R12

### Phase 7: Build & Package

- [ ] 12. **Finalize build configuration**
  - [ ] 12.1 Update TypeScript config if needed
    - Files: `tsconfig.json`
    - Ensure `cli.ts` is included in build
  - [ ] 12.2 Add shebang to cli.ts
    - Files: `src/cli.ts`
    - Add: `#!/usr/bin/env node`
  - [ ] 12.3 Update package.json scripts
    - Files: `package.json`
    - Update login script to use CLI
    - Add prepublish script

### Phase 8: Documentation

- [ ] 13. **Documentation**
  - [ ] 13.1 Update README with CLI usage
    - Files: `README.md`
    - Installation from GitHub
    - Basic usage examples
    - Command reference
  - [ ] 13.2 Update CLAUDE.md
    - Files: `CLAUDE.md`
    - Note about dual entry points (MCP + CLI)
  - [ ] 13.3 Add entry to Awolve Handbook
    - Location: Awolve Handbook (tools section)
    - Content: Installation, setup, usage examples
    - Link to GitHub repo

## Implementation Order

Recommended sequence:

1. **Phase 2 first** - Extract handler so we can test incrementally
2. **Phase 3** - Get basic CLI working with login/status
3. **Phase 4** - Start with mail (most used)
4. **Phase 10** - Add formatter early for consistent output
5. **Phases 5-9** - Add remaining command groups
6. **Phase 11** - Polish error handling
7. **Phases 7, 8** - Finalize build and docs

## Dependencies

**npm packages to add:**
- `commander` - CLI framework (~50KB)

**No other new dependencies needed** - existing auth and tools are reused.

**Distribution:**
- Install from GitHub: `npm install -g github:awolve/ops-personal-m365-mcp`
- Document in Awolve Handbook

## Verification Checklist

After implementation, verify:

- [ ] `npm run build` succeeds
- [ ] `npm link` installs `myoffice` command
- [ ] `myoffice --help` shows all categories
- [ ] `myoffice mail --help` shows mail actions
- [ ] `myoffice mail list --help` shows options
- [ ] `myoffice login` initiates device code flow
- [ ] `myoffice status` shows auth state
- [ ] `myoffice mail list` returns emails (human format)
- [ ] `myoffice mail list --json` returns emails (JSON)
- [ ] Commands without auth show helpful error
- [ ] Invalid commands show helpful error
- [ ] MCP server still works (`npm run dev` + Claude Code)
