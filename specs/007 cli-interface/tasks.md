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
    - Subcommands: list, search, get, create, update
    - Requirements: R7, R14

- [ ] 10. **Implement teams commands**
  - [ ] 10.1 Create teams command group
    - Files: `src/cli/commands/teams.ts` (new)
    - Subcommands: list, channels, messages, post
    - Requirements: R7, R14

- [ ] 11. **Implement chats commands**
  - [ ] 11.1 Create chats command group
    - Files: `src/cli/commands/chats.ts` (new)
    - Subcommands: list, messages, send, create
    - Requirements: R7, R14

- [ ] 12. **Implement planner commands**
  - [ ] 12.1 Create planner command group
    - Files: `src/cli/commands/planner.ts` (new)
    - Subcommands: plans, plan, buckets, bucket-create, bucket-update, bucket-delete, tasks, task, task-create, task-update, task-delete, task-details, task-details-update
    - Requirements: R7, R14

### Phase 5: Output Formatting

- [ ] 13. **Create output formatter**
  - [ ] 13.1 Implement formatter module
    - Files: `src/cli/formatter.ts` (new)
    - JSON output mode (simple stringify)
    - Human-readable tables for lists
    - Human-readable key-value for single items
    - Requirements: R9, R10
  - [ ] 13.2 Apply formatter to all commands
    - Files: `src/cli/commands/*.ts`
    - Use `--json` flag to switch modes

### Phase 6: Error Handling

- [ ] 14. **Implement error handling**
  - [ ] 14.1 Add auth check wrapper
    - Files: `src/cli.ts` or `src/cli/middleware.ts` (new)
    - Check auth before running commands (except login/status)
    - Show helpful error message with login instructions
    - Requirements: R1, R11, R12
  - [ ] 14.2 Handle API errors gracefully
    - Files: `src/cli/commands/*.ts`
    - Catch errors, format for human/JSON
    - Set exit code 1 on error
    - Requirements: R11, R12

### Phase 7: Build & Package

- [ ] 15. **Finalize build configuration**
  - [ ] 15.1 Update TypeScript config if needed
    - Files: `tsconfig.json`
    - Ensure `cli.ts` is included in build
  - [ ] 15.2 Add shebang to cli.ts
    - Files: `src/cli.ts`
    - Add: `#!/usr/bin/env node`
  - [ ] 15.3 Update package.json scripts
    - Files: `package.json`
    - Update login script to use CLI
    - Add prepublish script

### Phase 8: Documentation

- [ ] 16. **Documentation**
  - [ ] 16.1 Update README with CLI usage
    - Files: `README.md`
    - Installation from GitHub
    - Basic usage examples
    - Command reference
  - [ ] 16.2 Update CLAUDE.md
    - Files: `CLAUDE.md`
    - Note about dual entry points (MCP + CLI)
  - [ ] 16.3 Add entry to Awolve Handbook
    - Location: Awolve Handbook (tools section)
    - Content: Installation, setup, usage examples
    - Link to GitHub repo

### Phase 9: Claude Code Skill

- [ ] 17. **Create Claude Code skill for Awolve Office**
  - [ ] 17.1 Create skill file
    - Files: `skills/awolve-office.md` (new)
    - Skill name and description in frontmatter
    - Trigger conditions for personal Microsoft 365 operations
    - Clear distinction from: (1) personal Gmail/Google, (2) M365 MCP admin tool
    - Explain: MyOffice = personal user ops, M365 MCP = admin ops
    - Installation instructions (npm install from GitHub)
    - Environment setup (M365_CLIENT_ID, reference to handbook)
    - Authentication flow (myoffice login)
    - Error handling guidance (not installed, not authenticated, env var missing)
    - CLI command reference with --json flag usage
    - Example workflows for common operations
  - [ ] 17.2 Document skill installation
    - Add instructions for copying skill to `~/.claude/skills/` or project `.claude/skills/`

## Implementation Order

Recommended sequence:

1. **Phase 2 first** - Extract handler so we can test incrementally
2. **Phase 3** - Get basic CLI working with login/status
3. **Phase 4** - Start with mail (most used)
4. **Phase 5 (task 13)** - Add formatter early for consistent output
5. **Phases 4-12** - Add remaining command groups (calendar, tasks, files, sharepoint, contacts, teams, chats, planner)
6. **Phase 6 (task 14)** - Polish error handling
7. **Phases 7, 8** - Finalize build and docs
8. **Phase 9** - Create Claude Code skill for Awolve Office

## Dependencies

**npm packages to add:**
- `commander` - CLI framework (~50KB)

**No other new dependencies needed** - existing auth and tools are reused.

**Distribution:**
- Install from GitHub: `npm install -g github:awolve/ops-myoffice`
- Document in Awolve Handbook

## Verification Checklist

After implementation, verify:

- [ ] `npm run build` succeeds
- [ ] `npm link` installs `myoffice` command
- [ ] `myoffice --help` shows all categories (mail, calendar, tasks, files, sharepoint, contacts, teams, chats, planner)
- [ ] `myoffice mail --help` shows mail actions
- [ ] `myoffice mail list --help` shows options
- [ ] `myoffice login` initiates device code flow
- [ ] `myoffice status` shows auth state
- [ ] `myoffice debug` shows debug info
- [ ] `myoffice mail list` returns emails (human format)
- [ ] `myoffice mail list --json` returns emails (JSON)
- [ ] `myoffice teams list` returns teams
- [ ] `myoffice chats list` returns chats
- [ ] `myoffice planner plans` returns plans
- [ ] Commands without auth show helpful error
- [ ] Invalid commands show helpful error
- [ ] MCP server still works (`npm run dev` + Claude Code)
- [ ] Claude Code skill file exists at `skills/awolve-office.md`
- [ ] Skill has correct frontmatter (name, description, trigger conditions)
- [ ] Skill distinguishes MyOffice (personal) from M365 MCP (admin) and Gmail/Google
- [ ] Skill includes installation instructions (npm install from GitHub)
- [ ] Skill includes env var setup guidance (M365_CLIENT_ID)
- [ ] Skill includes authentication guidance (myoffice login)
- [ ] Skill handles error cases (not installed, not authenticated)
