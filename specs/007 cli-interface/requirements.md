# CLI Interface - Requirements

## Overview

Add a command-line interface (`myoffice`) that exposes the existing M365 MCP tools for direct terminal use. Primary use case is enabling Claude to call tools via the Bash tool without MCP protocol overhead, while also being usable directly by humans in terminal.

## User Stories

**As a** Claude Code user
**I want** Claude to call M365 tools directly via Bash
**So that** I can use M365 functionality without requiring MCP server setup

**As a** developer
**I want** to query my M365 data from the terminal
**So that** I can script and automate personal workflows

**As a** power user
**I want** human-readable output by default with JSON option
**So that** I can read results easily or pipe to other tools

## Acceptance Criteria

### Authentication

1. WHEN user runs any command without authentication THEN system SHALL display error with instructions to run `myoffice login`
2. WHEN user runs `myoffice login` THEN system SHALL initiate device code flow and cache tokens
3. WHEN token is expired THEN system SHALL silently refresh using cached refresh token
4. WHEN refresh fails THEN system SHALL display error with instructions to re-authenticate

### Command Structure

5. WHEN user runs `myoffice` with no arguments THEN system SHALL display help with available categories
6. WHEN user runs `myoffice <category>` THEN system SHALL display help for that category's commands
7. WHEN user runs `myoffice <category> <action>` THEN system SHALL execute the corresponding tool
8. WHEN user provides `--help` on any command THEN system SHALL display usage and available flags

### Output Formatting

9. WHEN command succeeds THEN system SHALL output human-readable format by default
10. WHEN user provides `--json` flag THEN system SHALL output JSON format
11. WHEN command fails THEN system SHALL output error message to stderr with non-zero exit code
12. IF output is JSON AND command fails THEN system SHALL output error as JSON object with `error` field

### Tool Coverage

13. WHEN any MCP tool exists THEN system SHALL expose equivalent CLI command
14. WHEN tool has optional parameters THEN system SHALL expose them as optional flags
15. WHEN tool has required parameters THEN system SHALL validate presence and show error if missing

## Command Categories

Based on existing tools:

| Category | Commands |
|----------|----------|
| `mail` | list, read, send, reply, search, move, delete |
| `calendar` | list, create, update, delete |
| `tasks` | lists, list, create, complete, delete |
| `files` | list, search, read, download, upload, shared |
| `sharepoint` | sites, libraries, files |
| `contacts` | list, search, create |

## Edge Cases

- Command called with invalid category → show "unknown category" + list valid ones
- Command called with invalid action → show "unknown command" + list valid ones for category
- Network error during API call → show error, exit code 1
- Token file missing or corrupted → prompt to re-authenticate
- API returns empty results → show "No results found" (human) or `[]` (JSON)
- API rate limiting → show error with retry suggestion

## Constraints

- Must reuse existing tool implementations (no duplication)
- Must reuse existing auth flow (device code + token caching)
- Must work on macOS (primary), Linux, Windows
- Must be installable globally via npm
- Exit codes: 0 = success, 1 = error
- No interactive prompts during command execution (except login)

## Out of Scope

- Interactive/TUI mode (e.g., browsing emails interactively)
- Shell completions (can add later)
- Config file for defaults (can add later)
- Multiple account support
- Offline mode / caching results
