# CLI Interface - Design

## Overview

Add a CLI entry point (`myoffice`) alongside the existing MCP server. The CLI reuses all existing tool implementations, auth layer, and Graph client - only adding argument parsing and output formatting.

## Architecture

```
                    ┌─────────────────┐
                    │   User/Claude   │
                    └────────┬────────┘
                             │
              ┌──────────────┴──────────────┐
              │                             │
    ┌─────────▼─────────┐       ┌───────────▼───────────┐
    │   src/cli.ts      │       │   src/index.ts        │
    │   (CLI entry)     │       │   (MCP server)        │
    └─────────┬─────────┘       └───────────┬───────────┘
              │                             │
              │     ┌───────────────────────┘
              │     │
    ┌─────────▼─────▼─────────┐
    │   src/core/handler.ts   │
    │   (shared tool dispatch) │
    └─────────────┬───────────┘
                  │
    ┌─────────────▼───────────┐
    │   src/tools/*           │
    │   (tool implementations) │
    └─────────────┬───────────┘
                  │
    ┌─────────────▼───────────┐
    │   src/auth/             │
    │   src/utils/graph-client│
    └─────────────────────────┘
```

Both entry points share:
- Tool implementations (unchanged)
- Auth layer (unchanged)
- Graph client (unchanged)

Only the presentation layer differs.

## Components

### New: src/cli.ts

**Purpose:** CLI entry point - parse arguments, invoke handler, format output

**Interface:**
```bash
myoffice <category> <action> [options]
myoffice mail list --unread --limit 10
myoffice calendar list --start 2024-01-01
myoffice login
myoffice --help
myoffice --version
```

**Key responsibilities:**
- Parse command-line arguments using Commander.js
- Map CLI flags to tool parameters
- Call shared handler
- Format output (human-readable or JSON)
- Handle errors with appropriate exit codes

### New: src/core/handler.ts

**Purpose:** Shared tool dispatch logic extracted from index.ts

**Interface:**
```typescript
interface ToolResult {
  success: boolean;
  data?: unknown;
  error?: string;
}

async function executeCommand(
  toolName: string,
  args: Record<string, unknown>
): Promise<ToolResult>
```

**Key responsibilities:**
- Tool name → function mapping (the big switch statement)
- Zod validation
- Error normalization

### New: src/cli/formatter.ts

**Purpose:** Format tool results for terminal output

**Interface:**
```typescript
function formatOutput(
  result: ToolResult,
  options: { json?: boolean }
): string
```

**Formats:**
- Default: Human-readable (tables for lists, key-value for objects)
- `--json`: Raw JSON (same as MCP output)

### Modified: src/index.ts

**Changes:**
- Extract tool dispatch logic to `core/handler.ts`
- Call handler instead of inline switch
- Keep MCP-specific wrapping (content array format)

### Modified: package.json

**Changes:**
```json
{
  "bin": {
    "myoffice": "dist/cli.js",
    "ops-personal-m365-mcp": "dist/index.js"
  },
  "scripts": {
    "login": "tsx src/cli.ts login"
  },
  "dependencies": {
    "commander": "^12.0.0"  // Add CLI framework
  }
}
```

## Data Models

No new data models. Tool inputs/outputs unchanged.

## Command Mapping

| CLI Command | Tool Name | Example |
|-------------|-----------|---------|
| **Mail** |||
| `mail list` | `mail_list` | `myoffice mail list --unread` |
| `mail read` | `mail_read` | `myoffice mail read --id ABC123` |
| `mail send` | `mail_send` | `myoffice mail send --to x@y.com --subject Hi --body Hello` |
| `mail reply` | `mail_reply` | `myoffice mail reply --id ABC123 --body "Thanks"` |
| `mail search` | `mail_search` | `myoffice mail search --query "invoice"` |
| `mail delete` | `mail_delete` | `myoffice mail delete --id ABC123` |
| `mail mark` | `mail_mark_read` | `myoffice mail mark --id ABC123 --unread` |
| **Calendar** |||
| `calendar list` | `calendar_list` | `myoffice calendar list --start 2024-01-01` |
| `calendar get` | `calendar_get` | `myoffice calendar get --id ABC123` |
| `calendar create` | `calendar_create` | `myoffice calendar create --subject "Meeting" --start ...` |
| `calendar update` | `calendar_update` | `myoffice calendar update --id ABC123 --subject "New title"` |
| `calendar delete` | `calendar_delete` | `myoffice calendar delete --id ABC123` |
| **Tasks (To Do)** |||
| `tasks lists` | `tasks_list_lists` | `myoffice tasks lists` |
| `tasks list` | `tasks_list` | `myoffice tasks list --list-id ABC123` |
| `tasks create` | `tasks_create` | `myoffice tasks create --title "Buy milk"` |
| `tasks update` | `tasks_update` | `myoffice tasks update --id ABC123 --title "New title"` |
| `tasks complete` | `tasks_complete` | `myoffice tasks complete --id ABC123` |
| `tasks delete` | `tasks_delete` | `myoffice tasks delete --id ABC123` |
| **Files (OneDrive)** |||
| `files list` | `onedrive_list` | `myoffice files list --path /Documents` |
| `files get` | `onedrive_get` | `myoffice files get --path /Documents/file.txt` |
| `files search` | `onedrive_search` | `myoffice files search --query "report"` |
| `files read` | `onedrive_read` | `myoffice files read --path /Documents/file.txt` |
| `files mkdir` | `onedrive_create_folder` | `myoffice files mkdir --name "New Folder"` |
| `files shared` | `onedrive_shared_with_me` | `myoffice files shared` |
| **SharePoint** |||
| `sharepoint sites` | `sharepoint_list_sites` | `myoffice sharepoint sites` |
| `sharepoint site` | `sharepoint_get_site` | `myoffice sharepoint site --id ABC123` |
| `sharepoint drives` | `sharepoint_list_drives` | `myoffice sharepoint drives --site-id ABC123` |
| `sharepoint files` | `sharepoint_list_files` | `myoffice sharepoint files --drive-id ABC123` |
| `sharepoint file` | `sharepoint_get_file` | `myoffice sharepoint file --drive-id ABC --path /doc.txt` |
| `sharepoint read` | `sharepoint_read_file` | `myoffice sharepoint read --drive-id ABC --path /doc.txt` |
| `sharepoint search` | `sharepoint_search_files` | `myoffice sharepoint search --drive-id ABC --query "budget"` |
| **Contacts** |||
| `contacts list` | `contacts_list` | `myoffice contacts list` |
| `contacts search` | `contacts_search` | `myoffice contacts search --query "john"` |
| `contacts get` | `contacts_get` | `myoffice contacts get --id ABC123` |
| `contacts create` | `contacts_create` | `myoffice contacts create --given-name John --email j@x.com` |
| `contacts update` | `contacts_update` | `myoffice contacts update --id ABC123 --job-title "Manager"` |
| **Teams** |||
| `teams list` | `teams_list` | `myoffice teams list` |
| `teams channels` | `teams_channels` | `myoffice teams channels --team-id ABC123` |
| `teams messages` | `teams_channel_messages` | `myoffice teams messages --team-id ABC --channel-id XYZ` |
| `teams post` | `teams_channel_post` | `myoffice teams post --team-id ABC --channel-id XYZ --content "Hello"` |
| **Chats** |||
| `chats list` | `chats_list` | `myoffice chats list` |
| `chats messages` | `chats_messages` | `myoffice chats messages --chat-id ABC123` |
| `chats send` | `chats_send` | `myoffice chats send --chat-id ABC123 --content "Hi"` |
| `chats create` | `chats_create` | `myoffice chats create --members user@example.com` |
| **Planner** |||
| `planner plans` | `planner_list_plans` | `myoffice planner plans` |
| `planner plan` | `planner_get_plan` | `myoffice planner plan --id ABC123` |
| `planner buckets` | `planner_list_buckets` | `myoffice planner buckets --plan-id ABC123` |
| `planner bucket-create` | `planner_create_bucket` | `myoffice planner bucket-create --plan-id ABC --name "To Do"` |
| `planner bucket-update` | `planner_update_bucket` | `myoffice planner bucket-update --id ABC --name "Done"` |
| `planner bucket-delete` | `planner_delete_bucket` | `myoffice planner bucket-delete --id ABC123` |
| `planner tasks` | `planner_list_tasks` | `myoffice planner tasks --plan-id ABC123` |
| `planner task` | `planner_get_task` | `myoffice planner task --id ABC123` |
| `planner task-create` | `planner_create_task` | `myoffice planner task-create --plan-id ABC --title "Do thing"` |
| `planner task-update` | `planner_update_task` | `myoffice planner task-update --id ABC --progress completed` |
| `planner task-delete` | `planner_delete_task` | `myoffice planner task-delete --id ABC123` |
| `planner task-details` | `planner_get_task_details` | `myoffice planner task-details --id ABC123` |
| `planner task-details-update` | `planner_update_task_details` | `myoffice planner task-details-update --id ABC --description "..."` |
| **System** |||
| `login` | (special) | `myoffice login` |
| `status` | `auth_status` | `myoffice status` |
| `debug` | `debug_info` | `myoffice debug` |

**Naming convention:**
- Category names are user-friendly (`files` not `onedrive`)
- Actions match tool suffix after `_`
- Parameters become `--flag-name` (kebab-case)
- Compound actions use hyphenated names (`bucket-create`, `task-details`)

## Help Output Specification

### Root Help: `myoffice --help`

```
myoffice - Access your Microsoft 365 data from the command line

Usage: myoffice <command> [options]

Commands:
  mail        Email operations (list, read, send, reply, search, delete, mark)
  calendar    Calendar events (list, get, create, update, delete)
  tasks       Microsoft To Do (lists, list, create, update, complete, delete)
  files       OneDrive files (list, get, search, read, mkdir, shared)
  sharepoint  SharePoint sites and document libraries
  contacts    Contacts (list, search, get, create, update)
  teams       Teams channels and messages
  chats       1:1 and group chats
  planner     Planner plans, buckets, and tasks
  login       Authenticate with Microsoft 365
  status      Check authentication status
  debug       Show debug information

Options:
  --json      Output as JSON (default: human-readable)
  --help      Show help
  --version   Show version

Examples:
  myoffice mail list --unread
  myoffice calendar list --start 2024-01-15
  myoffice files list --path /Documents
  myoffice planner tasks --plan-id ABC123
  myoffice login

Run 'myoffice <command> --help' for command-specific options.
```

### Category Help: `myoffice mail --help`

```
myoffice mail - Email operations

Usage: myoffice mail <action> [options]

Actions:
  list      List emails from a folder
  read      Read a specific email
  send      Send a new email
  reply     Reply to an email
  search    Search emails
  delete    Delete an email
  mark      Mark email as read/unread

Global Options:
  --json    Output as JSON

Run 'myoffice mail <action> --help' for action-specific options.
```

### Action Help: `myoffice mail list --help`

```
myoffice mail list - List emails from a folder

Usage: myoffice mail list [options]

Options:
  --folder <name>   Folder name (default: inbox)
  --limit <n>       Maximum emails to return (default: 25)
  --unread          Only show unread emails
  --json            Output as JSON

Examples:
  myoffice mail list
  myoffice mail list --unread --limit 10
  myoffice mail list --folder sentitems
```

### Example Output Formats

**Human-readable (default):**
```
$ myoffice mail list --limit 3

 ID          FROM                  SUBJECT                    DATE
 AAMk...abc  john@example.com      Meeting tomorrow           Jan 15, 10:30 AM
 AAMk...def  jane@company.com      Q4 Report                  Jan 14, 3:45 PM
 AAMk...ghi  notifications@...     Your order shipped         Jan 14, 9:00 AM

3 emails (2 unread)
```

**JSON output (`--json`):**
```json
{
  "emails": [
    {
      "id": "AAMk...abc",
      "from": "john@example.com",
      "subject": "Meeting tomorrow",
      "receivedDateTime": "2024-01-15T10:30:00Z",
      "isRead": false
    }
  ],
  "count": 3,
  "unreadCount": 2
}
```

## Key Decisions

### Decision: CLI Framework

**Context:** Need to parse nested commands with flags

**Options:**
1. **Commander.js** - Most popular, well-documented, native subcommand support
2. **yargs** - Powerful but heavier, more complex API
3. **cac** - Lightweight alternative, less ecosystem
4. **Manual (process.argv)** - No dependencies but tedious

**Chosen:** Commander.js
- Native subcommand support matches our `category action` pattern
- TypeScript support is good
- Well-maintained, large community

### Decision: Output Formatting

**Context:** Primary use is Claude Bash tool, but humans also use it

**Options:**
1. **JSON only** - Simple, consistent with MCP
2. **Human-readable only** - Better terminal UX, harder to parse
3. **Human default + --json flag** - Best of both worlds

**Chosen:** Human default + --json flag
- Claude can use `--json` for reliable parsing
- Humans get readable output by default
- Aligns with common CLI conventions (gh, aws cli, etc.)

### Decision: Category Naming

**Context:** Should CLI use internal names (onedrive_list) or user-friendly names (files list)?

**Options:**
1. **Internal names** - `myoffice onedrive_list` - Matches MCP tools exactly
2. **User-friendly** - `myoffice files list` - More intuitive

**Chosen:** User-friendly names
- `files` instead of `onedrive` (more intuitive)
- `sharepoint` kept as-is (well-known term)
- Mapping table handles translation

### Decision: Login Integration

**Context:** How should authentication work in CLI?

**Options:**
1. **Separate script** - Keep `npm run login` as-is
2. **Integrated command** - `myoffice login` calls same flow
3. **Auto-prompt** - Auto-start login if not authenticated

**Chosen:** Integrated command (`myoffice login`)
- Single entry point for users
- Reuse existing device code flow
- Other commands show error with login instructions if not authenticated

## Error Handling

| Error Type | Human Output | JSON Output | Exit Code |
|------------|--------------|-------------|-----------|
| Not authenticated | "Not authenticated. Run: myoffice login" | `{"error": "...", "code": "AUTH_REQUIRED"}` | 1 |
| Invalid command | "Unknown command. See: myoffice --help" | `{"error": "...", "code": "INVALID_COMMAND"}` | 1 |
| Missing required arg | "Missing required: --id" | `{"error": "...", "code": "MISSING_ARG"}` | 1 |
| API error | "API error: <message>" | `{"error": "...", "code": "API_ERROR"}` | 1 |
| Network error | "Network error: <message>" | `{"error": "...", "code": "NETWORK_ERROR"}` | 1 |

All errors go to stderr, success output to stdout.

## Distribution & Installation

### Installation Method

Distributed via **Awolve Handbook** - install directly from GitHub:

```bash
# Install globally from GitHub
npm install -g github:awolve/ops-myoffice

# Or clone and link for development
git clone https://github.com/awolve/ops-myoffice.git
cd ops-myoffice
npm install && npm run build
npm link
```

### Requirements

Users need:
- Node.js 18+ installed
- Azure AD app registration (Awolve shared or personal)
- Environment variable: `M365_CLIENT_ID`

### Post-Installation Setup

```bash
# 1. Set up Azure AD credentials (see Awolve Handbook for shared credentials)
export M365_CLIENT_ID="your-client-id"
# Optional: export M365_TENANT_ID="your-tenant-id"

# 2. Authenticate
myoffice login

# 3. Start using
myoffice mail list
```

### Awolve Handbook Entry

Add documentation to the handbook covering:
- What the tool does
- Installation command (`npm install -g github:awolve/ops-personal-m365-mcp`)
- Azure AD app setup or shared credentials
- Common usage examples
- Link to GitHub repo for issues/updates

## Testing Strategy

**Manual testing checklist:**
1. Install globally: `npm link`
2. Test without auth: should prompt login
3. Test login flow: device code works
4. Test each category's basic commands
5. Test `--json` flag produces valid JSON
6. Test `--help` on all levels
7. Test error cases (invalid args, network errors)

**Future: Automated tests**
- Unit tests for formatter
- Unit tests for arg parsing
- Integration tests mocking Graph API

## File Structure (Final)

```
src/
├── index.ts              # MCP server entry (modified)
├── cli.ts                # CLI entry (new)
├── core/
│   └── handler.ts        # Shared tool dispatch (new, extracted)
├── cli/
│   ├── commands/         # Command definitions (new)
│   │   ├── mail.ts
│   │   ├── calendar.ts
│   │   ├── tasks.ts
│   │   ├── files.ts
│   │   ├── sharepoint.ts
│   │   ├── contacts.ts
│   │   ├── teams.ts
│   │   ├── chats.ts
│   │   └── planner.ts
│   └── formatter.ts      # Output formatting (new)
├── tools/                # Unchanged
├── auth/                 # Unchanged
└── utils/                # Unchanged

skills/
└── awolve-office.md      # Claude Code skill for MyOffice operations (new)
```

## Claude Code Skill

A skill file that helps Claude Code users interact with their Awolve/work Microsoft 365 account via the `myoffice` CLI.

### Skill: `awolve-office.md`

**Purpose:** Guide Claude to use the `myoffice` CLI for personal Microsoft 365 operations (the user's own mail, calendar, files, etc.), avoiding confusion with:
1. Personal Gmail/Google Calendar
2. M365 MCP (the separate admin MCP for tenant-level operations)

**Important distinction:**
- **MyOffice CLI** = personal user operations (your own inbox, calendar, OneDrive, Teams chats)
- **M365 MCP** = admin operations (managing users, groups, policies, tenant settings)

**Trigger conditions:**
- User asks about their own work email or Outlook
- User asks about their own work calendar
- User asks about their Teams channels/chats
- User asks about their SharePoint or OneDrive files
- User asks about their Planner tasks
- User mentions "Awolve" in context of personal email/calendar/files
- User asks to check/send their own messages

**Key behaviors:**
- Always use `myoffice` CLI with `--json` flag for parsing
- Distinguish personal Microsoft 365 (MyOffice) from admin operations (M365 MCP)
- Distinguish work (Microsoft 365/Outlook) from personal accounts (Gmail/Google Calendar)
- Check if CLI is installed before using (`which myoffice`)
- Check auth status before operations (`myoffice status --json`)
- Guide user through installation and authentication if needed
- Format results in human-readable way after parsing JSON

**Installation & Setup guidance the skill should provide:**

1. **Check if installed:** `which myoffice`
2. **Install if missing:**
   ```bash
   npm install -g github:awolve/ops-myoffice
   ```
3. **Set environment variable:** User needs `M365_CLIENT_ID` set (refer to Awolve Handbook for shared credentials)
4. **Authenticate:** `myoffice login` - initiates device code flow
5. **Verify:** `myoffice status` - confirms authentication

**Error handling:**
- "command not found: myoffice" → Guide through installation
- "Not authenticated" error → Guide to run `myoffice login`
- "M365_CLIENT_ID not set" → Explain env var setup, point to handbook
- Auth token expired → Suggest `myoffice login` to re-authenticate

**Example interactions:**
- "Check my work email" → First check `myoffice status --json`, then `myoffice mail list --json`
- "What meetings do I have today?" (work context) → `myoffice calendar list --json`
- "Send a message to the team" → `myoffice teams post` or `myoffice chats send`
- "What's on my Planner?" → `myoffice planner plans` then `myoffice planner tasks`
- "How do I set up myoffice?" → Walk through installation and authentication steps
