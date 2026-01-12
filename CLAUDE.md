# MyOffice MCP Server

## Overview

A lightweight MCP (Model Context Protocol) server providing personal Microsoft 365 access via Microsoft Graph API. Uses delegated authentication - users authenticate as themselves and can only access their own data.

**Two entry points:**
- `myoffice-mcp` - MCP server for AI assistants (Claude Code MCP integration)
- `myoffice` - CLI for terminal use (Claude Code can call via Bash)

## Architecture

```
src/
├── index.ts           # MCP server entry point, tool definitions
├── cli.ts             # CLI entry point (Commander.js)
├── core/
│   └── handler.ts     # Shared tool dispatch logic
├── cli/
│   └── formatter.ts   # Human-readable output formatting
├── auth/
│   ├── config.ts      # Azure AD client config and scopes
│   ├── device-code.ts # Device code authentication flow
│   ├── token-manager.ts # Token caching and refresh
│   ├── login.ts       # CLI login script
│   └── index.ts       # Auth exports
├── tools/             # Graph API tool implementations
│   ├── mail.ts        # Email operations
│   ├── calendar.ts    # Calendar events
│   ├── tasks.ts       # Microsoft To Do
│   ├── planner.ts     # Microsoft Planner (plans, buckets, tasks)
│   ├── onedrive.ts    # OneDrive files + shared files
│   ├── sharepoint.ts  # SharePoint sites and document libraries
│   ├── contacts.ts    # Contacts
│   ├── teams.ts       # Teams channels and messages
│   ├── chats.ts       # 1:1 and group chats
│   └── index.ts       # Tool exports
└── utils/
    ├── graph-client.ts # Authenticated Graph API client
    └── version.ts      # Dynamic version from package.json
```

## Key Files

- `src/index.ts` - MCP server entry point, tool definitions (TOOLS array)
- `src/cli.ts` - CLI entry point with all commands
- `src/core/handler.ts` - Shared tool dispatch (used by both MCP and CLI)
- `src/cli/formatter.ts` - Human-readable output formatting
- `src/auth/config.ts` - Azure AD scopes (add new permissions here)
- `src/utils/graph-client.ts` - All Graph API calls go through this
- `src/tools/*.ts` - Individual tool implementations

## Authentication

- Device code flow for initial authentication (`npm run login`)
- Tokens cached at `~/.config/myoffice-mcp/token.json`
- Requires Azure AD app registration with delegated permissions
- Environment variables: `M365_CLIENT_ID` (required), `M365_TENANT_ID` (optional, defaults to "common")

## Adding New Tools

1. Create or update tool module in `src/tools/`
2. Define Zod schema for input validation
3. Implement function that calls Graph API via `graphClient.fetch()`
4. Add tool definition to `TOOLS` array in `src/index.ts`
5. Add case to switch statement in CallToolRequestSchema handler
6. If new permissions needed, add scopes to `src/auth/config.ts`

## Development

```bash
npm run dev    # Run with tsx (TypeScript directly)
npm run build  # Compile to dist/
npm run login  # Authenticate with Microsoft
```

## CLI Usage

The `myoffice` CLI provides terminal access to all Microsoft 365 tools.

### Installation & Authentication

```bash
npm install -g awolve-myoffice-cli

# Add to ~/.zshrc (required for all commands)
export M365_CLIENT_ID="your-azure-app-client-id"

# Authenticate (opens browser for device code flow)
myoffice login
```

**Note:** `M365_CLIENT_ID` must be set for every command. Add it to your shell profile for convenience.

### Commands

| Command | Description |
|---------|-------------|
| `myoffice login` | Authenticate with Microsoft 365 |
| `myoffice status` | Check authentication status |
| `myoffice debug` | Show server and auth info |

**Mail:**
- `myoffice mail list [--folder <name>] [--unread]` - List emails
- `myoffice mail read <id>` - Read email
- `myoffice mail search <query>` - Search emails
- `myoffice mail send --to <addr> --subject <subj> --body <body>` - Send email
- `myoffice mail reply <id> --body <body> [--all]` - Reply to email
- `myoffice mail delete <id>` - Delete email
- `myoffice mail mark <id> [--unread]` - Mark as read/unread

**Calendar:**
- `myoffice calendar list [--start <date>] [--end <date>]` - List events
- `myoffice calendar get <id>` - Get event details
- `myoffice calendar create --subject <subj> --start <dt> --end <dt>` - Create event
- `myoffice calendar update <id> [--subject <s>] [--start <dt>]` - Update event
- `myoffice calendar delete <id>` - Delete event

**Tasks (Microsoft To Do):**
- `myoffice tasks lists` - List task lists
- `myoffice tasks list [--list <id>] [--completed]` - List tasks
- `myoffice tasks create <title> [--list <id>] [--due <date>]` - Create task
- `myoffice tasks update <id> [--title <t>] [--due <date>]` - Update task
- `myoffice tasks complete <id>` - Mark task complete
- `myoffice tasks delete <id>` - Delete task

**Files (OneDrive):**
- `myoffice files list [path]` - List files
- `myoffice files get <path>` - Get file metadata
- `myoffice files search <query>` - Search files
- `myoffice files read <path>` - Read text file content
- `myoffice files mkdir <name> [--parent <path>]` - Create folder
- `myoffice files shared` - List files shared with me
- `myoffice files upload --file <path> [--dest <path>]` - Upload local file (any size)

**SharePoint:**
- `myoffice sharepoint sites [--search <query>]` - List sites
- `myoffice sharepoint site <id>` - Get site details
- `myoffice sharepoint drives <siteId>` - List document libraries
- `myoffice sharepoint files <driveId> [path]` - List files
- `myoffice sharepoint file <driveId> <path>` - Get file metadata
- `myoffice sharepoint read <driveId> <path>` - Read file content
- `myoffice sharepoint search <driveId> <query>` - Search files

**Contacts:**
- `myoffice contacts list` - List contacts
- `myoffice contacts search <query>` - Search contacts
- `myoffice contacts get <id>` - Get contact details

**Teams:**
- `myoffice teams list` - List teams
- `myoffice teams channels <teamId>` - List channels
- `myoffice teams messages <teamId> <channelId>` - List channel messages
- `myoffice teams post <teamId> <channelId> <message>` - Post message

**Chats:**
- `myoffice chats list` - List chats
- `myoffice chats messages <chatId>` - List chat messages
- `myoffice chats send <chatId> <message>` - Send message
- `myoffice chats create <email> [message]` - Create/start chat

**Planner:**
- `myoffice planner plans [--group <id>]` - List plans
- `myoffice planner plan <id>` - Get plan details
- `myoffice planner buckets <planId>` - List buckets
- `myoffice planner bucket-create <planId> <name>` - Create bucket
- `myoffice planner bucket-update <id> <name>` - Update bucket
- `myoffice planner bucket-delete <id>` - Delete bucket
- `myoffice planner tasks <planId> [--bucket <id>]` - List tasks
- `myoffice planner task <id>` - Get task details
- `myoffice planner task-create <planId> <title> [--bucket <id>]` - Create task
- `myoffice planner task-update <id> [--title <t>] [--progress <p>]` - Update task
- `myoffice planner task-delete <id>` - Delete task
- `myoffice planner task-details <id>` - Get task details (description, checklist, attachments)
- `myoffice planner task-details-update <id> [--description <d>]` - Update details
- `myoffice planner attach --id <taskId> --url <url> [--alias <name>]` - Add link/attachment to task
- `myoffice planner detach --id <taskId> --url <url>` - Remove attachment from task
- `myoffice planner upload --id <taskId> --file <path> [--alias <name>]` - Upload file and attach to task

### Output Format

By default, CLI outputs human-readable tables. Use `--json` flag for JSON output:

```bash
myoffice mail list --json
myoffice --json calendar list
```

## Testing

No automated tests currently. Test manually by running the MCP server and calling tools via Claude Code.

Use `debug_info` tool to check server status, version, and auth state.

## Specs

Feature specs are in `specs/` directory:
- `001-personal-m365-mcp/` - Initial implementation (requirements, design, tasks)
- `002 shared-onedrive-sharepoint/` - SharePoint and shared files access
- `003 version-management/` - Dynamic version from package.json
- `007 cli-interface/` - CLI for terminal access
- `008 teams-integration/` - Teams and chats support
- `009-planner-integration/` - Microsoft Planner support (plans, buckets, tasks)

## Common Tasks

### Before pushing new features
1. Bump the version in `package.json`
2. Check `specs/` for a spec that describes the current changes
3. If no relevant spec exists, run `/retro-spec` to document the work
4. Commit and push

### Add new Graph API permission
Add scope to `scopes` array in `src/auth/config.ts`. User must re-authenticate.

### Debug MCP server issues
Call the `debug_info` tool to see server info, environment variables, and auth status.

## Planner Integration

Microsoft Planner provides team-oriented task management. Key concepts:

- **Plans** - Read-only. Plans belong to M365 Groups, users can only see plans in groups they're members of.
- **Buckets** - Columns within a plan. Full CRUD supported.
- **Tasks** - Items within buckets. Full CRUD with assignments, due dates, priority, progress.
- **Task Details** - Extended info: description, checklist items, references (attachments).
- **References/Attachments** - Links to files or URLs. Planner doesn't store files directly; it stores references (URLs) to files in OneDrive, SharePoint, or external sites.

**Important notes:**
- All updates/deletes require ETags (handled internally - no user action needed)
- Task assignments accept email addresses (resolved to user IDs automatically)
- Progress values: `notStarted`, `inProgress`, `completed`
- Priority values: `urgent`, `important`, `medium`, `low`
- Plans cannot be created via MCP (would require M365 Group creation)
- Attachments are stored as "references" (URLs). The `planner upload` command uploads to the plan's SharePoint site (accessible to all plan members) at `Planner Attachments/<Plan Name>/<filename>`. File type is auto-detected from the URL.
