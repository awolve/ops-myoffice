# MyOffice MCP Server

## Overview

A lightweight MCP (Model Context Protocol) server providing personal Microsoft 365 access via Microsoft Graph API. Uses delegated authentication - users authenticate as themselves and can only access their own data.

## Architecture

```
src/
├── index.ts           # MCP server entry point, tool definitions and routing
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

- `src/index.ts` - MCP server entry point, tool definitions (TOOLS array), and routing (switch statement)
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

## Testing

No automated tests currently. Test manually by running the MCP server and calling tools via Claude Code.

Use `debug_info` tool to check server status, version, and auth state.

## Specs

Feature specs are in `specs/` directory:
- `001-personal-m365-mcp/` - Initial implementation (requirements, design, tasks)
- `002 shared-onedrive-sharepoint/` - SharePoint and shared files access
- `003 version-management/` - Dynamic version from package.json
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
- **Task Details** - Extended info: description, checklist items.

**Important notes:**
- All updates/deletes require ETags (handled internally - no user action needed)
- Task assignments accept email addresses (resolved to user IDs automatically)
- Progress values: `notStarted`, `inProgress`, `completed`
- Priority values: `urgent`, `important`, `medium`, `low`
- Plans cannot be created via MCP (would require M365 Group creation)
