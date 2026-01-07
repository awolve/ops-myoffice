# Personal M365 MCP Server

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
│   ├── onedrive.ts    # OneDrive files + shared files
│   ├── sharepoint.ts  # SharePoint sites and document libraries
│   ├── contacts.ts    # Contacts
│   └── index.ts       # Tool exports
└── utils/
    ├── graph-client.ts # Authenticated Graph API client
    └── version.ts      # Dynamic version from package.json
```

## Key Files

- `src/index.ts:37-47` - Server registration with MCP SDK
- `src/index.ts:50-481` - Tool definitions (JSON Schema)
- `src/index.ts:489-675` - Tool call routing and execution
- `src/auth/config.ts` - Azure AD scopes (add new permissions here)
- `src/utils/graph-client.ts` - All Graph API calls go through this

## Authentication

- Device code flow for initial authentication (`npm run login`)
- Tokens cached at `~/.config/ops-personal-m365-mcp/token.json`
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

## Common Tasks

### Update version
Edit `version` in `package.json`. Server reads it dynamically.

**Important:** Always bump the version in `package.json` when making changes. This helps users verify they're running the updated code via `debug_info`.

### Add new Graph API permission
Add scope to `scopes` array in `src/auth/config.ts`. User must re-authenticate.

### Debug MCP server issues
Call the `debug_info` tool to see server info, environment variables, and auth status.
