# Personal M365 MCP - Design

> Retroactive spec - documented after implementation

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        Claude Code                               │
└────────────────────────────┬────────────────────────────────────┘
                             │ stdio
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│                   ops-personal-m365-mcp                          │
│  ┌──────────────────────────────────────────────────────────┐   │
│  │                     MCP Server                            │   │
│  │  - Tool definitions                                       │   │
│  │  - Request routing                                        │   │
│  └────────────────────────────┬─────────────────────────────┘   │
│                               │                                  │
│  ┌────────────────────────────┴─────────────────────────────┐   │
│  │                      Tools Layer                          │   │
│  │  ┌─────┐ ┌─────────┐ ┌─────┐ ┌────────┐ ┌────────┐       │   │
│  │  │Mail │ │Calendar │ │Tasks│ │OneDrive│ │Contacts│       │   │
│  │  └──┬──┘ └────┬────┘ └──┬──┘ └───┬────┘ └───┬────┘       │   │
│  └─────┼─────────┼─────────┼────────┼──────────┼────────────┘   │
│        │         │         │        │          │                 │
│  ┌─────┴─────────┴─────────┴────────┴──────────┴────────────┐   │
│  │                   Graph Client                            │   │
│  │  - HTTP requests to Graph API                             │   │
│  │  - Pagination handling                                    │   │
│  └────────────────────────────┬─────────────────────────────┘   │
│                               │                                  │
│  ┌────────────────────────────┴─────────────────────────────┐   │
│  │                   Auth Layer                              │   │
│  │  - Device code flow                                       │   │
│  │  - Token management                                       │   │
│  │  - Token caching                                          │   │
│  └──────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
                             │
                             ▼
                 ┌───────────────────────┐
                 │  Microsoft Graph API  │
                 │  graph.microsoft.com  │
                 └───────────────────────┘
```

## Directory Structure

```
ops-personal-m365-mcp/
├── package.json
├── tsconfig.json
├── src/
│   ├── index.ts              # MCP server entry point
│   ├── auth/
│   │   ├── index.ts          # Auth exports
│   │   ├── config.ts         # Configuration & token storage
│   │   ├── device-code.ts    # Device code flow implementation
│   │   ├── token-manager.ts  # Token acquisition & refresh
│   │   └── login.ts          # Standalone login script
│   ├── tools/
│   │   ├── index.ts          # Tool exports
│   │   ├── mail.ts           # Email operations
│   │   ├── calendar.ts       # Calendar operations
│   │   ├── tasks.ts          # To Do tasks operations
│   │   ├── onedrive.ts       # File operations
│   │   └── contacts.ts       # Contact operations
│   └── utils/
│       └── graph-client.ts   # Graph API HTTP client
└── specs/
    └── 001-personal-m365-mcp/
```

## Key Design Decisions

### 1. Device Code Flow for Auth

**Choice:** Device code flow over other OAuth methods.

**Rationale:**
- Works well with CLI tools (no browser redirect needed)
- User-friendly: "Go to URL, enter code"
- No localhost server required
- Works in remote/headless environments

### 2. MSAL.js for Token Management

**Choice:** Use `@azure/msal-node` instead of raw OAuth.

**Rationale:**
- Handles token refresh automatically
- Proper cache management
- Well-tested by Microsoft
- Supports all Azure AD scenarios

### 3. Zod for Schema Validation

**Choice:** Use Zod for input validation.

**Rationale:**
- Runtime type checking
- Self-documenting schemas
- Good TypeScript integration
- Used by MCP SDK

### 4. Flat Tool Structure

**Choice:** All tools in single namespace (mail_*, calendar_*, etc.)

**Rationale:**
- Simpler than nested resources
- Clear naming convention
- Easier to discover
- Matches other MCP patterns

### 5. Local Token Storage

**Choice:** Store tokens in `~/.config/ops-personal-m365-mcp/token.json`

**Rationale:**
- XDG-compliant location
- User-specific
- File permissions (0600) protect credentials
- Easy to clear/reset

## Tool Naming Convention

```
{domain}_{action}

Examples:
- mail_list
- mail_send
- calendar_create
- tasks_complete
```

## Error Handling

1. **Graph API errors:** Parsed and returned with meaningful messages
2. **Auth errors:** Trigger re-authentication via device code
3. **Validation errors:** Zod provides clear error messages
4. **All errors:** Returned via MCP error response format

## Security Considerations

1. **No client secrets:** Public client, so no secrets to protect
2. **Token file permissions:** 0600 (owner read/write only)
3. **Scoped permissions:** Only delegated permissions, no admin access
4. **Tool descriptions:** Clearly mark destructive operations for AI confirmation
