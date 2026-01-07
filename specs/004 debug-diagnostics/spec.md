# Debug Diagnostics

> Retroactive spec - documented after implementation

## Summary

Added comprehensive debugging tools to help diagnose MCP server issues, particularly authentication and Graph API connectivity problems.

## Why

When the MCP server hangs or fails silently, users have no visibility into what's wrong. Need diagnostic tools that surface errors with timeouts so debugging doesn't require reading source code.

## Changes Made

### Commits (27b66d1 - present)

- `src/index.ts`: Added `debug_info` tool with:
  - Server info (version, pid, uptime, start time)
  - Environment variable status (with placeholder detection)
  - Authentication status
  - Graph API connectivity test with 10s timeout
- `src/index.ts`: Added startup logging to stderr
- `src/utils/version.ts`: Dynamic version reading (related: spec 003)

### debug_info Tool Output

```json
{
  "server": {
    "name": "ops-personal-m365-mcp",
    "version": "0.2.1",
    "nodeVersion": "v20.x.x",
    "platform": "darwin",
    "pid": 12345,
    "startedAt": "2024-01-07T10:00:00.000Z",
    "uptime": "1h 30m 45s"
  },
  "environment": {
    "M365_CLIENT_ID": "SET (a1b2c3d4...)",
    "M365_TENANT_ID": "common"
  },
  "auth": {
    "authenticated": true,
    "user": { ... },
    "tokenCachePath": "~/.config/ops-personal-m365-mcp/token.json"
  },
  "graphApiTest": {
    "status": "OK",
    "responseTimeMs": 234,
    "user": "Björn Allvin"
  },
  "tools": 35
}
```

## Key Features

- **Timeout protection**: Graph API test has 10 second timeout to prevent hanging
- **Placeholder detection**: Catches unresolved `${VAR}` in environment variables
- **Uptime tracking**: Shows when server started and how long it's been running
- **Error surfacing**: Returns actual error messages from Graph API failures

## Tasks

- [x] Add debug_info tool
- [x] Add startup logging
- [x] Add uptime tracking
- [x] Detect unresolved environment variable placeholders
- [x] Add Graph API connectivity test with timeout
- [x] Return error messages on failure

## Notes

- Logs go to stderr (console.error) so they don't interfere with MCP stdio protocol
- Version 0.2.1 includes the Graph API test feature
