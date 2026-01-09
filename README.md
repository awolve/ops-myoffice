# MyOffice MCP

A lightweight MCP (Model Context Protocol) server and CLI for personal Microsoft 365 access, designed for AI assistants like Claude Code.

Unlike admin-focused M365 tools, this uses **delegated authentication** - users authenticate as themselves and can only access their own data.

## Features

- **Email** - List, read, search, send, reply, delete, mark read/unread
- **Calendar** - List, create, update, delete events (with Teams meetings)
- **Tasks** - Manage Microsoft To Do lists and tasks
- **Planner** - Access plans, buckets, and tasks
- **OneDrive** - Browse, search, read files
- **SharePoint** - Access sites and document libraries
- **Teams** - List teams, channels, read/post messages
- **Chats** - 1:1 and group chats
- **Contacts** - List, search, create, update contacts

## Installation

### From npm (recommended)

```bash
npm install -g awolve-myoffice-cli
```

### From source

```bash
git clone https://github.com/awolve/ops-myoffice.git
cd ops-myoffice
npm install
npm run build
npm link
```

## Prerequisites

- Node.js 18+
- An Azure AD app registration with delegated permissions
- Microsoft 365 account

## Configuration

Add to your shell profile (`~/.zshrc` or `~/.bashrc`):

```bash
export M365_CLIENT_ID="your-app-client-id"
# Optional: export M365_TENANT_ID="your-tenant-id"
```

## Authentication

```bash
myoffice login
```

Opens a browser for device code authentication. Token is cached at `~/.config/myoffice-mcp/msal-cache.json`.

## CLI Usage

```bash
# Check status
myoffice status

# Email
myoffice mail list
myoffice mail list --unread
myoffice mail read <id>
myoffice mail send --to user@example.com --subject "Hi" --body "Hello"

# Calendar
myoffice calendar list
myoffice calendar list --start 2024-01-15 --end 2024-01-20

# Tasks (To Do)
myoffice tasks lists
myoffice tasks list
myoffice tasks create "Buy milk"

# Files (OneDrive)
myoffice files list
myoffice files search "report"

# Teams & Chats
myoffice teams list
myoffice chats list

# Planner
myoffice planner plans
myoffice planner tasks <planId>

# JSON output (for scripting)
myoffice mail list --json
```

Run `myoffice --help` for all commands.

## MCP Server Usage

Add to your Claude Code MCP settings:

```json
{
  "mcpServers": {
    "myoffice": {
      "command": "myoffice-mcp",
      "env": {
        "M365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

Or with full path if not installed globally:

```json
{
  "mcpServers": {
    "myoffice": {
      "command": "node",
      "args": ["/path/to/ops-myoffice/dist/index.js"],
      "env": {
        "M365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

## Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Click **New registration**
   - Name: `MyOffice MCP` (or your choice)
   - Supported account types: **Accounts in any organizational directory**
   - Redirect URI: Leave blank (uses device code flow)
3. Note the **Application (client) ID**
4. Go to **Authentication** > Enable **Allow public client flows** = Yes
5. Go to **API permissions** > Add Microsoft Graph delegated permissions:
   - `Mail.ReadWrite`, `Mail.Send`
   - `Calendars.ReadWrite`
   - `Tasks.ReadWrite`
   - `Files.ReadWrite`, `Sites.Read.All`
   - `Contacts.ReadWrite`
   - `Team.ReadBasic.All`, `Channel.ReadBasic.All`
   - `ChannelMessage.Read.All`, `ChannelMessage.Send`
   - `Chat.Create`, `Chat.ReadBasic`, `Chat.Read`, `ChatMessage.Send`
   - `Tasks.Read`, `Tasks.ReadWrite`, `Group.Read.All` (for Planner)
   - `User.Read`, `offline_access`

## Development

```bash
npm run dev      # Run with tsx
npm run build    # Compile TypeScript
npm run login    # Authenticate
```

## Publishing

To publish a new version to npm:

```bash
# 1. Update version in package.json
npm version patch  # or minor/major

# 2. Build and publish
npm publish --access public

# 3. Commit and push
git push && git push --tags
```

## Security

- Uses delegated permissions only - users can only access their own data
- Tokens stored locally with restrictive permissions
- No client secrets required (public client)
- Destructive operations require explicit confirmation

## License

MIT
