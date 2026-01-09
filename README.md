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
npm install      # Install dependencies
npm run dev      # Run with tsx (no build needed)
npm run build    # Compile TypeScript to dist/
npm run login    # Authenticate with Microsoft
```

### Project Structure

```
src/
├── index.ts           # MCP server entry point
├── cli.ts             # CLI entry point (Commander.js)
├── core/
│   └── handler.ts     # Shared tool dispatch logic
├── cli/
│   └── formatter.ts   # Human-readable output formatting
├── auth/              # Authentication (device code flow, token cache)
├── tools/             # Microsoft Graph API implementations
└── utils/             # Graph client, version helper
```

### Testing Locally

```bash
# Test CLI directly (no build needed)
npx tsx src/cli.ts mail list

# Or build first then test
npm run build
node dist/cli.js mail list

# Test MCP server
npm run dev
```

## Deploying to npm

The package is published to npm as `awolve-myoffice-cli`.

### First-Time Setup

1. Create npm account at https://www.npmjs.com/signup
2. Login from terminal:
   ```bash
   npm login
   ```

### Publishing a New Version

```bash
# 1. Make sure you're on main branch with clean working directory
git checkout main
git pull
git status  # Should be clean

# 2. Run tests / verify everything works
npm run build
node dist/cli.js --help

# 3. Bump version (choose one)
npm version patch  # 1.1.0 -> 1.1.1 (bug fixes)
npm version minor  # 1.1.0 -> 1.2.0 (new features)
npm version major  # 1.1.0 -> 2.0.0 (breaking changes)

# 4. Publish to npm (builds automatically via prepublishOnly)
npm publish --access public

# 5. Push version commit and tag to GitHub
git push && git push --tags
```

### What Gets Published

The `files` field in package.json controls what's included:
- `dist/**/*` - Compiled JavaScript
- `README.md` - Documentation

Source code (`src/`), specs, and dev files are NOT published.

### Verifying Publication

```bash
# Check package on npm
npm view awolve-myoffice-cli

# Test fresh install
npm install -g awolve-myoffice-cli
myoffice --version
```

### Important: Version Immutability

npm does not allow republishing the same version. Once `1.1.0` is published, that version number is permanently taken. If you need to fix something:

1. Bump to a new version (`npm version patch` → `1.1.1`)
2. Publish the new version

There is no way to "update" an existing version.

### Troubleshooting

- **"You must be logged in"** - Run `npm login`
- **"Package name already exists"** - Name is taken, choose another
- **"Cannot publish over existing version"** - Bump version with `npm version patch`
- **Build fails during publish** - Fix TypeScript errors first

## Security

- Uses delegated permissions only - users can only access their own data
- Tokens stored locally with restrictive permissions
- No client secrets required (public client)
- Destructive operations require explicit confirmation

## License

MIT
