# MyOffice MCP

A lightweight MCP (Model Context Protocol) server for personal Microsoft 365 access, designed for AI assistants like Claude Code.

Unlike admin-focused M365 tools, this uses **delegated authentication** - users authenticate as themselves and can only access their own data.

## Features

- **Email** - List, read, search, send, delete
- **Calendar** - List, create, update, delete events (with Teams meetings)
- **Tasks** - Manage Microsoft To Do lists and tasks
- **OneDrive** - Browse, search, read files
- **Contacts** - List and search contacts

## Prerequisites

- Node.js 18+
- An Azure AD app registration with delegated permissions (see setup below)
- Microsoft 365 account

## Installation

```bash
git clone https://github.com/awolve/ops-myoffice.git
cd ops-myoffice
npm install
npm run build
```

## Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Click **New registration**
   - Name: `Personal M365 MCP` (or your choice)
   - Supported account types: **Accounts in any organizational directory** (for multi-tenant)
   - Redirect URI: Leave blank (we use device code flow)
3. After creation, note the **Application (client) ID**
4. Go to **Authentication** > Advanced settings
   - Enable **Allow public client flows** = Yes
5. Go to **API permissions** > Add a permission > Microsoft Graph > Delegated permissions
   - Add these permissions:
     - `Mail.ReadWrite`
     - `Mail.Send`
     - `Calendars.ReadWrite`
     - `Tasks.ReadWrite`
     - `Files.ReadWrite`
     - `Contacts.Read`
     - `User.Read`
     - `offline_access`
6. Click **Grant admin consent** (optional - users can consent themselves)

## Configuration

Set environment variables:

```bash
export M365_CLIENT_ID="your-app-client-id"
export M365_TENANT_ID="common"  # or your tenant ID for single-tenant
```

Or create a `.env` file:

```
M365_CLIENT_ID=your-app-client-id
M365_TENANT_ID=common
```

## First-Time Authentication

Run the login script to authenticate:

```bash
npm run login
```

You'll see:
```
AUTHENTICATION REQUIRED
========================================
To sign in, use a web browser to open the page
https://microsoft.com/devicelogin and enter the code XXXXXXXX
========================================
```

After signing in, your token is cached at `~/.config/myoffice-mcp/token.json`

## Usage with Claude Code

Add to your Claude Code MCP settings:

```json
{
  "mcpServers": {
    "myoffice-mcp": {
      "command": "node",
      "args": ["/path/to/ops-myoffice/dist/index.js"],
      "env": {
        "M365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

## Available Tools

### Email
| Tool | Description |
|------|-------------|
| `mail_list` | List emails from a folder |
| `mail_read` | Read a specific email |
| `mail_search` | Search emails |
| `mail_send` | Send an email |
| `mail_delete` | Delete an email |

### Calendar
| Tool | Description |
|------|-------------|
| `calendar_list` | List events in date range |
| `calendar_get` | Get event details |
| `calendar_create` | Create an event |
| `calendar_update` | Update an event |
| `calendar_delete` | Delete an event |

### Tasks
| Tool | Description |
|------|-------------|
| `tasks_list_lists` | List all task lists |
| `tasks_list` | List tasks from a list |
| `tasks_create` | Create a task |
| `tasks_update` | Update a task |
| `tasks_complete` | Mark task complete |
| `tasks_delete` | Delete a task |

### OneDrive
| Tool | Description |
|------|-------------|
| `onedrive_list` | List files/folders |
| `onedrive_get` | Get file metadata |
| `onedrive_search` | Search files |
| `onedrive_read` | Read text file content |
| `onedrive_create_folder` | Create a folder |

### Contacts
| Tool | Description |
|------|-------------|
| `contacts_list` | List contacts |
| `contacts_search` | Search contacts |
| `contacts_get` | Get contact details |

### Auth
| Tool | Description |
|------|-------------|
| `auth_status` | Check auth status |

## Security

- Uses delegated permissions only - users can only access their own data
- Tokens stored locally with restrictive permissions (0600)
- No client secrets required (public client)
- Destructive operations (send, delete) are clearly marked for AI confirmation

## Development

```bash
# Run in development mode
npm run dev

# Build
npm run build

# Run login flow
npm run login
```

## License

MIT
