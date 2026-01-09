---
description: Guide for accessing personal Microsoft 365 data (work email, calendar, OneDrive, Teams, Planner) via the MyOffice CLI. Use when user asks about their own work/Awolve email, calendar, files, Teams channels/chats, or Planner tasks. This is for PERSONAL user operations, not admin operations (use M365 MCP for admin tasks like managing users/groups/policies).
trigger:
  - user asks about their work email or Outlook
  - user asks about their work calendar
  - user asks about Teams channels or chats
  - user asks about SharePoint or OneDrive files
  - user asks about Planner tasks
  - user mentions "Awolve" with email/calendar/files context
  - user asks to check, read, or send work messages
---

# MyOffice CLI - Personal Microsoft 365 Access

Use the `myoffice` CLI to access the user's personal Microsoft 365 data: their own inbox, calendar, OneDrive, SharePoint, Teams, and Planner.

## Important Distinctions

- **MyOffice CLI** = Personal user operations (user's own inbox, calendar, OneDrive, Teams chats)
- **M365 MCP** = Admin operations (managing users, groups, policies, tenant settings)
- **Gmail/Google Calendar** = Personal Google account (completely separate)

This skill is for **personal Microsoft 365 operations only**.

## Before Using MyOffice

### 1. Check if Installed

```bash
which myoffice
```

If not installed, guide user through installation:

```bash
npm install -g github:awolve/ops-myoffice
```

### 2. Check Environment Variable

The CLI requires `M365_CLIENT_ID` for **every command** (including reading the token cache). Check if set:

```bash
echo $M365_CLIENT_ID
```

If not set, user **must** add to their shell profile (`~/.zshrc` or `~/.bashrc`):

```bash
echo 'export M365_CLIENT_ID="client-id-from-awolve-handbook"' >> ~/.zshrc
source ~/.zshrc
```

Refer user to Awolve Handbook for shared Azure AD credentials.

**Important:** Without this env var, even after successful login, commands will fail with "Not authenticated".

### 3. Check Authentication Status

```bash
myoffice status --json
```

If not authenticated, run:

```bash
myoffice login
```

This opens a browser for Microsoft login via device code flow.

## CLI Usage

Always use `--json` flag for reliable parsing:

```bash
myoffice <command> --json
```

### Mail Commands

```bash
myoffice mail list --json                    # List inbox
myoffice mail list --unread --json           # Unread only
myoffice mail list --folder sentitems --json # Sent folder
myoffice mail read <id> --json               # Read specific email
myoffice mail search <query> --json          # Search emails
myoffice mail send --to addr --subject subj --body body --json
myoffice mail reply <id> --body "text" --json
myoffice mail delete <id> --json
myoffice mail mark <id> --json               # Mark as read
myoffice mail mark <id> --unread --json      # Mark as unread
```

### Calendar Commands

```bash
myoffice calendar list --json                 # Next 7 days
myoffice calendar list --start 2024-01-15 --end 2024-01-20 --json
myoffice calendar get <id> --json
myoffice calendar create --subject "Meeting" --start "2024-01-15T10:00:00" --end "2024-01-15T11:00:00" --json
myoffice calendar update <id> --subject "New title" --json
myoffice calendar delete <id> --json
```

### Tasks (Microsoft To Do)

```bash
myoffice tasks lists --json                   # List task lists
myoffice tasks list --json                    # Tasks from default list
myoffice tasks list --list <listId> --json    # Tasks from specific list
myoffice tasks create <title> --json
myoffice tasks create <title> --due 2024-01-20 --json
myoffice tasks update <id> --title "new" --json
myoffice tasks complete <id> --json
myoffice tasks delete <id> --json
```

### Files (OneDrive)

```bash
myoffice files list --json                    # Root folder
myoffice files list /Documents --json         # Specific folder
myoffice files get /path/to/file.txt --json   # File metadata
myoffice files read /path/to/file.txt --json  # Read text content
myoffice files search <query> --json
myoffice files mkdir <name> --json
myoffice files shared --json                  # Files shared with me
```

### SharePoint

```bash
myoffice sharepoint sites --json              # List sites
myoffice sharepoint sites --search query --json
myoffice sharepoint site <siteId> --json      # Site details
myoffice sharepoint drives <siteId> --json    # Document libraries
myoffice sharepoint files <driveId> --json    # Files in library
myoffice sharepoint files <driveId> /path --json
myoffice sharepoint file <driveId> /path --json
myoffice sharepoint read <driveId> /path --json
myoffice sharepoint search <driveId> <query> --json
```

### Teams

```bash
myoffice teams list --json                    # List teams
myoffice teams channels <teamId> --json       # List channels
myoffice teams messages <teamId> <channelId> --json
myoffice teams post <teamId> <channelId> "message" --json
```

### Chats (1:1 and Group)

```bash
myoffice chats list --json                    # List chats
myoffice chats messages <chatId> --json       # Chat history
myoffice chats send <chatId> "message" --json
myoffice chats create user@example.com --json # Start new chat
myoffice chats create user@example.com "Hi!" --json
```

### Planner

```bash
myoffice planner plans --json                 # List plans
myoffice planner plan <planId> --json         # Plan details
myoffice planner buckets <planId> --json      # List buckets
myoffice planner bucket-create <planId> "name" --json
myoffice planner bucket-update <id> "new name" --json
myoffice planner bucket-delete <id> --json
myoffice planner tasks <planId> --json        # List tasks
myoffice planner tasks <planId> --bucket <id> --json
myoffice planner task <taskId> --json         # Task details
myoffice planner task-create <planId> "title" --json
myoffice planner task-create <planId> "title" --bucket <id> --due 2024-01-20 --json
myoffice planner task-update <id> --title "new" --progress completed --json
myoffice planner task-delete <id> --json
myoffice planner task-details <id> --json     # Description, checklist
myoffice planner task-details-update <id> --description "text" --json
```

### System Commands

```bash
myoffice status --json                        # Auth status
myoffice debug --json                         # Full debug info
myoffice login                                # Authenticate (interactive)
```

## Error Handling

| Error | Solution |
|-------|----------|
| `command not found: myoffice` | Install: `npm install -g github:awolve/ops-myoffice` |
| `Not authenticated` (after login worked) | `M365_CLIENT_ID` not set - add to `~/.zshrc` |
| `Not authenticated` (never logged in) | Run: `myoffice login` |
| `M365_CLIENT_ID environment variable is required` | Set env var in shell profile (see Awolve Handbook) |
| API errors | Check `myoffice debug --json` for details |

## Workflow Example

```bash
# 1. Check status
myoffice status --json

# 2. If not authenticated
myoffice login

# 3. List recent emails
myoffice mail list --limit 10 --json

# 4. Read specific email
myoffice mail read "AAMkAG..." --json

# 5. Check calendar
myoffice calendar list --json
```
