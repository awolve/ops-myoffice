# Personal M365 MCP - Tasks

> Retroactive spec - documented after implementation

## Completed Tasks

### Phase 1: Project Setup
- [x] Create GitHub repo (awolve/ops-personal-m365-mcp)
- [x] Initialize package.json with dependencies
- [x] Configure TypeScript (tsconfig.json)
- [x] Create directory structure

### Phase 2: Authentication
- [x] Implement config and token storage (`src/auth/config.ts`)
- [x] Implement device code flow (`src/auth/device-code.ts`)
- [x] Implement token manager with refresh (`src/auth/token-manager.ts`)
- [x] Create standalone login script (`src/auth/login.ts`)

### Phase 3: Graph Client
- [x] Create HTTP client wrapper (`src/utils/graph-client.ts`)
- [x] Implement pagination support (graphList)
- [x] Add error handling

### Phase 4: Tools Implementation
- [x] Mail tools (list, read, search, send, delete)
- [x] Calendar tools (list, get, create, update, delete)
- [x] Tasks tools (list lists, list, create, update, complete, delete)
- [x] OneDrive tools (list, get, search, read, create folder)
- [x] Contacts tools (list, search, get)

### Phase 5: MCP Server
- [x] Create MCP server entry point
- [x] Define all tool schemas
- [x] Implement request routing
- [x] Add auth_status tool

## Remaining Tasks

### Phase 6: Documentation & Polish
- [ ] Create README.md with setup instructions
- [ ] Add .env.example file
- [ ] Add .gitignore
- [ ] Initial commit and push

### Phase 7: Azure AD Setup
- [ ] Document Azure AD app registration steps
- [ ] Create app registration in Awolve tenant
- [ ] Configure delegated permissions
- [ ] Test authentication flow

### Phase 8: Integration
- [ ] Add to awolve-general plugin
- [ ] Create `/setup-personal-m365` skill
- [ ] Test with Claude Code
- [ ] Document employee onboarding process

### Phase 9: Enhancements (Future)
- [ ] Add mail reply/forward tools
- [ ] Add file upload to OneDrive
- [ ] Add contact create/update tools
- [ ] Add calendar availability check
- [ ] Add meeting response (accept/decline)

## Test Checklist

- [ ] Device code login flow
- [ ] Token refresh after expiry
- [ ] List inbox emails
- [ ] Read specific email
- [ ] Search emails
- [ ] Send test email (to self)
- [ ] List calendar events
- [ ] Create calendar event
- [ ] List task lists
- [ ] Create and complete task
- [ ] List OneDrive files
- [ ] Read text file from OneDrive
- [ ] List contacts
- [ ] Search contacts
