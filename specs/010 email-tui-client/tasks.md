# Email TUI Client - Tasks

## Overview
Estimated scope: Medium (6 phases)

## Tasks

- [ ] 1. Foundation
  - [ ] 1.1 Add dependencies to package.json
    - Files: `package.json`
    - Add: ink, react, @inkjs/ui, @types/react
  - [ ] 1.2 Create TUI entry point
    - Files: `src/tui/index.tsx`
    - Auth check, render App
  - [ ] 1.3 Create state management
    - Files: `src/tui/state/types.ts`, `src/tui/state/reducer.ts`, `src/tui/state/context.tsx`
    - State types, reducer, AppContext
  - [ ] 1.4 Create App shell with view router
    - Files: `src/tui/App.tsx`
    - AppProvider, view switching
  - [ ] 1.5 Add `tui` command to CLI
    - Files: `src/cli.ts`
    - New subcommand that launches TUI

- [ ] 2. Inbox View
  - [ ] 2.1 Create InboxView component
    - Files: `src/tui/views/InboxView.tsx`
    - Requirements: R1, R2, R5
  - [ ] 2.2 Create EmailListItem component
    - Files: `src/tui/components/EmailListItem.tsx`
    - Display: from, subject, date, read status
  - [ ] 2.3 Implement useMailData hook
    - Files: `src/tui/hooks/useMailData.ts`
    - Wrap mail.ts functions for TUI use
  - [ ] 2.4 Add keyboard navigation
    - Files: `src/tui/hooks/useKeyboard.ts`
    - Up/down arrow, Enter to open
  - [ ] 2.5 Create StatusBar and HelpBar
    - Files: `src/tui/components/StatusBar.tsx`, `src/tui/components/HelpBar.tsx`
    - Top: folder name, unread count. Bottom: shortcuts

- [ ] 3. Email View
  - [ ] 3.1 Create EmailView component
    - Files: `src/tui/views/EmailView.tsx`
    - Requirements: R3, R4
  - [ ] 3.2 Create EmailHeader component
    - Files: `src/tui/components/EmailHeader.tsx`
    - From, To, Subject, Date display
  - [ ] 3.3 Create ScrollableText component
    - Files: `src/tui/components/ScrollableText.tsx`
    - Arrow/PgUp/PgDn scrolling
  - [ ] 3.4 Create html-to-text utility
    - Files: `src/tui/utils/html-to-text.ts`
    - Strip HTML tags, decode entities
  - [ ] 3.5 Add back navigation
    - Update: `src/tui/views/EmailView.tsx`
    - Esc/Backspace returns to inbox

- [ ] 4. Compose View
  - [ ] 4.1 Create ComposeView component
    - Files: `src/tui/views/ComposeView.tsx`
    - Requirements: R6, R7, R8, R9, R10
  - [ ] 4.2 Create input fields (To, CC, Subject)
    - Files: `src/tui/components/TextInput.tsx`
    - Wrap @inkjs/ui TextInput
  - [ ] 4.3 Create BodyEditor component
    - Files: `src/tui/components/BodyEditor.tsx`
    - Multi-line text input with cursor
  - [ ] 4.4 Implement send functionality
    - Update: `src/tui/hooks/useMailData.ts`
    - Call sendMail/replyMail
  - [ ] 4.5 Implement reply/reply-all modes
    - Update: `src/tui/views/ComposeView.tsx`
    - Pre-fill from selected email

- [ ] 5. Search & Folders
  - [ ] 5.1 Create SearchView component
    - Files: `src/tui/views/SearchView.tsx`
    - Requirements: R13, R14, R15
  - [ ] 5.2 Create FolderSelectView component
    - Files: `src/tui/views/FolderSelectView.tsx`
    - Requirements: R11, R12
  - [ ] 5.3 Implement search functionality
    - Update: `src/tui/hooks/useMailData.ts`
    - Call searchMail, display results
  - [ ] 5.4 Implement folder switching
    - Update: `src/tui/state/reducer.ts`
    - Reload emails on folder change

- [ ] 6. Polish
  - [ ] 6.1 Add loading states
    - Files: `src/tui/components/LoadingSpinner.tsx`
    - Requirements: R16
  - [ ] 6.2 Add error handling
    - Files: `src/tui/components/ErrorModal.tsx`
    - Requirements: R17
  - [ ] 6.3 Add empty states
    - Files: `src/tui/components/EmptyState.tsx`
    - "No emails" message
  - [ ] 6.4 Handle terminal size
    - Files: `src/tui/hooks/useTerminalSize.ts`
    - Minimum 80x24 check
  - [ ] 6.5 Success feedback
    - Update views
    - Requirements: R18

## Implementation Order

Recommended sequence:
1. **Foundation first** - Get basic rendering working before adding features
2. **Inbox View** - Core navigation and data fetching
3. **Email View** - Read-only operations before write operations
4. **Compose View** - Write operations after read is solid
5. **Search & Folders** - Secondary features
6. **Polish** - Error handling and edge cases last

## Dependencies

New npm packages to install:
- `ink` (^5.0.1) - React renderer for CLI
- `react` (^18.3.1) - Required by Ink
- `@inkjs/ui` (^2.0.0) - TextInput, Select, Spinner
- `@types/react` (^18.3.0) - TypeScript types (dev)
