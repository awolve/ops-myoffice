# Email TUI Client - Design

## Overview

An interactive TUI email client built with Ink (React for CLI). Provides full-screen view switching between inbox list, email reader, and compose views. Built on top of existing myoffice email tools.

## Architecture

```
┌─────────────────────────────────────────────────┐
│                 User Terminal                    │
└──────────────────────┬──────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────┐
│            src/tui/index.tsx                     │
│            (TUI Entry Point)                     │
└──────────────────────┬──────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────┐
│            src/tui/App.tsx                       │
│            - AppProvider (state context)         │
│            - View Router                         │
└──────────────────────┬──────────────────────────┘
                       │
        ┌──────────────┼──────────────┐
        │              │              │
┌───────▼────┐  ┌──────▼─────┐  ┌─────▼──────┐
│ InboxView  │  │ EmailView  │  │ComposeView │
└───────┬────┘  └──────┬─────┘  └─────┬──────┘
        │              │              │
        └──────────────┼──────────────┘
                       │
┌──────────────────────▼──────────────────────────┐
│            src/tools/mail.ts                     │
│            (Existing email tools)                │
└─────────────────────────────────────────────────┘
```

## Components

### Entry Point
- **Purpose:** Launch TUI, check auth, render App
- **Location:** `src/tui/index.tsx`
- **Interface:** `myoffice tui` command

### App (Root Component)
- **Purpose:** State management, view routing, global keyboard handling
- **Location:** `src/tui/App.tsx`
- **Interface:** Wraps all views in AppProvider context

### InboxView
- **Purpose:** Display email list with selection highlighting
- **Location:** `src/tui/views/InboxView.tsx`
- **Interface:** Arrow keys navigate, Enter opens email

### EmailView
- **Purpose:** Display full email content with scrolling
- **Location:** `src/tui/views/EmailView.tsx`
- **Interface:** Esc returns to inbox, r/R for reply

### ComposeView
- **Purpose:** New email and reply composition
- **Location:** `src/tui/views/ComposeView.tsx`
- **Interface:** Tab between fields, Ctrl+Enter sends

### SearchView
- **Purpose:** Search input and results display
- **Location:** `src/tui/views/SearchView.tsx`
- **Interface:** Enter executes search, Esc cancels

### FolderSelectView
- **Purpose:** Folder picker overlay
- **Location:** `src/tui/views/FolderSelectView.tsx`
- **Interface:** Arrow keys select, Enter confirms

## Data Models

```typescript
// State types
type View = 'inbox' | 'email' | 'compose' | 'search' | 'folder-select';
type ComposeMode = 'new' | 'reply' | 'reply-all';

interface Email {
  id: string;
  subject: string;
  from: string;
  fromName?: string;
  received: string;
  isRead: boolean;
  body?: string;
  to?: string[];
}

interface AppState {
  currentView: View;
  emails: Email[];
  selectedIndex: number;
  selectedEmail: Email | null;
  currentFolder: string;
  searchQuery: string;
  searchResults: Email[];
  compose: ComposeState | null;
  isLoading: boolean;
  error: { message: string; retry?: () => void } | null;
}

interface ComposeState {
  mode: ComposeMode;
  replyToId?: string;
  to: string[];
  cc: string[];
  subject: string;
  body: string;
}
```

## Key Decisions

### Decision: Entry Point Command
**Context:** How should users launch the TUI?
**Options:**
1. `myoffice tui` - Dedicated subcommand
2. `myoffice mail --interactive` - Flag on existing command
**Chosen:** `myoffice tui` - Cleaner separation, allows future expansion (calendar TUI, etc.)

### Decision: Data Layer Access
**Context:** How should TUI get email data?
**Options:**
1. Import mail.ts functions directly - Type-safe, efficient
2. Go through CLI handler - String-based dispatch, unnecessary overhead
**Chosen:** Direct imports from `src/tools/mail.ts` for type safety and efficiency

### Decision: HTML Email Handling
**Context:** Terminals can't render HTML emails
**Options:**
1. Use `html-to-text` npm package - Full-featured but heavy
2. Simple regex stripping - Lightweight, sufficient for triage
**Chosen:** Simple regex-based HTML stripping. Good enough for quick triage use case.

### Decision: Multi-line Text Input
**Context:** Ink's TextInput is single-line, email body needs multi-line
**Options:**
1. Use external package - More dependencies
2. Custom implementation - More control, tailored to needs
**Chosen:** Custom ScrollableText component with cursor management

## File Structure

```
src/tui/
├── index.tsx           # Entry point, auth check, render
├── App.tsx             # Root component, providers, router
├── state/
│   ├── types.ts        # State and action types
│   ├── reducer.ts      # State reducer
│   └── context.tsx     # AppContext, useApp hook
├── hooks/
│   ├── useMailData.ts  # Email operations wrapper
│   └── useKeyboard.ts  # Global keyboard handler
├── views/
│   ├── InboxView.tsx   # Email list
│   ├── EmailView.tsx   # Full email reader
│   ├── ComposeView.tsx # Compose/reply
│   ├── SearchView.tsx  # Search
│   └── FolderSelectView.tsx
├── components/
│   ├── StatusBar.tsx   # Top bar (folder, status)
│   ├── HelpBar.tsx     # Bottom bar (shortcuts)
│   ├── EmailListItem.tsx
│   ├── ScrollableText.tsx
│   ├── BodyEditor.tsx  # Multi-line input
│   ├── LoadingSpinner.tsx
│   └── ErrorModal.tsx
└── utils/
    ├── html-to-text.ts # HTML stripping
    └── truncate.ts     # Text truncation
```

## Keyboard Shortcuts

| Context | Key | Action |
|---------|-----|--------|
| Global | `q` | Quit |
| Global | `?` | Help |
| Inbox | `↑/↓` | Navigate |
| Inbox | `Enter` | Open email |
| Inbox | `c` | Compose |
| Inbox | `/` | Search |
| Inbox | `f` | Folders |
| Inbox | `r` | Refresh |
| Email | `Esc` | Back |
| Email | `r` | Reply |
| Email | `R` | Reply all |
| Email | `d` | Delete |
| Compose | `Esc` | Cancel |
| Compose | `Ctrl+Enter` | Send |
| Compose | `Tab` | Next field |

## Dependencies

```json
{
  "dependencies": {
    "ink": "^5.0.1",
    "react": "^18.3.1",
    "@inkjs/ui": "^2.0.0"
  },
  "devDependencies": {
    "@types/react": "^18.3.0"
  }
}
```

## Error Handling

- **Auth errors:** Show message, exit to `myoffice login`
- **Network errors:** Show retry option
- **API errors:** Display error modal with dismiss
- **Validation:** Highlight invalid fields

## Testing Strategy

1. Manual testing of all keyboard navigation flows
2. Test with different terminal sizes (80x24 minimum)
3. Test error states (network off, invalid data)
4. Test edge cases (empty inbox, long emails)
