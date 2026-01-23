# Email TUI Client - Requirements

## Overview

A text-based user interface (TUI) email client for quick inbox triage. Built on top of the existing myoffice CLI email tools, providing an interactive terminal experience for reading, replying, composing, and searching emails.

## User Stories

**As a** terminal-focused developer
**I want** an interactive email client in my terminal
**So that** I can triage my inbox without leaving my workflow

**As a** busy professional
**I want** to quickly scan and process emails
**So that** I can stay on top of my inbox efficiently

## Acceptance Criteria

### Navigation & Display

1. WHEN the client starts THEN system SHALL display a list of inbox emails showing sender, subject, date, and read status
2. WHEN user presses Up/Down arrows THEN system SHALL move selection highlight between emails
3. WHEN user presses Enter on an email THEN system SHALL display the full email content in a full-screen view
4. WHEN viewing an email AND user presses Escape/Backspace THEN system SHALL return to the inbox list
5. WHEN emails are displayed THEN system SHALL show unread emails with a visual indicator

### Email Actions

6. WHEN viewing an email AND user presses 'r' THEN system SHALL open reply composer
7. WHEN viewing an email AND user presses 'R' THEN system SHALL open reply-all composer
8. WHEN in inbox list AND user presses 'c' THEN system SHALL open new email composer
9. WHEN composing AND user presses Ctrl+Enter or designated send key THEN system SHALL send the email
10. WHEN composing AND user presses Escape THEN system SHALL cancel and return to previous view

### Folder Navigation

11. WHEN user presses designated folder key THEN system SHALL show folder selection (inbox, sent, drafts)
12. WHEN folder is selected THEN system SHALL load and display emails from that folder

### Search

13. WHEN user presses '/' or designated search key THEN system SHALL show search input
14. WHEN search query is entered THEN system SHALL display matching emails
15. WHEN in search results AND user presses Escape THEN system SHALL return to previous view

### Status & Feedback

16. WHEN an action is in progress THEN system SHALL show loading indicator
17. WHEN an error occurs THEN system SHALL display error message with recovery option
18. WHEN email is sent successfully THEN system SHALL show confirmation and return to inbox

## Edge Cases

- Empty inbox: Show "No emails" message
- Network error: Show error with retry option
- Very long email body: Scrollable view
- HTML-only emails: Display text content or indicate HTML-only
- Offline state: Show cached data or connection error
- Large inbox: Paginate or virtual scroll (initial load of 25 emails)

## Constraints

### Technical
- Must use Ink (React for CLI) as the TUI framework
- Must reuse existing myoffice CLI email tools (no direct Graph API calls)
- Must work in standard terminal sizes (80x24 minimum)
- Must support authentication check on startup

### UX
- Arrow key navigation (no vim keybindings required)
- Full-screen view switching (not split panes)
- Response time: List should load within 3 seconds on good connection

## Out of Scope

- Attachment viewing/downloading
- Folder management (create/delete folders)
- Email threading/conversation view
- Multiple account support
- Offline caching
- Vim keybindings
- Drag-and-drop
- Rich HTML rendering (text conversion only)
