# Personal M365 MCP - Requirements

> Retroactive spec - documented after implementation

## Summary

A lightweight MCP (Model Context Protocol) server for personal Microsoft 365 access, designed for AI assistants like Claude Code. Unlike the admin-focused M365 MCP, this uses **delegated authentication** so employees can access their own email, calendar, tasks, and files.

## Background

Awolve has an existing M365 MCP (`m365-core-mcp`) configured with application permissions for admin use cases. However, employees need a way for their personal AI assistants to access their own M365 data without admin privileges.

## Goals

1. Enable employees to use Claude Code with their personal M365 data
2. Use delegated auth (OAuth device code flow) - users authenticate as themselves
3. Provide core productivity tools: Email, Calendar, Tasks, OneDrive, Contacts
4. Keep it simple - personal use only, no admin features
5. Package for distribution via awolve-general plugin

## Non-Goals

- Admin operations (managing other users, groups, policies)
- Compliance/security features
- Application-level permissions
- Shared mailbox access

## Functional Requirements

### Authentication

- [x] Device code flow for initial authentication
- [x] Token caching with automatic refresh
- [x] Support for multi-tenant (any M365 organization)
- [x] Clear auth status feedback

### Email (Mail)

- [x] List emails from any folder (inbox, sent, drafts, etc.)
- [x] Read individual email with full body
- [x] Search emails by query
- [x] Send emails (with AI confirmation requirement)
- [x] Delete emails (with AI confirmation requirement)

### Calendar

- [x] List events within date range
- [x] Get event details including attendees
- [x] Create events with optional Teams meeting
- [x] Update existing events
- [x] Delete events (with AI confirmation requirement)

### Tasks (Microsoft To Do)

- [x] List task lists
- [x] List tasks from a list
- [x] Create tasks with due date and importance
- [x] Update tasks
- [x] Mark tasks as complete
- [x] Delete tasks

### OneDrive

- [x] List files and folders
- [x] Get file/folder metadata
- [x] Search files
- [x] Read text file content (< 1MB)
- [x] Create folders

### Contacts

- [x] List contacts
- [x] Search contacts by name/email/company
- [x] Get contact details

## Security Requirements

- Tokens stored locally with restrictive permissions (0600)
- No client secrets needed (public client)
- Delegated permissions only - users can only access their own data
- Sensitive actions (send, delete) marked in tool descriptions for AI confirmation

## Permissions Required (Azure AD)

| Permission | Type | Purpose |
|------------|------|---------|
| Mail.ReadWrite | Delegated | Read and manage email |
| Mail.Send | Delegated | Send email |
| Calendars.ReadWrite | Delegated | Read and manage calendar |
| Tasks.ReadWrite | Delegated | Read and manage To Do tasks |
| Files.ReadWrite | Delegated | Read and manage OneDrive files |
| Contacts.Read | Delegated | Read contacts |
| User.Read | Delegated | Get current user info |
| offline_access | Delegated | Refresh tokens |
