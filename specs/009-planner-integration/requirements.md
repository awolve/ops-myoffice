# Microsoft Planner Integration - Requirements

## Overview

Add Microsoft Planner support to the MyOffice MCP server, enabling users to manage tasks within Planner plans they have access to. Unlike Microsoft To Do (personal task management), Planner provides team-oriented task management with plans, buckets, and rich task features.

## User Stories

**As a** Microsoft 365 user
**I want** to view and manage my Planner tasks through the MCP server
**So that** I can organize and track team work without switching to the Planner UI

**As a** team member
**I want** to create, update, and organize tasks within existing plans
**So that** I can contribute to project management efficiently

**As a** project coordinator
**I want** to organize tasks into buckets
**So that** I can maintain clear task categorization (e.g., To Do, In Progress, Done)

## Acceptance Criteria

### Plans (Read-Only)

1. WHEN user requests plan list THEN system SHALL return all Planner plans the user has access to
2. WHEN user requests a specific plan THEN system SHALL return plan details including title, owner group, and creation date
3. IF user requests a plan they don't have access to THEN system SHALL return an appropriate error

### Buckets (Full CRUD)

4. WHEN user requests buckets for a plan THEN system SHALL return all buckets in that plan with their names and order
5. WHEN user creates a bucket THEN system SHALL create it in the specified plan with the given name
6. WHEN user updates a bucket THEN system SHALL update the bucket name (requires ETag)
7. WHEN user deletes a bucket THEN system SHALL remove the bucket and its tasks (requires ETag)
8. IF bucket operation fails due to ETag mismatch THEN system SHALL return a concurrency error

### Tasks (Full CRUD)

9. WHEN user requests tasks for a plan THEN system SHALL return all tasks with key fields (title, bucket, assignments, due date, progress, priority)
10. WHEN user creates a task THEN system SHALL create it in the specified plan with provided details
11. WHEN user updates a task THEN system SHALL update the specified fields (requires ETag)
12. WHEN user deletes a task THEN system SHALL remove the task (requires ETag)
13. WHEN user assigns a task THEN system SHALL update the assignments using user IDs
14. WHEN user changes task progress THEN system SHALL update the percentComplete field (0, 50, or 100)

### Task Details (Read/Update)

15. WHEN user requests task details THEN system SHALL return description, checklist items, and references
16. WHEN user updates task details THEN system SHALL update the description and/or checklist (requires ETag)

## Edge Cases

- User has no plans (should return empty list, not error)
- User loses access to a plan mid-operation (should return 403/404 error gracefully)
- Concurrent modification by another user (ETag mismatch - should explain the conflict)
- Plan has no buckets (should return empty list)
- Task has no assignments (should handle gracefully)
- Very long task titles or descriptions (Graph API has limits)
- Special characters in bucket/task names

## Constraints

### Technical Constraints
- All Planner updates require ETags for optimistic concurrency control
- Plans must belong to M365 Groups - cannot exist standalone
- Task assignments use Azure AD user IDs, not email addresses
- Task progress is limited to 0%, 50%, or 100%
- Labels are defined at plan level (category descriptions), tasks reference them by index

### Business Constraints
- No plan creation (requires M365 Group creation which is out of scope)
- No plan deletion (too destructive, manage via Planner UI)
- User can only access plans in groups they're a member of

### API Permissions Required
- `Tasks.ReadWrite` - Already exists, may work for Planner
- `Group.Read.All` - New, needed to list groups/plans user has access to

## Out of Scope

- Creating new Planner plans (requires M365 Group)
- Deleting plans
- Managing M365 Groups
- Plan templates
- Planner premium features (goals, portfolios)
- Exporting plans
- Plan analytics/insights
