# Microsoft Planner Integration - Design

## Overview

Add a new `planner.ts` tool module following the existing pattern (tasks.ts, mail.ts). The module will interact with Microsoft Graph Planner API endpoints, handling the ETag-based concurrency model that Planner requires.

## Architecture

```
src/
├── index.ts              # Add planner tool definitions + case handlers
├── auth/
│   └── config.ts         # Add Group.Read.All scope
└── tools/
    └── planner.ts        # NEW: Planner tool implementations
```

The Planner module follows the same pattern as other tool modules:
- Zod schemas for input validation
- TypeScript interfaces for Graph API responses
- Functions that call graphRequest/graphList utilities
- Export schemas and functions for use in index.ts

## Components

### Plans Component

**Purpose:** List and retrieve Planner plans the user has access to

**Graph API Endpoints:**
- `GET /me/planner/plans` - List all plans user has access to
- `GET /planner/plans/{planId}` - Get specific plan details

**Functions:**
- `listPlans()` - Returns all plans with id, title, owner group
- `getPlan(planId)` - Returns single plan details

### Buckets Component

**Purpose:** Manage buckets (columns) within a plan

**Graph API Endpoints:**
- `GET /planner/plans/{planId}/buckets` - List buckets
- `POST /planner/buckets` - Create bucket
- `PATCH /planner/buckets/{bucketId}` - Update bucket
- `DELETE /planner/buckets/{bucketId}` - Delete bucket

**Functions:**
- `listBuckets(planId)` - Returns buckets with id, name, orderHint
- `createBucket(planId, name)` - Creates new bucket
- `updateBucket(bucketId, name)` - Updates bucket (fetches ETag internally)
- `deleteBucket(bucketId)` - Deletes bucket (fetches ETag internally)

### Tasks Component

**Purpose:** Full CRUD on Planner tasks

**Graph API Endpoints:**
- `GET /planner/plans/{planId}/tasks` - List tasks in plan
- `GET /planner/tasks/{taskId}` - Get specific task
- `POST /planner/tasks` - Create task
- `PATCH /planner/tasks/{taskId}` - Update task
- `DELETE /planner/tasks/{taskId}` - Delete task

**Functions:**
- `listPlannerTasks(planId)` - Returns tasks with key fields
- `getPlannerTask(taskId)` - Returns single task
- `createPlannerTask(planId, title, bucketId?, ...)` - Creates task
- `updatePlannerTask(taskId, updates)` - Updates task (fetches ETag internally)
- `deletePlannerTask(taskId)` - Deletes task (fetches ETag internally)

### Task Details Component

**Purpose:** Manage extended task information (description, checklist)

**Graph API Endpoints:**
- `GET /planner/tasks/{taskId}/details` - Get task details
- `PATCH /planner/tasks/{taskId}/details` - Update task details

**Functions:**
- `getPlannerTaskDetails(taskId)` - Returns description, checklist, references
- `updatePlannerTaskDetails(taskId, updates)` - Updates details (fetches ETag internally)

## Data Models

```typescript
// Plan
interface PlannerPlan {
  id: string;
  title: string;
  createdDateTime: string;
  owner: string; // Group ID
  createdBy: { user: { id: string } };
}

// Bucket
interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint: string;
}

// Task
interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  percentComplete: number; // 0, 50, or 100
  priority: number; // 0-10, where 1=urgent, 3=important, 5=medium, 9=low
  startDateTime?: string;
  dueDateTime?: string;
  assignments: Record<string, { assignedBy: object; assignedDateTime: string }>;
  createdDateTime: string;
  orderHint: string;
}

// Task Details
interface PlannerTaskDetails {
  id: string;
  description: string;
  checklist: Record<string, {
    isChecked: boolean;
    title: string;
    orderHint: string;
  }>;
  references: Record<string, {
    alias: string;
    type: string;
  }>;
}
```

## Key Decisions

### Decision: ETag Handling Strategy

**Context:** Planner API requires ETags for all update/delete operations to prevent concurrent modification conflicts.

**Options:**
1. Require client to provide ETag - Pros: No extra API call / Cons: Bad UX, client must track ETags
2. Fetch ETag before each operation - Pros: Transparent to user / Cons: Extra API call per operation
3. Optimistic with retry - Pros: Usually one call / Cons: Complex, still needs retry logic

**Chosen:** Option 2 - Fetch ETag internally before each update/delete. This matches the pattern users expect from other tools (just provide ID + changes) and the extra API call is acceptable for correctness.

### Decision: User Assignment Format

**Context:** Planner assignments use Azure AD user IDs, not email addresses.

**Options:**
1. Accept only user IDs - Pros: Direct mapping / Cons: Users don't know their IDs
2. Accept email, resolve to ID - Pros: Better UX / Cons: Extra API call to resolve
3. Accept both formats - Pros: Flexible / Cons: Complex logic

**Chosen:** Option 2 - Accept email addresses and resolve to user IDs via `/users/{email}` endpoint. This provides a consistent UX with other tools that use email addresses.

### Decision: Progress Values

**Context:** Planner only supports 0%, 50%, or 100% complete.

**Options:**
1. Accept any number, map to nearest - Pros: Flexible / Cons: Confusing behavior
2. Accept only valid values (0, 50, 100) - Pros: Clear / Cons: Less flexible
3. Accept semantic values (notStarted, inProgress, completed) - Pros: Clear intent / Cons: Different from API

**Chosen:** Option 3 - Accept semantic values that map to percentComplete:
- `notStarted` → 0
- `inProgress` → 50
- `completed` → 100

This is clearer than raw numbers and matches how users think about task status.

### Decision: Priority Values

**Context:** Planner uses numeric priority (1=urgent to 9=low).

**Chosen:** Accept semantic values: `urgent`, `important`, `medium`, `low` mapping to 1, 3, 5, 9.

## Error Handling

| Error | Cause | Response |
|-------|-------|----------|
| 403 Forbidden | User doesn't have access to plan/task | "You don't have access to this resource" |
| 404 Not Found | Plan/bucket/task doesn't exist | "Resource not found" |
| 409 Conflict | ETag mismatch (concurrent edit) | "Resource was modified by another user. Please try again." |
| 400 Bad Request | Invalid input | Return Graph API error message |

All operations wrap errors with context about what operation failed.

## Testing Strategy

1. **Manual testing via Claude Code:**
   - List plans, verify correct plans returned
   - CRUD operations on buckets
   - CRUD operations on tasks
   - Update task details (description, checklist)
   - Test concurrent modification error handling

2. **Test scenarios:**
   - User with no plans
   - User with multiple plans across different groups
   - Bucket operations with special characters in names
   - Task assignments with valid/invalid email addresses
   - Progress and priority updates
   - Task details with empty/full checklists

3. **Permission testing:**
   - Verify Group.Read.All scope is requested on login
   - Verify user must re-authenticate after scope change
