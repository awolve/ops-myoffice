# Microsoft Planner Integration - Tasks

## Overview
Estimated scope: Medium (3-4 implementation sessions)

## Tasks

- [ ] 1. **Add Graph API Permission**
  - [ ] 1.1 Add Group.Read.All scope to auth config
    - Files: `src/auth/config.ts`
    - Requirements: Enables access to plans via groups
    - Note: Users will need to re-authenticate after this change

- [ ] 2. **Create Planner Tool Module**
  - [ ] 2.1 Create planner.ts with types and schemas
    - Files: `src/tools/planner.ts`
    - Requirements: R1-R16
    - Details:
      - Define TypeScript interfaces (PlannerPlan, PlannerBucket, PlannerTask, PlannerTaskDetails)
      - Define Zod schemas for all operations
      - Export schemas for use in index.ts

  - [ ] 2.2 Implement ETag helper function
    - Files: `src/tools/planner.ts`
    - Requirements: R6, R7, R8, R11, R12, R16
    - Details:
      - `getETag(endpoint)` - Fetches resource and extracts @odata.etag
      - Used by all update/delete operations

  - [ ] 2.3 Implement user ID resolver helper
    - Files: `src/tools/planner.ts`
    - Requirements: R13
    - Details:
      - `resolveUserId(emailOrId)` - Returns user ID from email address
      - Cache results within session to avoid repeated lookups

- [ ] 3. **Implement Plans Operations (Read-Only)**
  - [ ] 3.1 Implement listPlans
    - Files: `src/tools/planner.ts`
    - Requirements: R1
    - Endpoint: `GET /me/planner/plans`
    - Returns: Array of { id, title, owner, createdDateTime }

  - [ ] 3.2 Implement getPlan
    - Files: `src/tools/planner.ts`
    - Requirements: R2, R3
    - Endpoint: `GET /planner/plans/{planId}`
    - Returns: Full plan details

- [ ] 4. **Implement Buckets Operations (Full CRUD)**
  - [ ] 4.1 Implement listBuckets
    - Files: `src/tools/planner.ts`
    - Requirements: R4
    - Endpoint: `GET /planner/plans/{planId}/buckets`
    - Returns: Array of { id, name, planId, orderHint }

  - [ ] 4.2 Implement createBucket
    - Files: `src/tools/planner.ts`
    - Requirements: R5
    - Endpoint: `POST /planner/buckets`
    - Body: { planId, name }

  - [ ] 4.3 Implement updateBucket
    - Files: `src/tools/planner.ts`
    - Requirements: R6, R8
    - Endpoint: `PATCH /planner/buckets/{bucketId}`
    - Must fetch ETag first, include in If-Match header

  - [ ] 4.4 Implement deleteBucket
    - Files: `src/tools/planner.ts`
    - Requirements: R7, R8
    - Endpoint: `DELETE /planner/buckets/{bucketId}`
    - Must fetch ETag first, include in If-Match header

- [ ] 5. **Implement Tasks Operations (Full CRUD)**
  - [ ] 5.1 Implement listPlannerTasks
    - Files: `src/tools/planner.ts`
    - Requirements: R9
    - Endpoint: `GET /planner/plans/{planId}/tasks`
    - Returns: Array with key fields (title, bucketId, assignments, dueDateTime, percentComplete, priority)

  - [ ] 5.2 Implement getPlannerTask
    - Files: `src/tools/planner.ts`
    - Requirements: R9
    - Endpoint: `GET /planner/tasks/{taskId}`

  - [ ] 5.3 Implement createPlannerTask
    - Files: `src/tools/planner.ts`
    - Requirements: R10, R13, R14
    - Endpoint: `POST /planner/tasks`
    - Body: { planId, bucketId?, title, assignments?, dueDateTime?, priority? }
    - Handle assignment email-to-ID resolution
    - Handle progress/priority semantic values

  - [ ] 5.4 Implement updatePlannerTask
    - Files: `src/tools/planner.ts`
    - Requirements: R11, R13, R14
    - Endpoint: `PATCH /planner/tasks/{taskId}`
    - Must fetch ETag first
    - Handle partial updates

  - [ ] 5.5 Implement deletePlannerTask
    - Files: `src/tools/planner.ts`
    - Requirements: R12
    - Endpoint: `DELETE /planner/tasks/{taskId}`
    - Must fetch ETag first

- [ ] 6. **Implement Task Details Operations**
  - [ ] 6.1 Implement getPlannerTaskDetails
    - Files: `src/tools/planner.ts`
    - Requirements: R15
    - Endpoint: `GET /planner/tasks/{taskId}/details`
    - Returns: description, checklist, references

  - [ ] 6.2 Implement updatePlannerTaskDetails
    - Files: `src/tools/planner.ts`
    - Requirements: R16
    - Endpoint: `PATCH /planner/tasks/{taskId}/details`
    - Must fetch ETag first
    - Handle checklist item add/update/remove

- [ ] 7. **Register Tools in index.ts**
  - [ ] 7.1 Add tool definitions to TOOLS array
    - Files: `src/index.ts`
    - Tools to add:
      - `planner_list_plans` - List all accessible plans
      - `planner_get_plan` - Get plan details
      - `planner_list_buckets` - List buckets in a plan
      - `planner_create_bucket` - Create a bucket
      - `planner_update_bucket` - Update bucket name
      - `planner_delete_bucket` - Delete a bucket
      - `planner_list_tasks` - List tasks in a plan
      - `planner_get_task` - Get task details
      - `planner_create_task` - Create a task
      - `planner_update_task` - Update a task
      - `planner_delete_task` - Delete a task
      - `planner_get_task_details` - Get task description/checklist
      - `planner_update_task_details` - Update description/checklist

  - [ ] 7.2 Add case handlers in CallToolRequestSchema
    - Files: `src/index.ts`
    - Add switch cases for all 13 planner tools

  - [ ] 7.3 Add planner import
    - Files: `src/index.ts`
    - Add: `import * as planner from './tools/planner.js'`

- [ ] 8. **Update Exports**
  - [ ] 8.1 Export planner module from tools index
    - Files: `src/tools/index.ts`
    - Add: `export * from './planner.js'`

- [ ] 9. **Testing & Documentation**
  - [ ] 9.1 Manual testing with Claude Code
    - Test all operations against a real M365 tenant
    - Verify error handling for 403, 404, 409
    - Test concurrent modification scenario

  - [ ] 9.2 Update CLAUDE.md with planner info
    - Files: `CLAUDE.md`
    - Document new tools and their usage

## Implementation Order

Recommended sequence:

1. **Permission first (Task 1)** - Add the scope so subsequent testing works
2. **Module skeleton (Task 2)** - Types, schemas, helpers
3. **Plans (Task 3)** - Read-only, simplest to test
4. **Buckets (Task 4)** - Tests ETag handling pattern
5. **Tasks (Task 5)** - Core functionality, uses patterns from buckets
6. **Task Details (Task 6)** - Extends task capabilities
7. **Wire up (Tasks 7-8)** - Connect everything in index.ts
8. **Test & document (Task 9)** - Validate and document

## Dependencies

- Existing: `graphRequest`, `graphList` from `src/utils/graph-client.ts`
- Existing: Zod for schema validation
- New permission: `Group.Read.All` scope (users must re-authenticate)

## Notes

- ETag header name is `If-Match` for Graph API
- Graph returns ETag in `@odata.etag` property
- Checklist items use GUIDs as keys in the checklist object
- Assignment keys are user IDs, values are assignment metadata
