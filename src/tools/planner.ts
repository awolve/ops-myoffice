import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// ============================================================================
// Types
// ============================================================================

interface PlannerPlan {
  id: string;
  title: string;
  createdDateTime: string;
  owner: string;
  createdBy?: {
    user?: { id: string; displayName?: string };
  };
}

interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint: string;
  '@odata.etag'?: string;
}

interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  percentComplete: number;
  priority: number;
  startDateTime?: string;
  dueDateTime?: string;
  assignments: Record<string, { assignedBy?: object; assignedDateTime?: string }>;
  createdDateTime: string;
  orderHint: string;
  '@odata.etag'?: string;
}

interface PlannerTaskDetails {
  id: string;
  description: string;
  checklist: Record<
    string,
    {
      isChecked: boolean;
      title: string;
      orderHint: string;
    }
  >;
  references: Record<
    string,
    {
      alias?: string;
      type?: string;
    }
  >;
  '@odata.etag'?: string;
}

interface GraphUser {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

// ============================================================================
// Schemas
// ============================================================================

// Plans (read-only)
export const listPlansSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of plans. Default: 50'),
});

export const getPlanSchema = z.object({
  planId: z.string().describe('The plan ID'),
});

// Buckets
export const listBucketsSchema = z.object({
  planId: z.string().describe('The plan ID'),
});

export const createBucketSchema = z.object({
  planId: z.string().describe('The plan ID'),
  name: z.string().describe('Bucket name'),
});

export const updateBucketSchema = z.object({
  bucketId: z.string().describe('The bucket ID'),
  name: z.string().describe('New bucket name'),
});

export const deleteBucketSchema = z.object({
  bucketId: z.string().describe('The bucket ID'),
});

// Tasks
export const listPlannerTasksSchema = z.object({
  planId: z.string().describe('The plan ID'),
  bucketId: z.string().optional().describe('Filter by bucket ID'),
  maxItems: z.number().optional().describe('Maximum number of tasks. Default: 100'),
});

export const getPlannerTaskSchema = z.object({
  taskId: z.string().describe('The task ID'),
});

export const createPlannerTaskSchema = z.object({
  planId: z.string().describe('The plan ID'),
  title: z.string().describe('Task title'),
  bucketId: z.string().optional().describe('Bucket ID to place the task in'),
  assignments: z
    .array(z.string())
    .optional()
    .describe('Email addresses of users to assign'),
  dueDateTime: z.string().optional().describe('Due date (ISO format)'),
  startDateTime: z.string().optional().describe('Start date (ISO format)'),
  priority: z
    .enum(['urgent', 'important', 'medium', 'low'])
    .optional()
    .describe('Task priority'),
  progress: z
    .enum(['notStarted', 'inProgress', 'completed'])
    .optional()
    .describe('Task progress'),
});

export const updatePlannerTaskSchema = z.object({
  taskId: z.string().describe('The task ID'),
  title: z.string().optional().describe('New task title'),
  bucketId: z.string().optional().describe('Move to different bucket'),
  assignments: z
    .array(z.string())
    .optional()
    .describe('Email addresses of users to assign (replaces existing)'),
  dueDateTime: z.string().optional().describe('New due date (ISO format)'),
  startDateTime: z.string().optional().describe('New start date (ISO format)'),
  priority: z
    .enum(['urgent', 'important', 'medium', 'low'])
    .optional()
    .describe('New priority'),
  progress: z
    .enum(['notStarted', 'inProgress', 'completed'])
    .optional()
    .describe('New progress'),
});

export const deletePlannerTaskSchema = z.object({
  taskId: z.string().describe('The task ID'),
});

// Task Details
export const getPlannerTaskDetailsSchema = z.object({
  taskId: z.string().describe('The task ID'),
});

export const updatePlannerTaskDetailsSchema = z.object({
  taskId: z.string().describe('The task ID'),
  description: z.string().optional().describe('Task description'),
  checklist: z
    .array(
      z.object({
        title: z.string().describe('Checklist item title'),
        isChecked: z.boolean().optional().describe('Whether item is checked'),
      })
    )
    .optional()
    .describe('Checklist items (replaces existing checklist)'),
});

// ============================================================================
// Helpers
// ============================================================================

// Priority mapping: semantic -> numeric
const PRIORITY_MAP: Record<string, number> = {
  urgent: 1,
  important: 3,
  medium: 5,
  low: 9,
};

// Reverse priority mapping for display
const PRIORITY_REVERSE: Record<number, string> = {
  1: 'urgent',
  3: 'important',
  5: 'medium',
  9: 'low',
};

// Progress mapping: semantic -> percentComplete
const PROGRESS_MAP: Record<string, number> = {
  notStarted: 0,
  inProgress: 50,
  completed: 100,
};

// Reverse progress mapping for display
function progressToSemantic(percentComplete: number): string {
  if (percentComplete === 0) return 'notStarted';
  if (percentComplete === 100) return 'completed';
  return 'inProgress';
}

// Get ETag for a resource
async function getETag(endpoint: string): Promise<string> {
  const resource = await graphRequest<{ '@odata.etag': string }>(endpoint);
  const etag = resource['@odata.etag'];
  if (!etag) {
    throw new Error('Resource does not have an ETag');
  }
  return etag;
}

// Simple cache for user ID lookups within a session
const userIdCache = new Map<string, string>();

// Resolve email to user ID
async function resolveUserId(emailOrId: string): Promise<string> {
  // If it looks like a GUID, assume it's already a user ID
  if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(emailOrId)) {
    return emailOrId;
  }

  // Check cache
  const cached = userIdCache.get(emailOrId.toLowerCase());
  if (cached) {
    return cached;
  }

  // Look up user by email
  const user = await graphRequest<GraphUser>(`/users/${encodeURIComponent(emailOrId)}`);
  userIdCache.set(emailOrId.toLowerCase(), user.id);
  return user.id;
}

// ============================================================================
// Plans (Read-Only)
// ============================================================================

export async function listPlans(params: z.infer<typeof listPlansSchema>) {
  const { maxItems = 50 } = params;

  const plans = await graphList<PlannerPlan>('/me/planner/plans', { maxItems });

  return plans.map((p) => ({
    id: p.id,
    title: p.title,
    owner: p.owner,
    createdDateTime: p.createdDateTime,
  }));
}

export async function getPlan(params: z.infer<typeof getPlanSchema>) {
  const { planId } = params;

  const plan = await graphRequest<PlannerPlan>(`/planner/plans/${planId}`);

  return {
    id: plan.id,
    title: plan.title,
    owner: plan.owner,
    createdDateTime: plan.createdDateTime,
    createdBy: plan.createdBy?.user,
  };
}

// ============================================================================
// Buckets
// ============================================================================

export async function listBuckets(params: z.infer<typeof listBucketsSchema>) {
  const { planId } = params;

  const buckets = await graphList<PlannerBucket>(`/planner/plans/${planId}/buckets`);

  return buckets.map((b) => ({
    id: b.id,
    name: b.name,
    planId: b.planId,
    orderHint: b.orderHint,
  }));
}

export async function createBucket(params: z.infer<typeof createBucketSchema>) {
  const { planId, name } = params;

  const bucket = await graphRequest<PlannerBucket>('/planner/buckets', {
    method: 'POST',
    body: { planId, name },
  });

  return {
    success: true,
    bucketId: bucket.id,
    name: bucket.name,
  };
}

export async function updateBucket(params: z.infer<typeof updateBucketSchema>) {
  const { bucketId, name } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/buckets/${bucketId}`);

  const bucket = await graphRequest<PlannerBucket>(`/planner/buckets/${bucketId}`, {
    method: 'PATCH',
    body: { name },
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    bucketId: bucket.id,
    name: bucket.name,
  };
}

export async function deleteBucket(params: z.infer<typeof deleteBucketSchema>) {
  const { bucketId } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/buckets/${bucketId}`);

  await graphRequest(`/planner/buckets/${bucketId}`, {
    method: 'DELETE',
    headers: { 'If-Match': etag },
  });

  return { success: true, message: 'Bucket deleted' };
}

// ============================================================================
// Tasks
// ============================================================================

export async function listPlannerTasks(params: z.infer<typeof listPlannerTasksSchema>) {
  const { planId, bucketId, maxItems = 100 } = params;

  let tasks = await graphList<PlannerTask>(`/planner/plans/${planId}/tasks`, { maxItems });

  // Filter by bucket if specified
  if (bucketId) {
    tasks = tasks.filter((t) => t.bucketId === bucketId);
  }

  return tasks.map((t) => ({
    id: t.id,
    title: t.title,
    bucketId: t.bucketId,
    progress: progressToSemantic(t.percentComplete),
    percentComplete: t.percentComplete,
    priority: PRIORITY_REVERSE[t.priority] || 'medium',
    dueDateTime: t.dueDateTime,
    startDateTime: t.startDateTime,
    assignedTo: Object.keys(t.assignments || {}),
    createdDateTime: t.createdDateTime,
  }));
}

export async function getPlannerTask(params: z.infer<typeof getPlannerTaskSchema>) {
  const { taskId } = params;

  const task = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`);

  return {
    id: task.id,
    planId: task.planId,
    bucketId: task.bucketId,
    title: task.title,
    progress: progressToSemantic(task.percentComplete),
    percentComplete: task.percentComplete,
    priority: PRIORITY_REVERSE[task.priority] || 'medium',
    dueDateTime: task.dueDateTime,
    startDateTime: task.startDateTime,
    assignedTo: Object.keys(task.assignments || {}),
    createdDateTime: task.createdDateTime,
  };
}

export async function createPlannerTask(params: z.infer<typeof createPlannerTaskSchema>) {
  const { planId, title, bucketId, assignments, dueDateTime, startDateTime, priority, progress } =
    params;

  const body: Record<string, unknown> = {
    planId,
    title,
  };

  if (bucketId) body.bucketId = bucketId;
  if (dueDateTime) body.dueDateTime = dueDateTime;
  if (startDateTime) body.startDateTime = startDateTime;
  if (priority) body.priority = PRIORITY_MAP[priority];
  if (progress) body.percentComplete = PROGRESS_MAP[progress];

  // Resolve assignments to user IDs
  if (assignments && assignments.length > 0) {
    const assignmentObj: Record<string, { '@odata.type': string }> = {};
    for (const email of assignments) {
      const userId = await resolveUserId(email);
      assignmentObj[userId] = { '@odata.type': '#microsoft.graph.plannerAssignment' };
    }
    body.assignments = assignmentObj;
  }

  const task = await graphRequest<PlannerTask>('/planner/tasks', {
    method: 'POST',
    body,
  });

  return {
    success: true,
    taskId: task.id,
    title: task.title,
  };
}

export async function updatePlannerTask(params: z.infer<typeof updatePlannerTaskSchema>) {
  const { taskId, title, bucketId, assignments, dueDateTime, startDateTime, priority, progress } =
    params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}`);

  const body: Record<string, unknown> = {};

  if (title) body.title = title;
  if (bucketId) body.bucketId = bucketId;
  if (dueDateTime) body.dueDateTime = dueDateTime;
  if (startDateTime) body.startDateTime = startDateTime;
  if (priority) body.priority = PRIORITY_MAP[priority];
  if (progress) body.percentComplete = PROGRESS_MAP[progress];

  // Resolve assignments to user IDs
  if (assignments) {
    const assignmentObj: Record<string, { '@odata.type': string } | null> = {};
    for (const email of assignments) {
      const userId = await resolveUserId(email);
      assignmentObj[userId] = { '@odata.type': '#microsoft.graph.plannerAssignment' };
    }
    body.assignments = assignmentObj;
  }

  const task = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    taskId: task.id,
    title: task.title,
  };
}

export async function deletePlannerTask(params: z.infer<typeof deletePlannerTaskSchema>) {
  const { taskId } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}`);

  await graphRequest(`/planner/tasks/${taskId}`, {
    method: 'DELETE',
    headers: { 'If-Match': etag },
  });

  return { success: true, message: 'Task deleted' };
}

// ============================================================================
// Task Details
// ============================================================================

export async function getPlannerTaskDetails(params: z.infer<typeof getPlannerTaskDetailsSchema>) {
  const { taskId } = params;

  const details = await graphRequest<PlannerTaskDetails>(`/planner/tasks/${taskId}/details`);

  // Transform checklist to array for easier consumption
  const checklistItems = Object.entries(details.checklist || {}).map(([id, item]) => ({
    id,
    title: item.title,
    isChecked: item.isChecked,
  }));

  // Transform references to array
  const referenceItems = Object.entries(details.references || {}).map(([url, ref]) => ({
    url: decodeURIComponent(url),
    alias: ref.alias,
    type: ref.type,
  }));

  return {
    taskId,
    description: details.description,
    checklist: checklistItems,
    references: referenceItems,
  };
}

export async function updatePlannerTaskDetails(
  params: z.infer<typeof updatePlannerTaskDetailsSchema>
) {
  const { taskId, description, checklist } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}/details`);

  const body: Record<string, unknown> = {};

  if (description !== undefined) {
    body.description = description;
  }

  // Transform checklist array to object format
  if (checklist) {
    const checklistObj: Record<
      string,
      { '@odata.type': string; title: string; isChecked: boolean }
    > = {};
    checklist.forEach((item, index) => {
      // Generate a simple GUID-like key for new items
      const key = crypto.randomUUID();
      checklistObj[key] = {
        '@odata.type': 'microsoft.graph.plannerChecklistItem',
        title: item.title,
        isChecked: item.isChecked ?? false,
      };
    });
    body.checklist = checklistObj;
  }

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return { success: true, message: 'Task details updated' };
}
