import { z } from 'zod';
import { basename } from 'path';
import { readFile } from 'fs/promises';
import { graphRequest, graphList, graphUpload, graphUploadLarge } from '../utils/graph-client.js';

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
  conversationThreadId?: string;
  '@odata.etag'?: string;
}

interface ConversationPost {
  id: string;
  body: { content: string; contentType: string };
  from: { emailAddress: { name: string; address: string } };
  receivedDateTime: string;
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

interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl: string;
  file?: { mimeType: string };
}

interface Drive {
  id: string;
  name: string;
  webUrl: string;
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
  clearAssignments: z.boolean().optional().describe('Remove all assignments'),
  dueDateTime: z.string().optional().describe('New due date (ISO format)'),
  clearDue: z.boolean().optional().describe('Clear the due date'),
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

// Checklist operations
export const addChecklistItemSchema = z.object({
  taskId: z.string().describe('The task ID'),
  title: z.string().describe('Checklist item title'),
  isChecked: z.boolean().optional().describe('Whether item is checked. Default: false'),
});

export const removeChecklistItemSchema = z.object({
  taskId: z.string().describe('The task ID'),
  itemId: z.string().describe('The checklist item ID'),
});

export const toggleChecklistItemSchema = z.object({
  taskId: z.string().describe('The task ID'),
  itemId: z.string().describe('The checklist item ID'),
});

// Comments
export const listCommentsSchema = z.object({
  taskId: z.string().describe('The task ID'),
});

export const addCommentSchema = z.object({
  taskId: z.string().describe('The task ID'),
  comment: z.string().describe('The comment text'),
});

// References (attachments)
export const addPlannerTaskReferenceSchema = z.object({
  taskId: z.string().describe('The task ID'),
  url: z.string().url().describe('URL of the file or link to attach'),
  alias: z.string().optional().describe('Display name for the reference'),
  type: z
    .enum(['Word', 'Excel', 'PowerPoint', 'OneNote', 'Project', 'Visio', 'Pdf', 'TeamsHostedApp', 'Other'])
    .optional()
    .describe('Type of reference (auto-detected if not provided)'),
});

export const removePlannerTaskReferenceSchema = z.object({
  taskId: z.string().describe('The task ID'),
  url: z.string().url().describe('URL of the reference to remove'),
});

export const uploadAndAttachSchema = z.object({
  taskId: z.string().describe('The Planner task ID'),
  localPath: z.string().describe('Local file path to upload'),
  remotePath: z
    .string()
    .optional()
    .describe('Destination path in OneDrive. If omitted, uploads to "Planner Attachments/<filename>"'),
  alias: z.string().optional().describe('Display name for the attachment'),
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

// Reference type detection from URL
// Valid types per Microsoft API: Other, Word, Excel, PowerPoint, OneNote, Project, Visio, Pdf, TeamsHostedApp
function detectReferenceType(url: string): string {
  const lowerUrl = url.toLowerCase();
  if (lowerUrl.includes('.docx') || lowerUrl.includes('.doc')) return 'Word';
  if (lowerUrl.includes('.xlsx') || lowerUrl.includes('.xls')) return 'Excel';
  if (lowerUrl.includes('.pptx') || lowerUrl.includes('.ppt')) return 'PowerPoint';
  if (lowerUrl.includes('.one') || lowerUrl.includes('onenote')) return 'OneNote';
  if (lowerUrl.includes('.pdf')) return 'Pdf';
  if (lowerUrl.includes('.mpp')) return 'Project';
  if (lowerUrl.includes('.vsdx') || lowerUrl.includes('.vsd')) return 'Visio';
  // All other files (images, SharePoint URLs, OneDrive URLs, etc.) use 'Other'
  return 'Other';
}

// Encode URL for use as reference key (special encoding required by Planner)
// Only encode characters that are problematic as JSON object keys, preserve URL structure
function encodeReferenceUrl(url: string): string {
  // Microsoft Planner requires periods and some special chars to be encoded
  // but the URL must remain parseable (don't encode :, /, ?, &, =, etc.)
  return url
    .replace(/%/g, '%25') // Encode existing percent signs first
    .replace(/\./g, '%2E') // Periods must be encoded
    .replace(/#/g, '%23') // Hash/fragment
    .replace(/\s/g, '%20'); // Spaces
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
    const assignmentObj: Record<string, { '@odata.type': string; orderHint: string }> = {};
    for (const email of assignments) {
      const userId = await resolveUserId(email);
      assignmentObj[userId] = {
        '@odata.type': '#microsoft.graph.plannerAssignment',
        orderHint: ' !',
      };
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
  const { taskId, title, bucketId, assignments, clearAssignments, dueDateTime, clearDue, startDateTime, priority, progress } =
    params;

  // Fetch task to get ETag and current assignments (if clearing)
  const currentTask = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`);
  const etag = currentTask['@odata.etag'] as string;

  const body: Record<string, unknown> = {};

  if (title) body.title = title;
  if (bucketId) body.bucketId = bucketId;
  if (clearDue) {
    body.dueDateTime = null;
  } else if (dueDateTime) {
    body.dueDateTime = dueDateTime;
  }
  if (startDateTime) body.startDateTime = startDateTime;
  if (priority) body.priority = PRIORITY_MAP[priority];
  if (progress) body.percentComplete = PROGRESS_MAP[progress];

  // Handle assignments
  if (clearAssignments) {
    // Set all current assignments to null to remove them
    const assignmentObj: Record<string, null> = {};
    const currentAssignments = currentTask.assignments || {};
    for (const userId of Object.keys(currentAssignments)) {
      assignmentObj[userId] = null;
    }
    if (Object.keys(assignmentObj).length > 0) {
      body.assignments = assignmentObj;
    }
  } else if (assignments) {
    // Resolve assignments to user IDs
    const assignmentObj: Record<string, { '@odata.type': string; orderHint: string } | null> = {};
    for (const email of assignments) {
      const userId = await resolveUserId(email);
      assignmentObj[userId] = {
        '@odata.type': '#microsoft.graph.plannerAssignment',
        orderHint: ' !',
      };
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

// ============================================================================
// Checklist Operations
// ============================================================================

export async function addChecklistItem(params: z.infer<typeof addChecklistItemSchema>) {
  const { taskId, title, isChecked = false } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}/details`);

  // Generate a new item ID
  const itemId = crypto.randomUUID();

  const body = {
    checklist: {
      [itemId]: {
        '@odata.type': 'microsoft.graph.plannerChecklistItem',
        title,
        isChecked,
      },
    },
  };

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    message: 'Checklist item added',
    itemId,
    title,
    isChecked,
  };
}

export async function removeChecklistItem(params: z.infer<typeof removeChecklistItemSchema>) {
  const { taskId, itemId } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}/details`);

  // Set to null to remove
  const body = {
    checklist: {
      [itemId]: null,
    },
  };

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    message: 'Checklist item removed',
    itemId,
  };
}

export async function toggleChecklistItem(params: z.infer<typeof toggleChecklistItemSchema>) {
  const { taskId, itemId } = params;

  // Fetch current details to get the current state
  const details = await graphRequest<PlannerTaskDetails>(`/planner/tasks/${taskId}/details`);
  const etag = details['@odata.etag'];

  if (!etag) {
    throw new Error('Could not get ETag for task details');
  }

  const currentItem = details.checklist?.[itemId];
  if (!currentItem) {
    throw new Error(`Checklist item ${itemId} not found`);
  }

  const newCheckedState = !currentItem.isChecked;

  const body = {
    checklist: {
      [itemId]: {
        '@odata.type': 'microsoft.graph.plannerChecklistItem',
        title: currentItem.title,
        isChecked: newCheckedState,
      },
    },
  };

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    message: `Checklist item ${newCheckedState ? 'checked' : 'unchecked'}`,
    itemId,
    title: currentItem.title,
    isChecked: newCheckedState,
  };
}

// ============================================================================
// Comments
// ============================================================================

export async function listComments(params: z.infer<typeof listCommentsSchema>) {
  const { taskId } = params;

  // Get the task to find conversationThreadId and planId
  const task = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`);

  if (!task.conversationThreadId) {
    return { comments: [], message: 'No comments on this task' };
  }

  // Get the plan to find the group ID
  const plan = await graphRequest<PlannerPlan>(`/planner/plans/${task.planId}`);
  const groupId = plan.owner;

  // Get the conversation thread posts (comments)
  const posts = await graphList<ConversationPost>(
    `/groups/${groupId}/threads/${task.conversationThreadId}/posts`
  );

  return {
    comments: posts.map((post) => ({
      id: post.id,
      content: post.body.content,
      contentType: post.body.contentType,
      from: post.from.emailAddress.name || post.from.emailAddress.address,
      date: post.receivedDateTime,
    })),
  };
}

export async function addComment(params: z.infer<typeof addCommentSchema>) {
  const { taskId, comment } = params;

  // Get the task to find conversationThreadId and planId
  const task = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`);
  const etag = task['@odata.etag'] as string;

  // Get the plan to find the group ID
  const plan = await graphRequest<PlannerPlan>(`/planner/plans/${task.planId}`);
  const groupId = plan.owner;

  if (!task.conversationThreadId) {
    // No conversation thread yet - create one and link it to the task
    const newThread = await graphRequest<{ id: string }>(
      `/groups/${groupId}/threads`,
      {
        method: 'POST',
        body: {
          topic: task.title,
          posts: [
            {
              body: {
                contentType: 'text',
                content: comment,
              },
            },
          ],
        },
      }
    );

    // Link the new thread to the task
    await graphRequest(`/planner/tasks/${taskId}`, {
      method: 'PATCH',
      body: { conversationThreadId: newThread.id },
      headers: { 'If-Match': etag },
    });

    return {
      success: true,
      message: 'Comment added (conversation created)',
    };
  }

  // Reply to existing thread
  await graphRequest(
    `/groups/${groupId}/threads/${task.conversationThreadId}/reply`,
    {
      method: 'POST',
      body: {
        post: {
          body: {
            contentType: 'text',
            content: comment,
          },
        },
      },
    }
  );

  return {
    success: true,
    message: 'Comment added',
  };
}

// ============================================================================
// References (Attachments)
// ============================================================================

export async function addPlannerTaskReference(
  params: z.infer<typeof addPlannerTaskReferenceSchema>
) {
  const { taskId, url, alias, type } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}/details`);

  // Encode URL for use as object key
  const encodedUrl = encodeReferenceUrl(url);

  // Auto-detect type if not provided
  const referenceType = type || detectReferenceType(url);

  // Build reference object
  const reference: Record<string, unknown> = {
    '@odata.type': 'microsoft.graph.plannerExternalReference',
    type: referenceType,
    previewPriority: ' !',
  };
  if (alias) {
    reference.alias = alias;
  }

  const body = {
    references: {
      [encodedUrl]: reference,
    },
  };

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    message: 'Reference added to task',
    url,
    alias: alias || url,
    type: referenceType,
  };
}

export async function removePlannerTaskReference(
  params: z.infer<typeof removePlannerTaskReferenceSchema>
) {
  const { taskId, url } = params;

  // Fetch ETag first
  const etag = await getETag(`/planner/tasks/${taskId}/details`);

  // Encode URL for use as object key
  const encodedUrl = encodeReferenceUrl(url);

  // Set to null to remove
  const body = {
    references: {
      [encodedUrl]: null,
    },
  };

  await graphRequest(`/planner/tasks/${taskId}/details`, {
    method: 'PATCH',
    body,
    headers: { 'If-Match': etag },
  });

  return {
    success: true,
    message: 'Reference removed from task',
    url,
  };
}

export async function uploadAndAttach(params: z.infer<typeof uploadAndAttachSchema>) {
  const { taskId, localPath, remotePath, alias } = params;

  const filename = basename(localPath);

  // Step 1: Get task to find planId
  const task = await graphRequest<PlannerTask>(`/planner/tasks/${taskId}`);

  // Step 2: Get plan to find group (owner) and plan title
  const plan = await graphRequest<PlannerPlan>(`/planner/plans/${task.planId}`);
  const groupId = plan.owner;
  const planTitle = plan.title.replace(/[/\\?%*:|"<>]/g, '-'); // Sanitize for folder name

  // Step 3: Get the group's default drive (SharePoint document library)
  const drive = await graphRequest<Drive>(`/groups/${groupId}/drive`);

  // Step 4: Read local file
  const content = await readFile(localPath);

  // Step 5: Determine destination path in SharePoint
  const destPath = remotePath || `Planner Attachments/${planTitle}/${filename}`;

  // Step 6: Upload to group's SharePoint
  const MAX_SIMPLE_UPLOAD = 4 * 1024 * 1024;
  let uploaded: DriveItem;

  if (content.length <= MAX_SIMPLE_UPLOAD) {
    uploaded = await graphUpload<DriveItem>(
      `/drives/${drive.id}/root:/${destPath}:/content`,
      content
    );
  } else {
    uploaded = await graphUploadLarge<DriveItem>(
      `/drives/${drive.id}/root:/${destPath}`,
      content
    );
  }

  // Step 7: Attach the uploaded file's URL to the task
  const attachResult = await addPlannerTaskReference({
    taskId,
    url: uploaded.webUrl,
    alias: alias || filename,
  });

  return {
    success: true,
    message: 'File uploaded to SharePoint and attached to task',
    file: {
      name: uploaded.name,
      size: uploaded.size,
      webUrl: uploaded.webUrl,
    },
    attachment: {
      alias: alias || filename,
      type: attachResult.type,
    },
    location: {
      sharePointSite: drive.webUrl,
      path: destPath,
    },
  };
}
