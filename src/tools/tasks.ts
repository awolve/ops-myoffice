import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface TaskList {
  id: string;
  displayName: string;
}

interface Task {
  id: string;
  title: string;
  status: 'notStarted' | 'inProgress' | 'completed';
  importance: 'low' | 'normal' | 'high';
  createdDateTime: string;
  dueDateTime?: { dateTime: string; timeZone: string };
  body?: { content: string; contentType: string };
  completedDateTime?: { dateTime: string; timeZone: string };
}

// Schemas
export const listTaskListsSchema = z.object({});

export const listTasksSchema = z.object({
  listId: z.string().optional().describe('Task list ID. If not provided, uses the default task list'),
  includeCompleted: z.boolean().optional().describe('Include completed tasks. Default: false'),
  maxItems: z.number().optional().describe('Maximum number of tasks. Default: 50'),
});

export const createTaskSchema = z.object({
  title: z.string().describe('Task title'),
  listId: z.string().optional().describe('Task list ID. Default: default list'),
  dueDate: z.string().optional().describe('Due date (ISO format)'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('Task importance'),
  body: z.string().optional().describe('Task notes/description'),
});

export const updateTaskSchema = z.object({
  taskId: z.string().describe('The ID of the task to update'),
  listId: z.string().optional().describe('Task list ID. Default: default list'),
  title: z.string().optional().describe('New title'),
  dueDate: z.string().optional().describe('New due date'),
  importance: z.enum(['low', 'normal', 'high']).optional().describe('New importance'),
  body: z.string().optional().describe('New notes'),
});

export const completeTaskSchema = z.object({
  taskId: z.string().describe('The ID of the task to complete'),
  listId: z.string().optional().describe('Task list ID. Default: default list'),
});

export const deleteTaskSchema = z.object({
  taskId: z.string().describe('The ID of the task to delete'),
  listId: z.string().optional().describe('Task list ID. Default: default list'),
});

// Helper to get default task list
async function getDefaultListId(): Promise<string> {
  const lists = await graphList<TaskList>('/me/todo/lists');
  const defaultList = lists.find((l) => l.displayName === 'Tasks') || lists[0];
  if (!defaultList) {
    throw new Error('No task lists found');
  }
  return defaultList.id;
}

// Tool implementations
export async function listTaskLists() {
  const lists = await graphList<TaskList>('/me/todo/lists');
  return lists.map((l) => ({
    id: l.id,
    name: l.displayName,
  }));
}

export async function listTasks(params: z.infer<typeof listTasksSchema>) {
  const { listId, includeCompleted = false, maxItems = 50 } = params;

  const actualListId = listId || (await getDefaultListId());

  // Note: To Do API doesn't support combining $orderby with $filter well
  // So we fetch all and filter client-side if needed
  const path = `/me/todo/lists/${actualListId}/tasks?$select=id,title,status,importance,dueDateTime,createdDateTime&$top=${maxItems}`;

  let tasks = await graphList<Task>(path, { maxItems });

  // Filter completed tasks client-side
  if (!includeCompleted) {
    tasks = tasks.filter((t) => t.status !== 'completed');
  }

  return tasks.map((t) => ({
    id: t.id,
    title: t.title,
    status: t.status,
    importance: t.importance,
    dueDate: t.dueDateTime?.dateTime,
    created: t.createdDateTime,
  }));
}

export async function createTask(params: z.infer<typeof createTaskSchema>) {
  const { title, listId, dueDate, importance = 'normal', body } = params;

  const actualListId = listId || (await getDefaultListId());

  const taskData: Record<string, unknown> = {
    title,
    importance,
  };

  if (dueDate) {
    taskData.dueDateTime = { dateTime: dueDate, timeZone: 'UTC' };
  }

  if (body) {
    taskData.body = { content: body, contentType: 'text' };
  }

  const created = await graphRequest<Task>(`/me/todo/lists/${actualListId}/tasks`, {
    method: 'POST',
    body: taskData,
  });

  return {
    success: true,
    taskId: created.id,
    title: created.title,
  };
}

export async function updateTask(params: z.infer<typeof updateTaskSchema>) {
  const { taskId, listId, title, dueDate, importance, body } = params;

  const actualListId = listId || (await getDefaultListId());

  const updates: Record<string, unknown> = {};
  if (title) updates.title = title;
  if (importance) updates.importance = importance;
  if (dueDate) updates.dueDateTime = { dateTime: dueDate, timeZone: 'UTC' };
  if (body) updates.body = { content: body, contentType: 'text' };

  const updated = await graphRequest<Task>(`/me/todo/lists/${actualListId}/tasks/${taskId}`, {
    method: 'PATCH',
    body: updates,
  });

  return {
    success: true,
    taskId: updated.id,
    title: updated.title,
  };
}

export async function completeTask(params: z.infer<typeof completeTaskSchema>) {
  const { taskId, listId } = params;

  const actualListId = listId || (await getDefaultListId());

  await graphRequest(`/me/todo/lists/${actualListId}/tasks/${taskId}`, {
    method: 'PATCH',
    body: { status: 'completed' },
  });

  return { success: true, message: 'Task completed' };
}

export async function deleteTask(params: z.infer<typeof deleteTaskSchema>) {
  const { taskId, listId } = params;

  const actualListId = listId || (await getDefaultListId());

  await graphRequest(`/me/todo/lists/${actualListId}/tasks/${taskId}`, {
    method: 'DELETE',
  });

  return { success: true, message: 'Task deleted' };
}
