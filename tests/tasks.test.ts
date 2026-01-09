/**
 * Integration tests for Tasks tools
 * Run with: npx tsx tests/tasks.test.ts
 */

import {
  listTaskLists,
  listTasks,
  createTask,
  completeTask,
  deleteTask,
} from '../src/tools/tasks.js';

async function test(name: string, fn: () => Promise<void>) {
  try {
    await fn();
    console.log(`✅ ${name}`);
  } catch (error) {
    console.log(`❌ ${name}`);
    console.error(`   ${error instanceof Error ? error.message : error}`);
  }
}

async function runTests() {
  console.log('\n=== Tasks Integration Tests ===\n');

  let testListId: string | undefined;
  let testTaskId: string | undefined;

  // Test 1: List task lists
  await test('listTaskLists', async () => {
    const lists = await listTaskLists();
    if (!Array.isArray(lists) || lists.length === 0) {
      throw new Error('Expected at least one task list');
    }
    testListId = lists[0].id;
    console.log(`   Found ${lists.length} lists, using: ${lists[0].name}`);
  });

  // Test 2: List tasks (no params - uses default list)
  await test('listTasks (default list)', async () => {
    const tasks = await listTasks({});
    if (!Array.isArray(tasks)) {
      throw new Error('Expected array of tasks');
    }
    console.log(`   Found ${tasks.length} tasks`);
  });

  // Test 3: List tasks with explicit listId
  await test('listTasks (explicit listId)', async () => {
    if (!testListId) throw new Error('No list ID from previous test');
    const tasks = await listTasks({ listId: testListId });
    if (!Array.isArray(tasks)) {
      throw new Error('Expected array of tasks');
    }
    console.log(`   Found ${tasks.length} tasks`);
  });

  // Test 4: List tasks including completed
  await test('listTasks (includeCompleted)', async () => {
    const tasks = await listTasks({ includeCompleted: true });
    if (!Array.isArray(tasks)) {
      throw new Error('Expected array of tasks');
    }
    console.log(`   Found ${tasks.length} tasks (including completed)`);
  });

  // Test 5: Create a task
  await test('createTask', async () => {
    const result = await createTask({
      title: `Test task ${Date.now()}`,
      importance: 'low',
    });
    if (!result.success || !result.taskId) {
      throw new Error('Failed to create task');
    }
    testTaskId = result.taskId;
    console.log(`   Created task: ${result.title}`);
  });

  // Test 6: Complete the task
  await test('completeTask', async () => {
    if (!testTaskId) throw new Error('No task ID from previous test');
    const result = await completeTask({ taskId: testTaskId });
    if (!result.success) {
      throw new Error('Failed to complete task');
    }
  });

  // Test 7: Delete the task (cleanup)
  await test('deleteTask', async () => {
    if (!testTaskId) throw new Error('No task ID from previous test');
    const result = await deleteTask({ taskId: testTaskId });
    if (!result.success) {
      throw new Error('Failed to delete task');
    }
  });

  console.log('\n=== Tests Complete ===\n');
}

runTests().catch(console.error);
