#!/usr/bin/env node

import { Command } from 'commander';
import { getVersion } from './utils/version.js';
import { executeCommand } from './core/handler.js';
import { isAuthenticated, authenticateWithDeviceCode } from './auth/index.js';
import { formatOutput } from './cli/formatter.js';

const program = new Command();

// Global state for JSON output mode
let jsonOutput = false;
let currentToolName: string | undefined;

function output(data: unknown): void {
  if (jsonOutput) {
    console.log(JSON.stringify(data, null, 2));
  } else {
    console.log(formatOutput(data, currentToolName));
  }
}

function outputError(message: string, code?: string): void {
  if (jsonOutput) {
    console.error(JSON.stringify({ error: message, code: code || 'ERROR' }));
  } else {
    console.error(`Error: ${message}`);
  }
  process.exit(1);
}

async function requireAuth(): Promise<void> {
  const authenticated = await isAuthenticated();
  if (!authenticated) {
    outputError('Not authenticated. Run: myoffice login', 'AUTH_REQUIRED');
  }
}

async function runCommand(toolName: string, args: Record<string, unknown> = {}): Promise<void> {
  await requireAuth();

  currentToolName = toolName;
  const result = await executeCommand(toolName, args);

  if (result.success) {
    output(result.data);
  } else {
    outputError(result.error || 'Unknown error', result.code);
  }
}

// Main program setup
program
  .name('myoffice')
  .description('Access your Microsoft 365 data from the command line')
  .version(getVersion())
  .option('--json', 'Output as JSON')
  .hook('preAction', (thisCommand) => {
    const opts = thisCommand.opts();
    jsonOutput = opts.json || false;
  });

// Login command
program
  .command('login')
  .description('Authenticate with Microsoft 365')
  .action(async () => {
    try {
      console.log('Starting authentication...');
      console.log('');
      await authenticateWithDeviceCode();
      console.log('');
      console.log('Authentication successful!');
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      outputError(`Authentication failed: ${message}`, 'AUTH_FAILED');
    }
  });

// Status command
program
  .command('status')
  .description('Check authentication status')
  .action(async () => {
    currentToolName = 'auth_status';
    const result = await executeCommand('auth_status', {});
    if (result.success) {
      output(result.data);
    } else {
      outputError(result.error || 'Failed to get status', result.code);
    }
  });

// Debug command
program
  .command('debug')
  .description('Show debug information')
  .action(async () => {
    currentToolName = 'debug_info';
    const result = await executeCommand('debug_info', {});
    if (result.success) {
      output(result.data);
    } else {
      outputError(result.error || 'Failed to get debug info', result.code);
    }
  });

// Mail commands
const mailCmd = program
  .command('mail')
  .description('Email operations (list, read, send, reply, search, delete, mark)');

mailCmd
  .command('list')
  .description('List emails from a folder')
  .option('--folder <name>', 'Folder name (default: inbox)')
  .option('--limit <n>', 'Maximum emails to return', '25')
  .option('--unread', 'Only show unread emails')
  .action(async (opts) => {
    await runCommand('mail_list', {
      folder: opts.folder,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
      unreadOnly: opts.unread || undefined,
    });
  });

mailCmd
  .command('read')
  .description('Read a specific email')
  .requiredOption('--id <messageId>', 'The message ID')
  .action(async (opts) => {
    await runCommand('mail_read', { messageId: opts.id });
  });

mailCmd
  .command('search')
  .description('Search emails')
  .requiredOption('--query <query>', 'Search query')
  .option('--limit <n>', 'Maximum results', '25')
  .action(async (opts) => {
    await runCommand('mail_search', {
      query: opts.query,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

mailCmd
  .command('send')
  .description('Send an email')
  .requiredOption('--to <emails...>', 'Recipient emails')
  .requiredOption('--subject <subject>', 'Email subject')
  .requiredOption('--body <body>', 'Email body')
  .option('--cc <emails...>', 'CC recipients')
  .option('--bcc <emails...>', 'BCC recipients')
  .option('--no-html', 'Send as plain text')
  .option('--no-signature', 'Do not append signature')
  .action(async (opts) => {
    await runCommand('mail_send', {
      to: opts.to,
      subject: opts.subject,
      body: opts.body,
      cc: opts.cc,
      bcc: opts.bcc,
      isHtml: opts.html !== false,
      useSignature: opts.signature !== false,
    });
  });

mailCmd
  .command('reply')
  .description('Reply to an email')
  .requiredOption('--id <messageId>', 'The message ID to reply to')
  .requiredOption('--body <body>', 'Reply body')
  .option('--all', 'Reply to all recipients')
  .option('--no-html', 'Send as plain text')
  .option('--signature', 'Append signature')
  .action(async (opts) => {
    await runCommand('mail_reply', {
      messageId: opts.id,
      body: opts.body,
      replyAll: opts.all || false,
      isHtml: opts.html !== false,
      useSignature: opts.signature || false,
    });
  });

mailCmd
  .command('delete')
  .description('Delete an email')
  .requiredOption('--id <messageId>', 'The message ID to delete')
  .action(async (opts) => {
    await runCommand('mail_delete', { messageId: opts.id });
  });

mailCmd
  .command('mark')
  .description('Mark an email as read or unread')
  .requiredOption('--id <messageId>', 'The message ID')
  .option('--unread', 'Mark as unread (default: mark as read)')
  .action(async (opts) => {
    await runCommand('mail_mark_read', {
      messageId: opts.id,
      isRead: !opts.unread,
    });
  });

// Calendar commands
const calendarCmd = program
  .command('calendar')
  .description('Calendar events (list, get, create, update, delete)');

calendarCmd
  .command('list')
  .description('List calendar events')
  .option('--start <date>', 'Start date (ISO format)')
  .option('--end <date>', 'End date (ISO format)')
  .option('--limit <n>', 'Maximum events', '50')
  .action(async (opts) => {
    await runCommand('calendar_list', {
      startDate: opts.start,
      endDate: opts.end,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

calendarCmd
  .command('get')
  .description('Get details of a calendar event')
  .requiredOption('--id <eventId>', 'The event ID')
  .action(async (opts) => {
    await runCommand('calendar_get', { eventId: opts.id });
  });

calendarCmd
  .command('create')
  .description('Create a calendar event')
  .requiredOption('--subject <subject>', 'Event title')
  .requiredOption('--start <datetime>', 'Start datetime (ISO)')
  .requiredOption('--end <datetime>', 'End datetime (ISO)')
  .option('--timezone <tz>', 'Timezone (default: UTC)')
  .option('--location <location>', 'Event location')
  .option('--body <body>', 'Event description')
  .option('--attendees <emails...>', 'Attendee emails')
  .option('--online', 'Create Teams meeting')
  .action(async (opts) => {
    await runCommand('calendar_create', {
      subject: opts.subject,
      start: opts.start,
      end: opts.end,
      timeZone: opts.timezone,
      location: opts.location,
      body: opts.body,
      attendees: opts.attendees,
      isOnlineMeeting: opts.online || false,
    });
  });

calendarCmd
  .command('update')
  .description('Update a calendar event')
  .requiredOption('--id <eventId>', 'The event ID')
  .option('--subject <subject>', 'New title')
  .option('--start <datetime>', 'New start datetime')
  .option('--end <datetime>', 'New end datetime')
  .option('--location <location>', 'New location')
  .option('--body <body>', 'New description')
  .action(async (opts) => {
    await runCommand('calendar_update', {
      eventId: opts.id,
      subject: opts.subject,
      start: opts.start,
      end: opts.end,
      location: opts.location,
      body: opts.body,
    });
  });

calendarCmd
  .command('delete')
  .description('Delete a calendar event')
  .requiredOption('--id <eventId>', 'The event ID to delete')
  .action(async (opts) => {
    await runCommand('calendar_delete', { eventId: opts.id });
  });

// Tasks commands
const tasksCmd = program
  .command('tasks')
  .description('Microsoft To Do (lists, list, create, update, complete, delete)');

tasksCmd
  .command('lists')
  .description('List all task lists')
  .action(async () => {
    await runCommand('tasks_list_lists', {});
  });

tasksCmd
  .command('list')
  .description('List tasks from a task list')
  .option('--list-id <listId>', 'Task list ID (default: default list)')
  .option('--include-completed', 'Include completed tasks')
  .option('--limit <n>', 'Maximum tasks', '50')
  .action(async (opts) => {
    await runCommand('tasks_list', {
      listId: opts.listId,
      includeCompleted: opts.includeCompleted || false,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

tasksCmd
  .command('create')
  .description('Create a new task')
  .requiredOption('--title <title>', 'Task title')
  .option('--list-id <listId>', 'Task list ID')
  .option('--due <date>', 'Due date (ISO)')
  .option('--importance <level>', 'Importance: low, normal, high')
  .option('--body <body>', 'Task notes')
  .action(async (opts) => {
    await runCommand('tasks_create', {
      title: opts.title,
      listId: opts.listId,
      dueDate: opts.due,
      importance: opts.importance,
      body: opts.body,
    });
  });

tasksCmd
  .command('update')
  .description('Update a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .option('--list-id <listId>', 'Task list ID')
  .option('--title <title>', 'New title')
  .option('--due <date>', 'New due date')
  .option('--importance <level>', 'New importance: low, normal, high')
  .option('--body <body>', 'New notes')
  .action(async (opts) => {
    await runCommand('tasks_update', {
      taskId: opts.id,
      listId: opts.listId,
      title: opts.title,
      dueDate: opts.due,
      importance: opts.importance,
      body: opts.body,
    });
  });

tasksCmd
  .command('complete')
  .description('Mark a task as completed')
  .requiredOption('--id <taskId>', 'Task ID')
  .option('--list-id <listId>', 'Task list ID')
  .action(async (opts) => {
    await runCommand('tasks_complete', {
      taskId: opts.id,
      listId: opts.listId,
    });
  });

tasksCmd
  .command('delete')
  .description('Delete a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .option('--list-id <listId>', 'Task list ID')
  .action(async (opts) => {
    await runCommand('tasks_delete', {
      taskId: opts.id,
      listId: opts.listId,
    });
  });

// Files (OneDrive) commands
const filesCmd = program
  .command('files')
  .description('OneDrive files (list, get, search, read, mkdir, shared, upload)');

filesCmd
  .command('list')
  .description('List files and folders')
  .option('--path <path>', 'Folder path (default: root)')
  .option('--limit <n>', 'Maximum items', '50')
  .action(async (opts) => {
    await runCommand('onedrive_list', {
      path: opts.path,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

filesCmd
  .command('get')
  .description('Get metadata for a file or folder')
  .requiredOption('--path <path>', 'File/folder path')
  .action(async (opts) => {
    await runCommand('onedrive_get', { path: opts.path });
  });

filesCmd
  .command('search')
  .description('Search for files')
  .requiredOption('--query <query>', 'Search query')
  .option('--limit <n>', 'Maximum results', '25')
  .action(async (opts) => {
    await runCommand('onedrive_search', {
      query: opts.query,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

filesCmd
  .command('read')
  .description('Read content of a text file')
  .requiredOption('--path <path>', 'File path')
  .action(async (opts) => {
    await runCommand('onedrive_read', { path: opts.path });
  });

filesCmd
  .command('mkdir')
  .description('Create a new folder')
  .requiredOption('--name <name>', 'Folder name')
  .option('--parent <path>', 'Parent path (default: root)')
  .action(async (opts) => {
    await runCommand('onedrive_create_folder', {
      name: opts.name,
      parentPath: opts.parent,
    });
  });

filesCmd
  .command('shared')
  .description('List files shared with you')
  .option('--limit <n>', 'Maximum items', '50')
  .action(async (opts) => {
    await runCommand('onedrive_shared_with_me', {
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

filesCmd
  .command('upload')
  .description('Upload a local file to OneDrive')
  .requiredOption('--file <path>', 'Local file path to upload')
  .option('--dest <path>', 'Destination path in OneDrive')
  .action(async (opts) => {
    await runCommand('onedrive_upload', {
      localPath: opts.file,
      remotePath: opts.dest,
    });
  });

// SharePoint commands
const sharepointCmd = program
  .command('sharepoint')
  .description('SharePoint sites and document libraries');

sharepointCmd
  .command('sites')
  .description('List SharePoint sites')
  .option('--search <query>', 'Search query to find sites')
  .option('--limit <n>', 'Maximum sites', '50')
  .action(async (opts) => {
    await runCommand('sharepoint_list_sites', {
      search: opts.search,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

sharepointCmd
  .command('site')
  .description('Get details of a SharePoint site')
  .requiredOption('--id <siteId>', 'Site ID or hostname:path')
  .action(async (opts) => {
    await runCommand('sharepoint_get_site', { siteId: opts.id });
  });

sharepointCmd
  .command('drives')
  .description('List document libraries in a site')
  .requiredOption('--site-id <siteId>', 'Site ID')
  .option('--limit <n>', 'Maximum drives', '50')
  .action(async (opts) => {
    await runCommand('sharepoint_list_drives', {
      siteId: opts.siteId,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

sharepointCmd
  .command('files')
  .description('List files in a document library')
  .requiredOption('--drive-id <driveId>', 'Drive ID')
  .option('--path <path>', 'Folder path (default: root)')
  .option('--limit <n>', 'Maximum items', '50')
  .action(async (opts) => {
    await runCommand('sharepoint_list_files', {
      driveId: opts.driveId,
      path: opts.path,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

sharepointCmd
  .command('file')
  .description('Get metadata for a file')
  .requiredOption('--drive-id <driveId>', 'Drive ID')
  .requiredOption('--path <path>', 'File path')
  .action(async (opts) => {
    await runCommand('sharepoint_get_file', {
      driveId: opts.driveId,
      path: opts.path,
    });
  });

sharepointCmd
  .command('read')
  .description('Read content of a text file')
  .requiredOption('--drive-id <driveId>', 'Drive ID')
  .requiredOption('--path <path>', 'File path')
  .action(async (opts) => {
    await runCommand('sharepoint_read_file', {
      driveId: opts.driveId,
      path: opts.path,
    });
  });

sharepointCmd
  .command('search')
  .description('Search for files in a document library')
  .requiredOption('--drive-id <driveId>', 'Drive ID')
  .requiredOption('--query <query>', 'Search query')
  .option('--limit <n>', 'Maximum results', '25')
  .action(async (opts) => {
    await runCommand('sharepoint_search_files', {
      driveId: opts.driveId,
      query: opts.query,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

// Contacts commands
const contactsCmd = program
  .command('contacts')
  .description('Contacts (list, search, get, create, update)');

contactsCmd
  .command('list')
  .description('List contacts')
  .option('--limit <n>', 'Maximum contacts', '50')
  .action(async (opts) => {
    await runCommand('contacts_list', {
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

contactsCmd
  .command('search')
  .description('Search contacts')
  .requiredOption('--query <query>', 'Search query')
  .option('--limit <n>', 'Maximum results', '25')
  .action(async (opts) => {
    await runCommand('contacts_search', {
      query: opts.query,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

contactsCmd
  .command('get')
  .description('Get details of a contact')
  .requiredOption('--id <contactId>', 'Contact ID')
  .action(async (opts) => {
    await runCommand('contacts_get', { contactId: opts.id });
  });

contactsCmd
  .command('create')
  .description('Create a new contact')
  .option('--given-name <name>', 'First name')
  .option('--surname <name>', 'Last name')
  .option('--email <email>', 'Email address')
  .option('--mobile <phone>', 'Mobile phone')
  .option('--business-phone <phone>', 'Business phone')
  .option('--company <name>', 'Company name')
  .option('--job-title <title>', 'Job title')
  .option('--notes <text>', 'Personal notes about the contact')
  .action(async (opts) => {
    await runCommand('contacts_create', {
      givenName: opts.givenName,
      surname: opts.surname,
      email: opts.email,
      mobilePhone: opts.mobile,
      businessPhone: opts.businessPhone,
      companyName: opts.company,
      jobTitle: opts.jobTitle,
      notes: opts.notes,
    });
  });

contactsCmd
  .command('update')
  .description('Update an existing contact')
  .requiredOption('--id <contactId>', 'Contact ID')
  .option('--given-name <name>', 'First name')
  .option('--surname <name>', 'Last name')
  .option('--email <email>', 'Email address')
  .option('--mobile <phone>', 'Mobile phone')
  .option('--business-phone <phone>', 'Business phone')
  .option('--company <name>', 'Company name')
  .option('--job-title <title>', 'Job title')
  .option('--notes <text>', 'Personal notes about the contact')
  .action(async (opts) => {
    await runCommand('contacts_update', {
      contactId: opts.id,
      givenName: opts.givenName,
      surname: opts.surname,
      email: opts.email,
      mobilePhone: opts.mobile,
      businessPhone: opts.businessPhone,
      companyName: opts.company,
      jobTitle: opts.jobTitle,
      notes: opts.notes,
    });
  });

// Teams commands
const teamsCmd = program
  .command('teams')
  .description('Teams channels and messages');

teamsCmd
  .command('list')
  .description('List Teams you are a member of')
  .option('--limit <n>', 'Maximum teams', '50')
  .action(async (opts) => {
    await runCommand('teams_list', {
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

teamsCmd
  .command('channels')
  .description('List channels in a Team')
  .requiredOption('--team-id <teamId>', 'Team ID')
  .action(async (opts) => {
    await runCommand('teams_channels', { teamId: opts.teamId });
  });

teamsCmd
  .command('messages')
  .description('Read messages from a channel')
  .requiredOption('--team-id <teamId>', 'Team ID')
  .requiredOption('--channel-id <channelId>', 'Channel ID')
  .option('--limit <n>', 'Maximum messages', '25')
  .action(async (opts) => {
    await runCommand('teams_channel_messages', {
      teamId: opts.teamId,
      channelId: opts.channelId,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

teamsCmd
  .command('post')
  .description('Post a message to a channel')
  .requiredOption('--team-id <teamId>', 'Team ID')
  .requiredOption('--channel-id <channelId>', 'Channel ID')
  .requiredOption('--content <content>', 'Message content')
  .action(async (opts) => {
    await runCommand('teams_channel_post', {
      teamId: opts.teamId,
      channelId: opts.channelId,
      content: opts.content,
    });
  });

// Chats commands
const chatsCmd = program
  .command('chats')
  .description('1:1 and group chats');

chatsCmd
  .command('list')
  .description('List your chats')
  .option('--limit <n>', 'Maximum chats', '25')
  .action(async (opts) => {
    await runCommand('chats_list', {
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

chatsCmd
  .command('messages')
  .description('Read messages from a chat')
  .requiredOption('--chat-id <chatId>', 'Chat ID')
  .option('--limit <n>', 'Maximum messages', '25')
  .action(async (opts) => {
    await runCommand('chats_messages', {
      chatId: opts.chatId,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

chatsCmd
  .command('send')
  .description('Send a message in a chat')
  .requiredOption('--chat-id <chatId>', 'Chat ID')
  .requiredOption('--content <content>', 'Message content')
  .action(async (opts) => {
    await runCommand('chats_send', {
      chatId: opts.chatId,
      content: opts.content,
    });
  });

chatsCmd
  .command('create')
  .description('Create a new chat')
  .requiredOption('--members <emails...>', 'Email addresses of chat members')
  .option('--topic <topic>', 'Chat topic/title (for group chats)')
  .action(async (opts) => {
    await runCommand('chats_create', {
      members: opts.members,
      topic: opts.topic,
    });
  });

// Planner commands
const plannerCmd = program
  .command('planner')
  .description('Planner plans, buckets, tasks, and attachments');

plannerCmd
  .command('plans')
  .description('List all Planner plans')
  .option('--limit <n>', 'Maximum plans', '50')
  .action(async (opts) => {
    await runCommand('planner_list_plans', {
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

plannerCmd
  .command('plan')
  .description('Get details of a plan')
  .requiredOption('--id <planId>', 'Plan ID')
  .action(async (opts) => {
    await runCommand('planner_get_plan', { planId: opts.id });
  });

plannerCmd
  .command('buckets')
  .description('List buckets in a plan')
  .requiredOption('--plan-id <planId>', 'Plan ID')
  .action(async (opts) => {
    await runCommand('planner_list_buckets', { planId: opts.planId });
  });

plannerCmd
  .command('bucket-create')
  .description('Create a new bucket')
  .requiredOption('--plan-id <planId>', 'Plan ID')
  .requiredOption('--name <name>', 'Bucket name')
  .action(async (opts) => {
    await runCommand('planner_create_bucket', {
      planId: opts.planId,
      name: opts.name,
    });
  });

plannerCmd
  .command('bucket-update')
  .description('Update a bucket')
  .requiredOption('--id <bucketId>', 'Bucket ID')
  .requiredOption('--name <name>', 'New bucket name')
  .action(async (opts) => {
    await runCommand('planner_update_bucket', {
      bucketId: opts.id,
      name: opts.name,
    });
  });

plannerCmd
  .command('bucket-delete')
  .description('Delete a bucket')
  .requiredOption('--id <bucketId>', 'Bucket ID')
  .action(async (opts) => {
    await runCommand('planner_delete_bucket', { bucketId: opts.id });
  });

plannerCmd
  .command('tasks')
  .description('List tasks in a plan')
  .requiredOption('--plan-id <planId>', 'Plan ID')
  .option('--bucket-id <bucketId>', 'Filter by bucket')
  .option('--limit <n>', 'Maximum tasks', '100')
  .action(async (opts) => {
    await runCommand('planner_list_tasks', {
      planId: opts.planId,
      bucketId: opts.bucketId,
      maxItems: opts.limit ? parseInt(opts.limit, 10) : undefined,
    });
  });

plannerCmd
  .command('task')
  .description('Get details of a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .action(async (opts) => {
    await runCommand('planner_get_task', { taskId: opts.id });
  });

plannerCmd
  .command('task-create')
  .description('Create a new task')
  .requiredOption('--plan-id <planId>', 'Plan ID')
  .requiredOption('--title <title>', 'Task title')
  .option('--bucket-id <bucketId>', 'Bucket ID')
  .option('--assignments <emails...>', 'Assign to users (emails)')
  .option('--due <datetime>', 'Due date (ISO)')
  .option('--start <datetime>', 'Start date (ISO)')
  .option('--priority <level>', 'Priority: urgent, important, medium, low')
  .option('--progress <status>', 'Progress: notStarted, inProgress, completed')
  .action(async (opts) => {
    await runCommand('planner_create_task', {
      planId: opts.planId,
      title: opts.title,
      bucketId: opts.bucketId,
      assignments: opts.assignments,
      dueDateTime: opts.due,
      startDateTime: opts.start,
      priority: opts.priority,
      progress: opts.progress,
    });
  });

plannerCmd
  .command('task-update')
  .description('Update a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .option('--title <title>', 'New title')
  .option('--bucket-id <bucketId>', 'Move to bucket')
  .option('--assignments <emails...>', 'Assign to users (replaces existing)')
  .option('--due <datetime>', 'New due date')
  .option('--clear-due', 'Remove due date')
  .option('--start <datetime>', 'New start date')
  .option('--priority <level>', 'New priority')
  .option('--progress <status>', 'New progress')
  .action(async (opts) => {
    await runCommand('planner_update_task', {
      taskId: opts.id,
      title: opts.title,
      bucketId: opts.bucketId,
      assignments: opts.assignments,
      dueDateTime: opts.due,
      clearDue: opts.clearDue,
      startDateTime: opts.start,
      priority: opts.priority,
      progress: opts.progress,
    });
  });

plannerCmd
  .command('task-delete')
  .description('Delete a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .action(async (opts) => {
    await runCommand('planner_delete_task', { taskId: opts.id });
  });

plannerCmd
  .command('task-details')
  .description('Get task details (description, checklist)')
  .requiredOption('--id <taskId>', 'Task ID')
  .action(async (opts) => {
    await runCommand('planner_get_task_details', { taskId: opts.id });
  });

plannerCmd
  .command('task-details-update')
  .description('Update task description and checklist')
  .requiredOption('--id <taskId>', 'Task ID')
  .option('--description <text>', 'Task description')
  .action(async (opts) => {
    await runCommand('planner_update_task_details', {
      taskId: opts.id,
      description: opts.description,
    });
  });

plannerCmd
  .command('attach')
  .description('Add a link/attachment to a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .requiredOption('--url <url>', 'URL to attach (file URL or web link)')
  .option('--alias <name>', 'Display name for the attachment')
  .option(
    '--type <type>',
    'File type: Word, Excel, PowerPoint, OneNote, SharePoint, OneDrive, Pdf, Other'
  )
  .action(async (opts) => {
    await runCommand('planner_add_reference', {
      taskId: opts.id,
      url: opts.url,
      alias: opts.alias,
      type: opts.type,
    });
  });

plannerCmd
  .command('detach')
  .description('Remove an attachment/link from a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .requiredOption('--url <url>', 'URL to remove')
  .action(async (opts) => {
    await runCommand('planner_remove_reference', {
      taskId: opts.id,
      url: opts.url,
    });
  });

plannerCmd
  .command('upload')
  .description('Upload a local file and attach it to a task')
  .requiredOption('--id <taskId>', 'Task ID')
  .requiredOption('--file <path>', 'Local file path to upload')
  .option('--dest <path>', 'Destination path in OneDrive (default: "Planner Attachments/<filename>")')
  .option('--alias <name>', 'Display name for the attachment')
  .action(async (opts) => {
    await runCommand('planner_upload_attach', {
      taskId: opts.id,
      localPath: opts.file,
      remotePath: opts.dest,
      alias: opts.alias,
    });
  });

// Parse and execute
program.parse();
