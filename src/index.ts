#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { getVersion } from './utils/version.js';
import { executeCommand } from './core/handler.js';

const VERSION = getVersion();
const SERVER_START_TIME = new Date();

// Debug: Log environment on startup
console.error('=== MCP Server Starting ===');
console.error('Version:', VERSION);
console.error('Started at:', SERVER_START_TIME.toISOString());
const clientId = process.env.M365_CLIENT_ID;
if (!clientId) {
  console.error('M365_CLIENT_ID: NOT SET');
} else if (clientId.startsWith('${') || clientId.includes('$M365')) {
  console.error('M365_CLIENT_ID: UNRESOLVED PLACEHOLDER:', clientId);
} else {
  console.error('M365_CLIENT_ID: SET (' + clientId.substring(0, 8) + '...)');
}
console.error('M365_TENANT_ID:', process.env.M365_TENANT_ID || 'NOT SET (will use "common")');
console.error('===========================');

const server = new Server(
  {
    name: 'myoffice-mcp',
    version: VERSION,
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Tool definitions
const TOOLS = [
  // Mail
  {
    name: 'mail_list',
    description: 'List emails from a folder (inbox, sentitems, drafts, etc.)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        folder: { type: 'string', description: 'Folder name. Default: inbox' },
        maxItems: { type: 'number', description: 'Max emails to return. Default: 25' },
        unreadOnly: { type: 'boolean', description: 'Only return unread emails' },
      },
    },
  },
  {
    name: 'mail_read',
    description: 'Read a specific email by ID',
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: { type: 'string', description: 'The message ID' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'mail_search',
    description: 'Search emails by query',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: { type: 'string', description: 'Search query' },
        maxItems: { type: 'number', description: 'Max results. Default: 25' },
      },
      required: ['query'],
    },
  },
  {
    name: 'mail_send',
    description: 'Send an email. IMPORTANT: Requires user confirmation before sending. Automatically appends signature if configured.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        to: { type: 'array', items: { type: 'string' }, description: 'Recipient emails' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body (HTML by default)' },
        isHtml: { type: 'boolean', description: 'Is body HTML? Default: true' },
        cc: { type: 'array', items: { type: 'string' }, description: 'CC recipients' },
        bcc: { type: 'array', items: { type: 'string' }, description: 'BCC recipients' },
        useSignature: { type: 'boolean', description: 'Append email signature if configured. Default: true' },
      },
      required: ['to', 'subject', 'body'],
    },
  },
  {
    name: 'mail_reply',
    description: 'Reply to an email. IMPORTANT: Requires user confirmation before sending.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: { type: 'string', description: 'The message ID to reply to' },
        body: { type: 'string', description: 'Reply body (HTML by default)' },
        isHtml: { type: 'boolean', description: 'Is body HTML? Default: true' },
        replyAll: { type: 'boolean', description: 'Reply to all recipients. Default: false' },
        useSignature: { type: 'boolean', description: 'Append email signature. Default: false' },
      },
      required: ['messageId', 'body'],
    },
  },
  {
    name: 'mail_delete',
    description: 'Delete an email. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: { type: 'string', description: 'The message ID to delete' },
      },
      required: ['messageId'],
    },
  },
  {
    name: 'mail_mark_read',
    description: 'Mark an email as read or unread',
    inputSchema: {
      type: 'object' as const,
      properties: {
        messageId: { type: 'string', description: 'The message ID' },
        isRead: { type: 'boolean', description: 'Set to false to mark as unread. Default: true' },
      },
      required: ['messageId'],
    },
  },

  // Calendar
  {
    name: 'calendar_list',
    description: 'List calendar events within a date range',
    inputSchema: {
      type: 'object' as const,
      properties: {
        startDate: { type: 'string', description: 'Start date (ISO). Default: today' },
        endDate: { type: 'string', description: 'End date (ISO). Default: +7 days' },
        maxItems: { type: 'number', description: 'Max events. Default: 50' },
      },
    },
  },
  {
    name: 'calendar_get',
    description: 'Get details of a specific calendar event',
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: { type: 'string', description: 'The event ID' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_create',
    description: 'Create a calendar event. IMPORTANT: Requires user confirmation if inviting attendees.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        subject: { type: 'string', description: 'Event title' },
        start: { type: 'string', description: 'Start datetime (ISO)' },
        end: { type: 'string', description: 'End datetime (ISO)' },
        timeZone: { type: 'string', description: 'Timezone. Default: UTC' },
        location: { type: 'string', description: 'Event location' },
        body: { type: 'string', description: 'Event description' },
        attendees: { type: 'array', items: { type: 'string' }, description: 'Attendee emails' },
        isOnlineMeeting: { type: 'boolean', description: 'Create Teams meeting' },
      },
      required: ['subject', 'start', 'end'],
    },
  },
  {
    name: 'calendar_update',
    description: 'Update a calendar event',
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: { type: 'string', description: 'The event ID' },
        subject: { type: 'string', description: 'New title' },
        start: { type: 'string', description: 'New start datetime' },
        end: { type: 'string', description: 'New end datetime' },
        location: { type: 'string', description: 'New location' },
        body: { type: 'string', description: 'New description' },
      },
      required: ['eventId'],
    },
  },
  {
    name: 'calendar_delete',
    description: 'Delete a calendar event. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        eventId: { type: 'string', description: 'The event ID to delete' },
      },
      required: ['eventId'],
    },
  },

  // Tasks
  {
    name: 'tasks_list_lists',
    description: 'List all task lists (To Do lists)',
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
  },
  {
    name: 'tasks_list',
    description: 'List tasks from a task list',
    inputSchema: {
      type: 'object' as const,
      properties: {
        listId: { type: 'string', description: 'Task list ID. Default: default list' },
        includeCompleted: { type: 'boolean', description: 'Include completed tasks' },
        maxItems: { type: 'number', description: 'Max tasks. Default: 50' },
      },
    },
  },
  {
    name: 'tasks_create',
    description: 'Create a new task',
    inputSchema: {
      type: 'object' as const,
      properties: {
        title: { type: 'string', description: 'Task title' },
        listId: { type: 'string', description: 'Task list ID' },
        dueDate: { type: 'string', description: 'Due date (ISO)' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Importance' },
        body: { type: 'string', description: 'Task notes' },
      },
      required: ['title'],
    },
  },
  {
    name: 'tasks_update',
    description: 'Update a task',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'Task ID' },
        listId: { type: 'string', description: 'Task list ID' },
        title: { type: 'string', description: 'New title' },
        dueDate: { type: 'string', description: 'New due date' },
        importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'New importance' },
        body: { type: 'string', description: 'New notes' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'tasks_complete',
    description: 'Mark a task as completed',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'Task ID' },
        listId: { type: 'string', description: 'Task list ID' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'tasks_delete',
    description: 'Delete a task. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'Task ID' },
        listId: { type: 'string', description: 'Task list ID' },
      },
      required: ['taskId'],
    },
  },

  // OneDrive
  {
    name: 'onedrive_list',
    description: 'List files and folders in OneDrive',
    inputSchema: {
      type: 'object' as const,
      properties: {
        path: { type: 'string', description: 'Folder path. Default: root' },
        maxItems: { type: 'number', description: 'Max items. Default: 50' },
      },
    },
  },
  {
    name: 'onedrive_get',
    description: 'Get metadata for a file or folder',
    inputSchema: {
      type: 'object' as const,
      properties: {
        path: { type: 'string', description: 'File/folder path' },
      },
      required: ['path'],
    },
  },
  {
    name: 'onedrive_search',
    description: 'Search for files in OneDrive',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: { type: 'string', description: 'Search query' },
        maxItems: { type: 'number', description: 'Max results. Default: 25' },
      },
      required: ['query'],
    },
  },
  {
    name: 'onedrive_read',
    description: 'Read content of a text file (max 1MB)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        path: { type: 'string', description: 'File path' },
      },
      required: ['path'],
    },
  },
  {
    name: 'onedrive_create_folder',
    description: 'Create a new folder',
    inputSchema: {
      type: 'object' as const,
      properties: {
        name: { type: 'string', description: 'Folder name' },
        parentPath: { type: 'string', description: 'Parent path. Default: root' },
      },
      required: ['name'],
    },
  },
  {
    name: 'onedrive_shared_with_me',
    description: 'List files and folders shared with you by others',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max items. Default: 50' },
      },
    },
  },

  // SharePoint
  {
    name: 'sharepoint_list_sites',
    description: 'List SharePoint sites you follow or search for sites',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max sites. Default: 50' },
        search: { type: 'string', description: 'Search query to find sites' },
      },
    },
  },
  {
    name: 'sharepoint_get_site',
    description: 'Get details of a specific SharePoint site',
    inputSchema: {
      type: 'object' as const,
      properties: {
        siteId: { type: 'string', description: 'Site ID or hostname:path (e.g., "contoso.sharepoint.com:/sites/team")' },
      },
      required: ['siteId'],
    },
  },
  {
    name: 'sharepoint_list_drives',
    description: 'List document libraries in a SharePoint site',
    inputSchema: {
      type: 'object' as const,
      properties: {
        siteId: { type: 'string', description: 'Site ID' },
        maxItems: { type: 'number', description: 'Max drives. Default: 50' },
      },
      required: ['siteId'],
    },
  },
  {
    name: 'sharepoint_list_files',
    description: 'List files in a SharePoint document library (drive)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        driveId: { type: 'string', description: 'Drive ID from sharepoint_list_drives' },
        path: { type: 'string', description: 'Folder path. Default: root' },
        maxItems: { type: 'number', description: 'Max items. Default: 50' },
      },
      required: ['driveId'],
    },
  },
  {
    name: 'sharepoint_get_file',
    description: 'Get metadata for a file in SharePoint',
    inputSchema: {
      type: 'object' as const,
      properties: {
        driveId: { type: 'string', description: 'Drive ID' },
        path: { type: 'string', description: 'File path' },
      },
      required: ['driveId', 'path'],
    },
  },
  {
    name: 'sharepoint_read_file',
    description: 'Read content of a text file from SharePoint (max 1MB)',
    inputSchema: {
      type: 'object' as const,
      properties: {
        driveId: { type: 'string', description: 'Drive ID' },
        path: { type: 'string', description: 'File path' },
      },
      required: ['driveId', 'path'],
    },
  },
  {
    name: 'sharepoint_search_files',
    description: 'Search for files in a SharePoint document library',
    inputSchema: {
      type: 'object' as const,
      properties: {
        driveId: { type: 'string', description: 'Drive ID' },
        query: { type: 'string', description: 'Search query' },
        maxItems: { type: 'number', description: 'Max results. Default: 25' },
      },
      required: ['driveId', 'query'],
    },
  },

  // Contacts
  {
    name: 'contacts_list',
    description: 'List contacts',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max contacts. Default: 50' },
      },
    },
  },
  {
    name: 'contacts_search',
    description: 'Search contacts by name, email, or company',
    inputSchema: {
      type: 'object' as const,
      properties: {
        query: { type: 'string', description: 'Search query' },
        maxItems: { type: 'number', description: 'Max results. Default: 25' },
      },
      required: ['query'],
    },
  },
  {
    name: 'contacts_get',
    description: 'Get details of a specific contact',
    inputSchema: {
      type: 'object' as const,
      properties: {
        contactId: { type: 'string', description: 'Contact ID' },
      },
      required: ['contactId'],
    },
  },
  {
    name: 'contacts_create',
    description: 'Create a new contact',
    inputSchema: {
      type: 'object' as const,
      properties: {
        givenName: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        mobilePhone: { type: 'string', description: 'Mobile phone' },
        businessPhone: { type: 'string', description: 'Business phone' },
        companyName: { type: 'string', description: 'Company name' },
        jobTitle: { type: 'string', description: 'Job title' },
      },
    },
  },
  {
    name: 'contacts_update',
    description: 'Update an existing contact',
    inputSchema: {
      type: 'object' as const,
      properties: {
        contactId: { type: 'string', description: 'Contact ID' },
        givenName: { type: 'string', description: 'First name' },
        surname: { type: 'string', description: 'Last name' },
        email: { type: 'string', description: 'Email address' },
        mobilePhone: { type: 'string', description: 'Mobile phone' },
        businessPhone: { type: 'string', description: 'Business phone' },
        companyName: { type: 'string', description: 'Company name' },
        jobTitle: { type: 'string', description: 'Job title' },
      },
      required: ['contactId'],
    },
  },

  // Teams
  {
    name: 'teams_list',
    description: 'List Teams the user is a member of',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max teams to return. Default: 50' },
      },
    },
  },
  {
    name: 'teams_channels',
    description: 'List channels in a Team',
    inputSchema: {
      type: 'object' as const,
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
      },
      required: ['teamId'],
    },
  },
  {
    name: 'teams_channel_messages',
    description: 'Read recent messages from a Teams channel',
    inputSchema: {
      type: 'object' as const,
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        maxItems: { type: 'number', description: 'Max messages. Default: 25' },
      },
      required: ['teamId', 'channelId'],
    },
  },
  {
    name: 'teams_channel_post',
    description: 'Post a message to a Teams channel. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        teamId: { type: 'string', description: 'Team ID' },
        channelId: { type: 'string', description: 'Channel ID' },
        content: { type: 'string', description: 'Message content (supports HTML)' },
      },
      required: ['teamId', 'channelId', 'content'],
    },
  },

  // Chats
  {
    name: 'chats_list',
    description: 'List 1:1 and group chats',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max chats to return. Default: 25' },
      },
    },
  },
  {
    name: 'chats_messages',
    description: 'Read messages from a chat',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chatId: { type: 'string', description: 'Chat ID' },
        maxItems: { type: 'number', description: 'Max messages. Default: 25' },
      },
      required: ['chatId'],
    },
  },
  {
    name: 'chats_send',
    description: 'Send a message in a chat. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        chatId: { type: 'string', description: 'Chat ID' },
        content: { type: 'string', description: 'Message content (supports HTML)' },
      },
      required: ['chatId', 'content'],
    },
  },
  {
    name: 'chats_create',
    description: 'Create a new 1:1 or group chat. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        members: { type: 'array', items: { type: 'string' }, description: 'Email addresses of chat members' },
        topic: { type: 'string', description: 'Chat topic/title (for group chats)' },
      },
      required: ['members'],
    },
  },

  // Planner
  {
    name: 'planner_list_plans',
    description: 'List all Planner plans the user has access to',
    inputSchema: {
      type: 'object' as const,
      properties: {
        maxItems: { type: 'number', description: 'Max plans to return. Default: 50' },
      },
    },
  },
  {
    name: 'planner_get_plan',
    description: 'Get details of a specific Planner plan',
    inputSchema: {
      type: 'object' as const,
      properties: {
        planId: { type: 'string', description: 'The plan ID' },
      },
      required: ['planId'],
    },
  },
  {
    name: 'planner_list_buckets',
    description: 'List buckets (columns) in a Planner plan',
    inputSchema: {
      type: 'object' as const,
      properties: {
        planId: { type: 'string', description: 'The plan ID' },
      },
      required: ['planId'],
    },
  },
  {
    name: 'planner_create_bucket',
    description: 'Create a new bucket in a Planner plan',
    inputSchema: {
      type: 'object' as const,
      properties: {
        planId: { type: 'string', description: 'The plan ID' },
        name: { type: 'string', description: 'Bucket name' },
      },
      required: ['planId', 'name'],
    },
  },
  {
    name: 'planner_update_bucket',
    description: 'Update a bucket name',
    inputSchema: {
      type: 'object' as const,
      properties: {
        bucketId: { type: 'string', description: 'The bucket ID' },
        name: { type: 'string', description: 'New bucket name' },
      },
      required: ['bucketId', 'name'],
    },
  },
  {
    name: 'planner_delete_bucket',
    description: 'Delete a bucket and its tasks. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        bucketId: { type: 'string', description: 'The bucket ID' },
      },
      required: ['bucketId'],
    },
  },
  {
    name: 'planner_list_tasks',
    description: 'List tasks in a Planner plan',
    inputSchema: {
      type: 'object' as const,
      properties: {
        planId: { type: 'string', description: 'The plan ID' },
        bucketId: { type: 'string', description: 'Filter by bucket ID' },
        maxItems: { type: 'number', description: 'Max tasks to return. Default: 100' },
      },
      required: ['planId'],
    },
  },
  {
    name: 'planner_get_task',
    description: 'Get details of a specific Planner task',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'The task ID' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'planner_create_task',
    description: 'Create a new task in a Planner plan',
    inputSchema: {
      type: 'object' as const,
      properties: {
        planId: { type: 'string', description: 'The plan ID' },
        title: { type: 'string', description: 'Task title' },
        bucketId: { type: 'string', description: 'Bucket ID to place the task in' },
        assignments: { type: 'array', items: { type: 'string' }, description: 'Email addresses of users to assign' },
        dueDateTime: { type: 'string', description: 'Due date (ISO format)' },
        startDateTime: { type: 'string', description: 'Start date (ISO format)' },
        priority: { type: 'string', enum: ['urgent', 'important', 'medium', 'low'], description: 'Task priority' },
        progress: { type: 'string', enum: ['notStarted', 'inProgress', 'completed'], description: 'Task progress' },
      },
      required: ['planId', 'title'],
    },
  },
  {
    name: 'planner_update_task',
    description: 'Update a Planner task',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'The task ID' },
        title: { type: 'string', description: 'New task title' },
        bucketId: { type: 'string', description: 'Move to different bucket' },
        assignments: { type: 'array', items: { type: 'string' }, description: 'Email addresses of users to assign (replaces existing)' },
        dueDateTime: { type: 'string', description: 'New due date (ISO format)' },
        startDateTime: { type: 'string', description: 'New start date (ISO format)' },
        priority: { type: 'string', enum: ['urgent', 'important', 'medium', 'low'], description: 'New priority' },
        progress: { type: 'string', enum: ['notStarted', 'inProgress', 'completed'], description: 'New progress' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'planner_delete_task',
    description: 'Delete a Planner task. IMPORTANT: Requires user confirmation.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'The task ID' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'planner_get_task_details',
    description: 'Get task details including description and checklist',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'The task ID' },
      },
      required: ['taskId'],
    },
  },
  {
    name: 'planner_update_task_details',
    description: 'Update task description and checklist',
    inputSchema: {
      type: 'object' as const,
      properties: {
        taskId: { type: 'string', description: 'The task ID' },
        description: { type: 'string', description: 'Task description' },
        checklist: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Checklist item title' },
              isChecked: { type: 'boolean', description: 'Whether item is checked' },
            },
            required: ['title'],
          },
          description: 'Checklist items (replaces existing checklist)',
        },
      },
      required: ['taskId'],
    },
  },

  // Auth
  {
    name: 'auth_status',
    description: 'Check authentication status and current user',
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
  },
  {
    name: 'debug_info',
    description: 'Get debug information about the MCP server configuration',
    inputSchema: {
      type: 'object' as const,
      properties: {},
    },
  },
];

// Handler for listing tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: TOOLS };
});

// Handler for calling tools
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  const result = await executeCommand(name, args as Record<string, unknown>);

  if (result.success) {
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result.data, null, 2),
        },
      ],
    };
  } else {
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${result.error}`,
        },
      ],
      isError: true,
    };
  }
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Personal M365 MCP server started');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
