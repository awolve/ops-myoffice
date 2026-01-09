#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { getCurrentUser, isAuthenticated } from './auth/index.js';
import { getVersion } from './utils/version.js';
import { graphRequest } from './utils/graph-client.js';
import * as mail from './tools/mail.js';
import * as calendar from './tools/calendar.js';
import * as tasks from './tools/tasks.js';
import * as onedrive from './tools/onedrive.js';
import * as sharepoint from './tools/sharepoint.js';
import * as contacts from './tools/contacts.js';
import * as teams from './tools/teams.js';
import * as chats from './tools/chats.js';
import * as planner from './tools/planner.js';

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

  try {
    let result: unknown;

    switch (name) {
      // Mail
      case 'mail_list':
        result = await mail.listMails(mail.listMailsSchema.parse(args));
        break;
      case 'mail_read':
        result = await mail.readMail(mail.readMailSchema.parse(args));
        break;
      case 'mail_search':
        result = await mail.searchMail(mail.searchMailSchema.parse(args));
        break;
      case 'mail_send':
        result = await mail.sendMail(mail.sendMailSchema.parse(args));
        break;
      case 'mail_reply':
        result = await mail.replyMail(mail.replyMailSchema.parse(args));
        break;
      case 'mail_delete':
        result = await mail.deleteMail(mail.deleteMailSchema.parse(args));
        break;
      case 'mail_mark_read':
        result = await mail.markAsRead(mail.markAsReadSchema.parse(args));
        break;

      // Calendar
      case 'calendar_list':
        result = await calendar.listEvents(calendar.listEventsSchema.parse(args));
        break;
      case 'calendar_get':
        result = await calendar.getEvent(calendar.getEventSchema.parse(args));
        break;
      case 'calendar_create':
        result = await calendar.createEvent(calendar.createEventSchema.parse(args));
        break;
      case 'calendar_update':
        result = await calendar.updateEvent(calendar.updateEventSchema.parse(args));
        break;
      case 'calendar_delete':
        result = await calendar.deleteEvent(calendar.deleteEventSchema.parse(args));
        break;

      // Tasks
      case 'tasks_list_lists':
        result = await tasks.listTaskLists();
        break;
      case 'tasks_list':
        result = await tasks.listTasks(tasks.listTasksSchema.parse(args));
        break;
      case 'tasks_create':
        result = await tasks.createTask(tasks.createTaskSchema.parse(args));
        break;
      case 'tasks_update':
        result = await tasks.updateTask(tasks.updateTaskSchema.parse(args));
        break;
      case 'tasks_complete':
        result = await tasks.completeTask(tasks.completeTaskSchema.parse(args));
        break;
      case 'tasks_delete':
        result = await tasks.deleteTask(tasks.deleteTaskSchema.parse(args));
        break;

      // OneDrive
      case 'onedrive_list':
        result = await onedrive.listFiles(onedrive.listFilesSchema.parse(args));
        break;
      case 'onedrive_get':
        result = await onedrive.getFile(onedrive.getFileSchema.parse(args));
        break;
      case 'onedrive_search':
        result = await onedrive.searchFiles(onedrive.searchFilesSchema.parse(args));
        break;
      case 'onedrive_read':
        result = await onedrive.readFileContent(onedrive.readFileContentSchema.parse(args));
        break;
      case 'onedrive_create_folder':
        result = await onedrive.createFolder(onedrive.createFolderSchema.parse(args));
        break;
      case 'onedrive_shared_with_me':
        result = await onedrive.listSharedWithMe(onedrive.listSharedWithMeSchema.parse(args));
        break;

      // SharePoint
      case 'sharepoint_list_sites':
        result = await sharepoint.listSites(sharepoint.listSitesSchema.parse(args));
        break;
      case 'sharepoint_get_site':
        result = await sharepoint.getSite(sharepoint.getSiteSchema.parse(args));
        break;
      case 'sharepoint_list_drives':
        result = await sharepoint.listDrives(sharepoint.listDrivesSchema.parse(args));
        break;
      case 'sharepoint_list_files':
        result = await sharepoint.listDriveFiles(sharepoint.listDriveFilesSchema.parse(args));
        break;
      case 'sharepoint_get_file':
        result = await sharepoint.getDriveFile(sharepoint.getDriveFileSchema.parse(args));
        break;
      case 'sharepoint_read_file':
        result = await sharepoint.readDriveFile(sharepoint.readDriveFileSchema.parse(args));
        break;
      case 'sharepoint_search_files':
        result = await sharepoint.searchDriveFiles(sharepoint.searchDriveFilesSchema.parse(args));
        break;

      // Contacts
      case 'contacts_list':
        result = await contacts.listContacts(contacts.listContactsSchema.parse(args));
        break;
      case 'contacts_search':
        result = await contacts.searchContacts(contacts.searchContactsSchema.parse(args));
        break;
      case 'contacts_get':
        result = await contacts.getContact(contacts.getContactSchema.parse(args));
        break;
      case 'contacts_create':
        result = await contacts.createContact(contacts.createContactSchema.parse(args));
        break;
      case 'contacts_update':
        result = await contacts.updateContact(contacts.updateContactSchema.parse(args));
        break;

      // Teams
      case 'teams_list':
        result = await teams.listTeams(teams.listTeamsSchema.parse(args));
        break;
      case 'teams_channels':
        result = await teams.listChannels(teams.listChannelsSchema.parse(args));
        break;
      case 'teams_channel_messages':
        result = await teams.listChannelMessages(teams.listChannelMessagesSchema.parse(args));
        break;
      case 'teams_channel_post':
        result = await teams.postChannelMessage(teams.postChannelMessageSchema.parse(args));
        break;

      // Chats
      case 'chats_list':
        result = await chats.listChats(chats.listChatsSchema.parse(args));
        break;
      case 'chats_messages':
        result = await chats.listChatMessages(chats.listChatMessagesSchema.parse(args));
        break;
      case 'chats_send':
        result = await chats.sendChatMessage(chats.sendChatMessageSchema.parse(args));
        break;
      case 'chats_create':
        result = await chats.createChat(chats.createChatSchema.parse(args));
        break;

      // Planner
      case 'planner_list_plans':
        result = await planner.listPlans(planner.listPlansSchema.parse(args));
        break;
      case 'planner_get_plan':
        result = await planner.getPlan(planner.getPlanSchema.parse(args));
        break;
      case 'planner_list_buckets':
        result = await planner.listBuckets(planner.listBucketsSchema.parse(args));
        break;
      case 'planner_create_bucket':
        result = await planner.createBucket(planner.createBucketSchema.parse(args));
        break;
      case 'planner_update_bucket':
        result = await planner.updateBucket(planner.updateBucketSchema.parse(args));
        break;
      case 'planner_delete_bucket':
        result = await planner.deleteBucket(planner.deleteBucketSchema.parse(args));
        break;
      case 'planner_list_tasks':
        result = await planner.listPlannerTasks(planner.listPlannerTasksSchema.parse(args));
        break;
      case 'planner_get_task':
        result = await planner.getPlannerTask(planner.getPlannerTaskSchema.parse(args));
        break;
      case 'planner_create_task':
        result = await planner.createPlannerTask(planner.createPlannerTaskSchema.parse(args));
        break;
      case 'planner_update_task':
        result = await planner.updatePlannerTask(planner.updatePlannerTaskSchema.parse(args));
        break;
      case 'planner_delete_task':
        result = await planner.deletePlannerTask(planner.deletePlannerTaskSchema.parse(args));
        break;
      case 'planner_get_task_details':
        result = await planner.getPlannerTaskDetails(planner.getPlannerTaskDetailsSchema.parse(args));
        break;
      case 'planner_update_task_details':
        result = await planner.updatePlannerTaskDetails(planner.updatePlannerTaskDetailsSchema.parse(args));
        break;

      // Auth
      case 'auth_status':
        result = {
          authenticated: await isAuthenticated(),
          user: await getCurrentUser(),
        };
        break;

      case 'debug_info':
        const uptimeMs = Date.now() - SERVER_START_TIME.getTime();
        const uptimeSec = Math.floor(uptimeMs / 1000);
        const hours = Math.floor(uptimeSec / 3600);
        const minutes = Math.floor((uptimeSec % 3600) / 60);
        const seconds = uptimeSec % 60;
        const uptimeStr = `${hours}h ${minutes}m ${seconds}s`;

        // Test Graph API connectivity with 10 second timeout
        let graphTest: { status: string; responseTimeMs?: number; error?: string; user?: string };
        const testStartTime = Date.now();
        try {
          const timeoutPromise = new Promise<never>((_, reject) => {
            setTimeout(() => reject(new Error('Timeout after 10 seconds')), 10000);
          });

          const graphTestResult = await Promise.race([
            graphRequest<{ displayName?: string; mail?: string }>('/me?$select=displayName,mail'),
            timeoutPromise,
          ]);

          const responseTimeMs = Date.now() - testStartTime;
          graphTest = {
            status: 'OK',
            responseTimeMs,
            user: graphTestResult.displayName || graphTestResult.mail || 'unknown',
          };
        } catch (err) {
          const responseTimeMs = Date.now() - testStartTime;
          graphTest = {
            status: 'FAILED',
            responseTimeMs,
            error: err instanceof Error ? err.message : String(err),
          };
        }

        result = {
          server: {
            name: 'myoffice-mcp',
            version: VERSION,
            nodeVersion: process.version,
            platform: process.platform,
            pid: process.pid,
            startedAt: SERVER_START_TIME.toISOString(),
            uptime: uptimeStr,
          },
          environment: {
            M365_CLIENT_ID: (() => {
              const val = process.env.M365_CLIENT_ID;
              if (!val) return 'NOT SET';
              if (val.startsWith('${') || val.includes('$M365')) return `UNRESOLVED PLACEHOLDER: ${val}`;
              return `SET (${val.substring(0, 8)}...)`;
            })(),
            M365_TENANT_ID: process.env.M365_TENANT_ID || 'NOT SET (using "common")',
          },
          auth: {
            authenticated: await isAuthenticated(),
            user: await getCurrentUser(),
            tokenCachePath: '~/.config/myoffice-mcp/msal-cache.json',
          },
          graphApiTest: graphTest,
          tools: TOOLS.length,
        };
        break;

      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${message}`,
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
