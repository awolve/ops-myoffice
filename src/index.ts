#!/usr/bin/env node

// Debug: Log environment on startup
console.error('=== MCP Server Starting ===');
console.error('M365_CLIENT_ID:', process.env.M365_CLIENT_ID ? `SET (${process.env.M365_CLIENT_ID.substring(0, 8)}...)` : 'NOT SET');
console.error('M365_TENANT_ID:', process.env.M365_TENANT_ID || 'NOT SET (will use "common")');
console.error('===========================');

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';

import { getCurrentUser, isAuthenticated } from './auth/index.js';
import * as mail from './tools/mail.js';
import * as calendar from './tools/calendar.js';
import * as tasks from './tools/tasks.js';
import * as onedrive from './tools/onedrive.js';
import * as sharepoint from './tools/sharepoint.js';
import * as contacts from './tools/contacts.js';

const server = new Server(
  {
    name: 'ops-personal-m365-mcp',
    version: '0.1.0',
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
    description: 'Send an email. IMPORTANT: Requires user confirmation before sending.',
    inputSchema: {
      type: 'object' as const,
      properties: {
        to: { type: 'array', items: { type: 'string' }, description: 'Recipient emails' },
        subject: { type: 'string', description: 'Email subject' },
        body: { type: 'string', description: 'Email body' },
        isHtml: { type: 'boolean', description: 'Is body HTML? Default: false' },
        cc: { type: 'array', items: { type: 'string' }, description: 'CC recipients' },
        bcc: { type: 'array', items: { type: 'string' }, description: 'BCC recipients' },
      },
      required: ['to', 'subject', 'body'],
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

      // Auth
      case 'auth_status':
        result = {
          authenticated: await isAuthenticated(),
          user: getCurrentUser(),
        };
        break;

      case 'debug_info':
        result = {
          server: {
            name: 'ops-personal-m365-mcp',
            version: '0.1.0',
            nodeVersion: process.version,
            platform: process.platform,
            pid: process.pid,
          },
          environment: {
            M365_CLIENT_ID: process.env.M365_CLIENT_ID ? `SET (${process.env.M365_CLIENT_ID.substring(0, 8)}...)` : 'NOT SET',
            M365_TENANT_ID: process.env.M365_TENANT_ID || 'NOT SET (using "common")',
          },
          auth: {
            authenticated: await isAuthenticated(),
            user: getCurrentUser(),
            tokenCachePath: '~/.config/ops-personal-m365-mcp/token.json',
          },
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
