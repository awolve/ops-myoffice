import { getCurrentUser, isAuthenticated } from '../auth/index.js';
import { getVersion } from '../utils/version.js';
import { graphRequest } from '../utils/graph-client.js';
import * as mail from '../tools/mail.js';
import * as calendar from '../tools/calendar.js';
import * as tasks from '../tools/tasks.js';
import * as onedrive from '../tools/onedrive.js';
import * as sharepoint from '../tools/sharepoint.js';
import * as contacts from '../tools/contacts.js';
import * as teams from '../tools/teams.js';
import * as chats from '../tools/chats.js';
import * as planner from '../tools/planner.js';

export interface ToolResult {
  success: boolean;
  data?: unknown;
  error?: string;
  code?: string;
}

const SERVER_START_TIME = new Date();

export async function executeCommand(
  toolName: string,
  args: Record<string, unknown>
): Promise<ToolResult> {
  try {
    let result: unknown;

    switch (toolName) {
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
      case 'onedrive_upload':
        result = await onedrive.uploadFile(onedrive.uploadFileSchema.parse(args));
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
      case 'planner_add_reference':
        result = await planner.addPlannerTaskReference(planner.addPlannerTaskReferenceSchema.parse(args));
        break;
      case 'planner_remove_reference':
        result = await planner.removePlannerTaskReference(planner.removePlannerTaskReferenceSchema.parse(args));
        break;
      case 'planner_upload_attach':
        result = await planner.uploadAndAttach(planner.uploadAndAttachSchema.parse(args));
        break;

      // Auth
      case 'auth_status':
        result = {
          authenticated: await isAuthenticated(),
          user: await getCurrentUser(),
        };
        break;

      case 'debug_info':
        const VERSION = getVersion();
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
        };
        break;

      default:
        return {
          success: false,
          error: `Unknown tool: ${toolName}`,
          code: 'UNKNOWN_TOOL',
        };
    }

    return {
      success: true,
      data: result,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      success: false,
      error: message,
      code: 'EXECUTION_ERROR',
    };
  }
}
