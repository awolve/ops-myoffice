import { z } from 'zod';
import { graphRequest, graphList, graphDownload } from '../utils/graph-client.js';
import { getSignature } from '../utils/signature.js';
import { promises as fs } from 'fs';
import { resolve } from 'path';

// Types
interface EmailAddress {
  emailAddress: {
    name?: string;
    address: string;
  };
}

interface Message {
  id: string;
  subject: string;
  from?: EmailAddress;
  toRecipients?: EmailAddress[];
  receivedDateTime: string;
  isRead: boolean;
  bodyPreview?: string;
  body?: {
    contentType: string;
    content: string;
  };
  hasAttachments?: boolean;
}

interface Attachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
  '@odata.type': string;
}

// Schemas
export const listMailsSchema = z.object({
  folder: z.string().optional().describe('Folder to list (inbox, sentitems, drafts, etc.). Default: inbox'),
  maxItems: z.number().optional().describe('Maximum number of emails to return. Default: 25'),
  unreadOnly: z.boolean().optional().describe('Only return unread emails'),
});

export const readMailSchema = z.object({
  messageId: z.string().describe('The ID of the message to read'),
});

export const searchMailSchema = z.object({
  query: z.string().describe('Search query (searches subject, body, and participants)'),
  maxItems: z.number().optional().describe('Maximum number of results. Default: 25'),
});

export const createDraftSchema = z.object({
  to: z.array(z.string()).describe('List of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body (plain text or HTML)'),
  isHtml: z.boolean().optional().describe('Whether body is HTML. Default: true'),
  cc: z.array(z.string()).optional().describe('CC recipients'),
  bcc: z.array(z.string()).optional().describe('BCC recipients'),
  useSignature: z.boolean().optional().describe('Append email signature if configured. Default: true'),
  attachments: z.array(z.string()).optional().describe('List of file paths to attach'),
});

export const sendMailSchema = z.object({
  to: z.array(z.string()).describe('List of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body (plain text or HTML)'),
  isHtml: z.boolean().optional().describe('Whether body is HTML. Default: true'),
  cc: z.array(z.string()).optional().describe('CC recipients'),
  bcc: z.array(z.string()).optional().describe('BCC recipients'),
  useSignature: z.boolean().optional().describe('Append email signature if configured. Default: true'),
  attachments: z.array(z.string()).optional().describe('List of file paths to attach'),
});

export const replyMailSchema = z.object({
  messageId: z.string().describe('The ID of the message to reply to'),
  body: z.string().describe('Reply body (HTML by default)'),
  isHtml: z.boolean().optional().describe('Whether body is HTML. Default: true'),
  replyAll: z.boolean().optional().describe('Reply to all recipients. Default: false'),
  useSignature: z.boolean().optional().describe('Append email signature if configured. Default: false'),
});

export const deleteMailSchema = z.object({
  messageId: z.string().describe('The ID of the message to delete'),
});

export const markAsReadSchema = z.object({
  messageId: z.string().describe('The ID of the message to mark as read'),
  isRead: z.boolean().optional().describe('Set to false to mark as unread. Default: true'),
});

export const listAttachmentsSchema = z.object({
  messageId: z.string().describe('The ID of the message to list attachments from'),
});

export const downloadAttachmentSchema = z.object({
  messageId: z.string().describe('The ID of the message containing the attachment'),
  attachmentId: z.string().describe('The ID of the attachment to download'),
  outputPath: z.string().describe('Path where to save the attachment'),
});

export const moveMailSchema = z.object({
  messageId: z.string().describe('The ID of the message to move'),
  folderName: z.string().describe('Name of the destination folder (e.g., "Under Processing", "Captured", "Skipped")'),
});

// Helper: Get folder ID by name (searches if not a well-known folder)
async function getFolderIdByName(folderName: string): Promise<string | null> {
  interface MailFolder {
    id: string;
    displayName: string;
    parentFolderId?: string;
  }

  console.error(`[DEBUG] Resolving folder: "${folderName}"`);

  // Well-known folder names that Graph API accepts directly
  const wellKnownFolders = ['inbox', 'sentitems', 'drafts', 'deleteditems', 'junkemail', 'archive'];

  // If it's a well-known folder, return as-is
  if (wellKnownFolders.includes(folderName.toLowerCase())) {
    console.error(`[DEBUG] Using well-known folder ID: "${folderName.toLowerCase()}"`);
    return folderName.toLowerCase();
  }

  // List all folders to help debug
  console.error('[DEBUG] Fetching all mail folders...');
  const allFolders = await graphList<MailFolder>(
    `/me/mailFolders?$select=id,displayName,parentFolderId&$top=100`
  );

  console.error(`[DEBUG] Found ${allFolders.length} folders:`);
  allFolders.forEach(f => {
    console.error(`[DEBUG]   - "${f.displayName}" (ID: ${f.id})`);
  });

  // Search for folder by displayName (case-insensitive)
  const matchingFolder = allFolders.find(
    f => f.displayName.toLowerCase() === folderName.toLowerCase()
  );

  if (matchingFolder) {
    console.error(`[DEBUG] Resolved folder "${folderName}" to ID: ${matchingFolder.id}`);
    return matchingFolder.id;
  }

  console.error(`[DEBUG] Folder "${folderName}" not found`);
  return null;
}

export const listFoldersSchema = z.object({});

// Tool implementations
export async function listFolders() {
  interface MailFolder {
    id: string;
    displayName: string;
    parentFolderId: string;
    totalItemCount: number;
    unreadItemCount: number;
    childFolderCount: number;
  }

  const folders = await graphList<MailFolder>(
    `/me/mailFolders?$select=id,displayName,parentFolderId,totalItemCount,unreadItemCount,childFolderCount&$top=100`
  );

  // Fetch child folders for any folder with children
  const allFolders: Array<{
    id: string;
    name: string;
    parentId: string;
    totalItems: number;
    unreadItems: number;
  }> = [];

  for (const f of folders) {
    allFolders.push({
      id: f.id,
      name: f.displayName,
      parentId: f.parentFolderId,
      totalItems: f.totalItemCount,
      unreadItems: f.unreadItemCount,
    });

    if (f.childFolderCount > 0) {
      const children = await graphList<MailFolder>(
        `/me/mailFolders/${f.id}/childFolders?$select=id,displayName,parentFolderId,totalItemCount,unreadItemCount,childFolderCount&$top=100`
      );
      for (const child of children) {
        allFolders.push({
          id: child.id,
          name: `${f.displayName}/${child.displayName}`,
          parentId: child.parentFolderId,
          totalItems: child.totalItemCount,
          unreadItems: child.unreadItemCount,
        });
      }
    }
  }

  return allFolders;
}

export async function listMails(params: z.infer<typeof listMailsSchema>) {
  const { folder = 'inbox', maxItems = 25, unreadOnly = false } = params;

  console.error(`[DEBUG] listMails called with folder: "${folder}"`);

  // Resolve folder name to folder ID
  const folderId = await getFolderIdByName(folder);

  if (!folderId) {
    throw new Error(`Folder not found: ${folder}`);
  }

  // URL encode the folder ID in case it contains special characters
  const encodedFolderId = encodeURIComponent(folderId);
  console.error(`[DEBUG] Folder ID (raw): ${folderId}`);
  console.error(`[DEBUG] Folder ID (encoded): ${encodedFolderId}`);

  let path = `/me/mailFolders/${encodedFolderId}/messages?$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments&$orderby=receivedDateTime desc&$top=${maxItems}`;

  if (unreadOnly) {
    path += '&$filter=isRead eq false';
  }

  console.error(`[DEBUG] Calling Graph API with path: ${path}`);

  let messages: Message[];
  try {
    messages = await graphList<Message>(path, { maxItems });
    console.error(`[DEBUG] Retrieved ${messages.length} messages`);
  } catch (error) {
    console.error(`[DEBUG] Error calling Graph API:`, error);
    throw error;
  }

  return messages.map((m) => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    fromName: m.from?.emailAddress?.name,
    received: m.receivedDateTime,
    isRead: m.isRead,
    preview: m.bodyPreview?.substring(0, 200),
    hasAttachments: m.hasAttachments,
  }));
}

export async function readMail(params: z.infer<typeof readMailSchema>) {
  const { messageId } = params;

  const message = await graphRequest<Message>(
    `/me/messages/${messageId}?$select=id,subject,from,toRecipients,receivedDateTime,body,hasAttachments`
  );

  return {
    id: message.id,
    subject: message.subject,
    from: message.from?.emailAddress?.address,
    fromName: message.from?.emailAddress?.name,
    to: message.toRecipients?.map((r) => r.emailAddress.address),
    received: message.receivedDateTime,
    body: message.body?.content,
    bodyType: message.body?.contentType,
    hasAttachments: message.hasAttachments,
  };
}

export async function searchMail(params: z.infer<typeof searchMailSchema>) {
  const { query, maxItems = 25 } = params;

  const path = `/me/messages?$search="${encodeURIComponent(query)}"&$select=id,subject,from,receivedDateTime,isRead,bodyPreview&$top=${maxItems}`;

  const messages = await graphList<Message>(path, { maxItems });

  return messages.map((m) => ({
    id: m.id,
    subject: m.subject,
    from: m.from?.emailAddress?.address,
    received: m.receivedDateTime,
    isRead: m.isRead,
    preview: m.bodyPreview?.substring(0, 200),
  }));
}

// Convert plain text to HTML (escape special chars, convert newlines to <br>)
function textToHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}

// Check if string looks like HTML (contains tags)
function looksLikeHtml(text: string): boolean {
  return /<[a-z][\s\S]*>/i.test(text);
}

// Helper: Prepare attachments from file paths
async function prepareAttachments(filePaths: string[]): Promise<Array<{
  '@odata.type': string;
  name: string;
  contentType: string;
  contentBytes: string;
}>> {
  const attachments = [];

  for (const filePath of filePaths) {
    const absolutePath = resolve(filePath);
    const fileName = filePath.split('/').pop() || 'attachment';

    // Read file content
    const fileBuffer = await fs.readFile(absolutePath);
    const base64Content = fileBuffer.toString('base64');

    // Detect MIME type based on extension
    const ext = fileName.split('.').pop()?.toLowerCase();
    const mimeTypes: Record<string, string> = {
      'pdf': 'application/pdf',
      'doc': 'application/msword',
      'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'xls': 'application/vnd.ms-excel',
      'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'ppt': 'application/vnd.ms-powerpoint',
      'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'txt': 'text/plain',
      'csv': 'text/csv',
      'json': 'application/json',
      'xml': 'application/xml',
      'zip': 'application/zip',
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'png': 'image/png',
      'gif': 'image/gif',
      'svg': 'image/svg+xml',
    };

    const contentType = mimeTypes[ext || ''] || 'application/octet-stream';

    attachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: fileName,
      contentType,
      contentBytes: base64Content,
    });
  }

  return attachments;
}

export async function createDraft(params: z.infer<typeof createDraftSchema>) {
  const { to, subject, body, isHtml = true, cc, bcc, useSignature = true, attachments: attachmentPaths } = params;

  // Prepare body - convert plain text to HTML if needed
  let finalBody = body;
  if (isHtml && !looksLikeHtml(body)) {
    finalBody = textToHtml(body);
  }

  // Append signature if enabled and configured
  if (useSignature) {
    const signature = getSignature();
    if (signature) {
      finalBody = isHtml
        ? `${finalBody}<br><br>${signature}`
        : `${finalBody}\n\n--\n${signature}`;
    }
  }

  // Prepare attachments if provided
  const attachments = attachmentPaths && attachmentPaths.length > 0
    ? await prepareAttachments(attachmentPaths)
    : undefined;

  const message: Record<string, unknown> = {
    subject,
    body: {
      contentType: isHtml ? 'HTML' : 'Text',
      content: finalBody,
    },
    toRecipients: to.map((addr) => ({
      emailAddress: { address: addr },
    })),
    ccRecipients: cc?.map((addr) => ({
      emailAddress: { address: addr },
    })),
    bccRecipients: bcc?.map((addr) => ({
      emailAddress: { address: addr },
    })),
  };

  if (attachments) {
    message.attachments = attachments;
  }

  const result = await graphRequest<Message>('/me/messages', {
    method: 'POST',
    body: message,
  });

  return {
    success: true,
    message: `Draft created for ${to.join(', ')}`,
    draftId: result.id,
  };
}

export async function sendMail(params: z.infer<typeof sendMailSchema>) {
  const { to, subject, body, isHtml = true, cc, bcc, useSignature = true, attachments: attachmentPaths } = params;

  // Prepare body - convert plain text to HTML if needed
  let finalBody = body;
  if (isHtml && !looksLikeHtml(body)) {
    finalBody = textToHtml(body);
  }

  // Append signature if enabled and configured
  if (useSignature) {
    const signature = getSignature();
    if (signature) {
      finalBody = isHtml
        ? `${finalBody}<br><br>${signature}`
        : `${finalBody}\n\n--\n${signature}`;
    }
  }

  // Prepare attachments if provided
  const attachments = attachmentPaths && attachmentPaths.length > 0
    ? await prepareAttachments(attachmentPaths)
    : undefined;

  const messageContent: Record<string, unknown> = {
    subject,
    body: {
      contentType: isHtml ? 'HTML' : 'Text',
      content: finalBody,
    },
    toRecipients: to.map((addr) => ({
      emailAddress: { address: addr },
    })),
    ccRecipients: cc?.map((addr) => ({
      emailAddress: { address: addr },
    })),
    bccRecipients: bcc?.map((addr) => ({
      emailAddress: { address: addr },
    })),
  };

  if (attachments) {
    messageContent.attachments = attachments;
  }

  const message = {
    message: messageContent,
    saveToSentItems: true,
  };

  await graphRequest('/me/sendMail', {
    method: 'POST',
    body: message,
  });

  return { success: true, message: `Email sent to ${to.join(', ')}` };
}

export async function replyMail(params: z.infer<typeof replyMailSchema>) {
  const { messageId, body, isHtml = true, replyAll = false, useSignature = false } = params;

  // Prepare body - convert plain text to HTML if needed
  let finalBody = body;
  if (isHtml && !looksLikeHtml(body)) {
    finalBody = textToHtml(body);
  }

  // Append signature if enabled and configured
  if (useSignature) {
    const signature = getSignature();
    if (signature) {
      finalBody = isHtml
        ? `${finalBody}<br><br>${signature}`
        : `${finalBody}\n\n--\n${signature}`;
    }
  }

  const endpoint = replyAll
    ? `/me/messages/${messageId}/replyAll`
    : `/me/messages/${messageId}/reply`;

  await graphRequest(endpoint, {
    method: 'POST',
    body: {
      message: {
        body: {
          contentType: isHtml ? 'HTML' : 'Text',
          content: finalBody,
        },
      },
    },
  });

  return { success: true, message: replyAll ? 'Reply sent to all' : 'Reply sent' };
}

export async function deleteMail(params: z.infer<typeof deleteMailSchema>) {
  const { messageId } = params;

  await graphRequest(`/me/messages/${messageId}`, {
    method: 'DELETE',
  });

  return { success: true, message: 'Email deleted' };
}

export async function markAsRead(params: z.infer<typeof markAsReadSchema>) {
  const { messageId, isRead = true } = params;

  await graphRequest(`/me/messages/${messageId}`, {
    method: 'PATCH',
    body: { isRead },
  });

  return { success: true, message: isRead ? 'Email marked as read' : 'Email marked as unread' };
}

export async function listAttachments(params: z.infer<typeof listAttachmentsSchema>) {
  const { messageId } = params;

  const attachments = await graphList<Attachment>(
    `/me/messages/${messageId}/attachments?$select=id,name,contentType,size,isInline`
  );

  return attachments.map((a) => ({
    id: a.id,
    name: a.name,
    contentType: a.contentType,
    size: a.size,
    isInline: a.isInline,
  }));
}

export async function downloadAttachment(params: z.infer<typeof downloadAttachmentSchema>) {
  const { messageId, attachmentId, outputPath } = params;

  // Get attachment with content (contentBytes is included by default for fileAttachment)
  const attachment = await graphRequest<Attachment & { contentBytes?: string }>(
    `/me/messages/${messageId}/attachments/${attachmentId}`
  );

  // For file attachments, download the content
  if (attachment['@odata.type'] === '#microsoft.graph.fileAttachment' && attachment.contentBytes) {
    const buffer = Buffer.from(attachment.contentBytes, 'base64');
    const absolutePath = resolve(outputPath);
    await fs.writeFile(absolutePath, buffer);

    return {
      success: true,
      message: `Attachment saved to ${absolutePath}`,
      name: attachment.name,
      size: buffer.length,
      path: absolutePath,
    };
  } else {
    throw new Error('Unsupported attachment type. Only file attachments can be downloaded.');
  }
}

// Helper: Get or create a mail folder by name
async function getOrCreateFolder(folderName: string): Promise<string> {
  interface MailFolder {
    id: string;
    displayName: string;
  }

  // First, try to find existing folder (but don't use well-known folder logic)
  const folders = await graphList<MailFolder>(
    `/me/mailFolders?$filter=displayName eq '${folderName.replace(/'/g, "''")}'&$select=id,displayName`
  );

  if (folders.length > 0) {
    return folders[0].id;
  }

  // Folder doesn't exist, create it under root
  const newFolder = await graphRequest<MailFolder>('/me/mailFolders', {
    method: 'POST',
    body: {
      displayName: folderName,
    },
  });

  return newFolder.id;
}

export async function moveMail(params: z.infer<typeof moveMailSchema>) {
  const { messageId, folderName } = params;

  // Get or create the destination folder
  const folderId = await getOrCreateFolder(folderName);

  // Move the message
  await graphRequest(`/me/messages/${messageId}/move`, {
    method: 'POST',
    body: {
      destinationId: folderId,
    },
  });

  return {
    success: true,
    messageId,
    folder: folderName,
  };
}
