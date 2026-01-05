import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

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

export const sendMailSchema = z.object({
  to: z.array(z.string()).describe('List of recipient email addresses'),
  subject: z.string().describe('Email subject'),
  body: z.string().describe('Email body (plain text or HTML)'),
  isHtml: z.boolean().optional().describe('Whether body is HTML. Default: false'),
  cc: z.array(z.string()).optional().describe('CC recipients'),
  bcc: z.array(z.string()).optional().describe('BCC recipients'),
});

export const deleteMailSchema = z.object({
  messageId: z.string().describe('The ID of the message to delete'),
});

export const markAsReadSchema = z.object({
  messageId: z.string().describe('The ID of the message to mark as read'),
  isRead: z.boolean().optional().describe('Set to false to mark as unread. Default: true'),
});

// Tool implementations
export async function listMails(params: z.infer<typeof listMailsSchema>) {
  const { folder = 'inbox', maxItems = 25, unreadOnly = false } = params;

  let path = `/me/mailFolders/${folder}/messages?$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments&$orderby=receivedDateTime desc&$top=${maxItems}`;

  if (unreadOnly) {
    path += '&$filter=isRead eq false';
  }

  const messages = await graphList<Message>(path, { maxItems });

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

export async function sendMail(params: z.infer<typeof sendMailSchema>) {
  const { to, subject, body, isHtml = false, cc, bcc } = params;

  const message = {
    message: {
      subject,
      body: {
        contentType: isHtml ? 'HTML' : 'Text',
        content: body,
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
    },
    saveToSentItems: true,
  };

  await graphRequest('/me/sendMail', {
    method: 'POST',
    body: message,
  });

  return { success: true, message: `Email sent to ${to.join(', ')}` };
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
