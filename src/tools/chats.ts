import { z } from 'zod';
import { graphRequest, graphList, graphDownload } from '../utils/graph-client.js';

// Types
interface ChatMember {
  id: string;
  displayName?: string;
  email?: string;
}

interface Chat {
  id: string;
  topic?: string;
  chatType: string;
  createdDateTime: string;
  lastUpdatedDateTime?: string;
  members?: ChatMember[];
}

interface ChatMessageFrom {
  user?: {
    id: string;
    displayName?: string;
  };
  application?: {
    id: string;
    displayName?: string;
  };
}

interface ChatMessageAttachment {
  id: string;
  contentType?: string;
  contentUrl?: string;
  content?: string | null; // for messageReference: a JSON string with the quote
  name?: string;
}

/**
 * When a message is a quote-reply, Teams attaches a `messageReference`
 * attachment whose `content` is a JSON string holding the quoted message's id,
 * preview text, and sender. Pull that out so callers can read the quote.
 */
function extractQuotedMessage(attachments?: ChatMessageAttachment[]) {
  const ref = attachments?.find((a) => a.contentType === 'messageReference' && a.content);
  if (!ref?.content) return null;
  try {
    const c = JSON.parse(ref.content);
    return {
      messageId: String(c.messageId ?? ref.id ?? ''),
      preview: String(c.messagePreview ?? ''),
      sender:
        c.messageSender?.user?.displayName ??
        c.messageSender?.application?.displayName ??
        null,
    };
  } catch {
    return null;
  }
}

interface ChatMessage {
  id: string;
  createdDateTime: string;
  from?: ChatMessageFrom;
  body: {
    contentType: string;
    content: string;
  };
  attachments?: ChatMessageAttachment[];
}

interface HostedContent {
  id: string;
  contentType?: string;
}

// Schemas
export const listChatsSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of chats to return. Default: 25'),
});

export const listChatMessagesSchema = z.object({
  chatId: z.string().describe('The ID of the chat'),
  maxItems: z.number().optional().describe('Maximum number of messages to return. Default: 25'),
});

export const sendChatMessageSchema = z.object({
  chatId: z.string().describe('The ID of the chat'),
  content: z.string().describe('Message content (supports HTML)'),
  images: z
    .array(z.string())
    .optional()
    .describe('Paths to image files to embed inline in the message'),
});

export const downloadChatImagesSchema = z.object({
  chatId: z.string().describe('The ID of the chat'),
  messageId: z.string().optional().describe('Download images from this message only; omit to scan recent messages'),
  maxItems: z.number().optional().describe('Number of recent messages to scan when no messageId is given. Default: 25'),
  outputDir: z.string().describe('Directory to save downloaded images to'),
});

export const createChatSchema = z.object({
  members: z.array(z.string()).describe('Email addresses of chat members (current user is added automatically)'),
  topic: z.string().optional().describe('Chat topic/title (required for group chats with 3+ members)'),
});

// Tool implementations
export async function listChats(params: z.infer<typeof listChatsSchema>) {
  const { maxItems = 25 } = params;

  const chats = await graphList<Chat>(
    `/me/chats?$expand=members&$select=id,topic,chatType,createdDateTime,lastUpdatedDateTime&$top=${maxItems}`,
    { maxItems }
  );

  return chats.map((c) => ({
    id: c.id,
    topic: c.topic || '(No topic)',
    chatType: c.chatType,
    createdAt: c.createdDateTime,
    lastUpdated: c.lastUpdatedDateTime,
    members: c.members?.map((m) => m.displayName || m.email).filter(Boolean),
  }));
}

export async function listChatMessages(params: z.infer<typeof listChatMessagesSchema>) {
  const { chatId, maxItems = 25 } = params;

  const messages = await graphList<ChatMessage>(
    `/chats/${chatId}/messages?$top=${maxItems}`,
    { maxItems }
  );

  return messages.map((m) => ({
    id: m.id,
    createdAt: m.createdDateTime,
    from: m.from?.user?.displayName || m.from?.application?.displayName || 'Unknown',
    // Stable Azure AD object id of the sender (NOT spoofable, unlike the display
    // name). Null for system/event messages with no `from`. Use this — not
    // `from` — for any authorization decision.
    fromId: m.from?.user?.id ?? m.from?.application?.id ?? null,
    fromType: m.from?.user ? 'user' : m.from?.application ? 'application' : 'unknown',
    content: m.body.content,
    contentType: m.body.contentType,
    // When this is a quote-reply, the quoted message ({ messageId, preview,
    // sender }); null otherwise.
    quotedMessage: extractQuotedMessage(m.attachments),
  }));
}

const IMAGE_EXTENSIONS: Record<string, string> = {
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'image/gif': '.gif',
  'image/webp': '.webp',
  'image/bmp': '.bmp',
};

function sanitizeFilename(name: string): string {
  return name.replace(/[^a-zA-Z0-9._-]+/g, '_');
}

// hostedContents listings often omit contentType — sniff the magic bytes
function extFromBuffer(buffer: Buffer): string | null {
  if (buffer.length >= 4 && buffer[0] === 0x89 && buffer[1] === 0x50) return '.png';
  if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8) return '.jpg';
  if (buffer.length >= 3 && buffer.subarray(0, 3).toString('ascii') === 'GIF') return '.gif';
  if (buffer.length >= 12 && buffer.subarray(8, 12).toString('ascii') === 'WEBP') return '.webp';
  return null;
}

export async function downloadChatImages(params: z.infer<typeof downloadChatImagesSchema>) {
  const { chatId, messageId, maxItems = 25, outputDir } = params;
  const { writeFile, mkdir } = await import('fs/promises');
  const { join } = await import('path');

  await mkdir(outputDir, { recursive: true });

  let messages: ChatMessage[];
  if (messageId) {
    messages = [await graphRequest<ChatMessage>(`/chats/${chatId}/messages/${messageId}`)];
  } else {
    messages = await graphList<ChatMessage>(`/chats/${chatId}/messages?$top=${maxItems}`, { maxItems });
  }

  const files: Array<{
    messageId: string;
    from: string;
    createdAt: string;
    source: 'hostedContent' | 'attachment';
    path: string;
    bytes: number;
  }> = [];
  const errors: Array<{ messageId: string; error: string }> = [];

  for (const m of messages) {
    const from = m.from?.user?.displayName || m.from?.application?.displayName || 'Unknown';

    // Inline pasted images (hosted contents)
    try {
      const hosted = await graphList<HostedContent>(
        `/chats/${chatId}/messages/${m.id}/hostedContents`
      );
      let n = 0;
      for (const hc of hosted) {
        if (hc.contentType && !hc.contentType.startsWith('image/')) continue;
        const buffer = await graphDownload(
          `/chats/${chatId}/messages/${m.id}/hostedContents/${hc.id}/$value`
        );
        n += 1;
        const ext = IMAGE_EXTENSIONS[hc.contentType || ''] || extFromBuffer(buffer) || '.png';
        const outputPath = join(outputDir, `${m.id}-${n}${ext}`);
        await writeFile(outputPath, buffer);
        files.push({
          messageId: m.id,
          from,
          createdAt: m.createdDateTime,
          source: 'hostedContent',
          path: outputPath,
          bytes: buffer.length,
        });
      }
    } catch (err) {
      errors.push({ messageId: m.id, error: err instanceof Error ? err.message : String(err) });
    }

    // Image file attachments (shared via OneDrive/SharePoint reference)
    for (const att of m.attachments ?? []) {
      if (att.contentType !== 'reference' || !att.contentUrl) continue;
      const name = att.name || '';
      if (!/\.(png|jpe?g|gif|webp|bmp|heic)$/i.test(name)) continue;
      try {
        const shareToken =
          'u!' +
          Buffer.from(att.contentUrl)
            .toString('base64')
            .replace(/=+$/, '')
            .replace(/\//g, '_')
            .replace(/\+/g, '-');
        const buffer = await graphDownload(`/shares/${shareToken}/driveItem/content`);
        const outputPath = join(outputDir, `${m.id}-${sanitizeFilename(name)}`);
        await writeFile(outputPath, buffer);
        files.push({
          messageId: m.id,
          from,
          createdAt: m.createdDateTime,
          source: 'attachment',
          path: outputPath,
          bytes: buffer.length,
        });
      } catch (err) {
        errors.push({ messageId: m.id, error: err instanceof Error ? err.message : String(err) });
      }
    }
  }

  return {
    success: true,
    outputDir,
    count: files.length,
    files,
    ...(errors.length > 0 ? { errors } : {}),
  };
}

const MIME_TYPES: Record<string, string> = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp',
  '.bmp': 'image/bmp',
};

export async function sendChatMessage(params: z.infer<typeof sendChatMessageSchema>) {
  const { chatId, content, images = [] } = params;

  let html = content;
  const hostedContents: Array<Record<string, string>> = [];

  if (images.length > 0) {
    const { readFile } = await import('fs/promises');
    const { extname } = await import('path');

    for (let i = 0; i < images.length; i++) {
      const buffer = await readFile(images[i]);
      const ext = extname(images[i]).toLowerCase();
      const mimeType = MIME_TYPES[ext] || (extFromBuffer(buffer) ? MIME_TYPES[extFromBuffer(buffer)!] : undefined);
      if (!mimeType) {
        throw new Error(`Unsupported image type: ${images[i]} (expected png, jpg, gif, webp, or bmp)`);
      }
      const temporaryId = String(i + 1);
      hostedContents.push({
        '@microsoft.graph.temporaryId': temporaryId,
        contentBytes: buffer.toString('base64'),
        contentType: mimeType,
      });
      html += `<p><img src="../hostedContents/${temporaryId}/$value" alt="${sanitizeFilename(images[i].split('/').pop() || 'image')}"></p>`;
    }
  }

  const message = await graphRequest<ChatMessage>(
    `/chats/${chatId}/messages`,
    {
      method: 'POST',
      body: {
        body: {
          contentType: 'html',
          content: html,
        },
        ...(hostedContents.length > 0 ? { hostedContents } : {}),
      },
    }
  );

  return {
    success: true,
    messageId: message.id,
    message: images.length > 0 ? `Message sent with ${images.length} image(s)` : 'Message sent',
  };
}

interface User {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

export const whoamiSchema = z.object({});

/**
 * Identity of the authenticated user. `id` is the stable Azure AD object id —
 * the value to compare against a chat message's `fromId` for authorization.
 */
export async function whoami(_params: z.infer<typeof whoamiSchema>) {
  const me = await graphRequest<User>('/me?$select=id,displayName,mail,userPrincipalName');
  return {
    id: me.id,
    displayName: me.displayName ?? null,
    mail: me.mail ?? null,
    userPrincipalName: me.userPrincipalName ?? null,
  };
}

export async function createChat(params: z.infer<typeof createChatSchema>) {
  const { members, topic } = params;

  // Get current user email
  const currentUser = await graphRequest<User>('/me?$select=id,mail,userPrincipalName');
  const currentUserEmail = currentUser.mail || currentUser.userPrincipalName;

  // Build member bindings using email addresses directly
  const memberBindings = [
    {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${currentUserEmail}')`,
    },
  ];

  for (const email of members) {
    memberBindings.push({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${email}')`,
    });
  }

  // Determine chat type
  const isGroup = memberBindings.length > 2;
  const chatType = isGroup ? 'group' : 'oneOnOne';

  const requestBody: Record<string, unknown> = {
    chatType,
    members: memberBindings,
  };

  if (isGroup && topic) {
    requestBody.topic = topic;
  }

  const chat = await graphRequest<Chat>('/chats', {
    method: 'POST',
    body: requestBody,
  });

  return {
    success: true,
    chatId: chat.id,
    chatType: chat.chatType,
    topic: chat.topic,
    message: isGroup ? 'Group chat created' : 'Chat created',
  };
}
