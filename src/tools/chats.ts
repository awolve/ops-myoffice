import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

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

interface ChatMessage {
  id: string;
  createdDateTime: string;
  from?: ChatMessageFrom;
  body: {
    contentType: string;
    content: string;
  };
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
    content: m.body.content,
    contentType: m.body.contentType,
  }));
}

export async function sendChatMessage(params: z.infer<typeof sendChatMessageSchema>) {
  const { chatId, content } = params;

  const message = await graphRequest<ChatMessage>(
    `/chats/${chatId}/messages`,
    {
      method: 'POST',
      body: {
        body: {
          contentType: 'html',
          content,
        },
      },
    }
  );

  return {
    success: true,
    messageId: message.id,
    message: 'Message sent',
  };
}

interface User {
  id: string;
  mail?: string;
  userPrincipalName?: string;
}

export async function createChat(params: z.infer<typeof createChatSchema>) {
  const { members, topic } = params;

  // Get current user
  const currentUser = await graphRequest<User>('/me?$select=id');

  // Resolve email addresses to user IDs
  const memberBindings = [
    {
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${currentUser.id}')`,
    },
  ];

  for (const email of members) {
    const user = await graphRequest<User>(`/users/${encodeURIComponent(email)}?$select=id`);
    memberBindings.push({
      '@odata.type': '#microsoft.graph.aadUserConversationMember',
      roles: ['owner'],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${user.id}')`,
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
