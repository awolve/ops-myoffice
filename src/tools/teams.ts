import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface Team {
  id: string;
  displayName: string;
  description?: string;
}

interface Channel {
  id: string;
  displayName: string;
  description?: string;
  membershipType?: string;
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
  subject?: string;
}

// Schemas
export const listTeamsSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of teams to return. Default: 50'),
});

export const listChannelsSchema = z.object({
  teamId: z.string().describe('The ID of the team'),
});

export const listChannelMessagesSchema = z.object({
  teamId: z.string().describe('The ID of the team'),
  channelId: z.string().describe('The ID of the channel'),
  maxItems: z.number().optional().describe('Maximum number of messages to return. Default: 25'),
});

export const postChannelMessageSchema = z.object({
  teamId: z.string().describe('The ID of the team'),
  channelId: z.string().describe('The ID of the channel'),
  content: z.string().describe('Message content (supports HTML)'),
});

// Tool implementations
export async function listTeams(params: z.infer<typeof listTeamsSchema>) {
  const { maxItems = 50 } = params;

  const teams = await graphList<Team>(
    '/me/joinedTeams?$select=id,displayName,description',
    { maxItems }
  );

  return teams.map((t) => ({
    id: t.id,
    displayName: t.displayName,
    description: t.description,
  }));
}

export async function listChannels(params: z.infer<typeof listChannelsSchema>) {
  const { teamId } = params;

  const channels = await graphList<Channel>(
    `/teams/${teamId}/channels?$select=id,displayName,description,membershipType`
  );

  return channels.map((c) => ({
    id: c.id,
    displayName: c.displayName,
    description: c.description,
    membershipType: c.membershipType,
  }));
}

export async function listChannelMessages(params: z.infer<typeof listChannelMessagesSchema>) {
  const { teamId, channelId, maxItems = 25 } = params;

  const messages = await graphList<ChatMessage>(
    `/teams/${teamId}/channels/${channelId}/messages?$top=${maxItems}`,
    { maxItems }
  );

  return messages.map((m) => ({
    id: m.id,
    createdAt: m.createdDateTime,
    from: m.from?.user?.displayName || m.from?.application?.displayName || 'Unknown',
    content: m.body.content,
    contentType: m.body.contentType,
    subject: m.subject,
  }));
}

export async function postChannelMessage(params: z.infer<typeof postChannelMessageSchema>) {
  const { teamId, channelId, content } = params;

  const message = await graphRequest<ChatMessage>(
    `/teams/${teamId}/channels/${channelId}/messages`,
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
    message: 'Message posted to channel',
  };
}
