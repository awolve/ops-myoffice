import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';
import { createdSince, graphDateTime } from './mail.js';

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
  lastModifiedDateTime?: string;
  replyToId?: string | null;
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
  since: z.string().optional().describe('Only messages created at or after this time (ISO 8601)'),
  includeReplies: z.boolean().optional().describe('Also return thread replies, fetched for each top-level message read'),
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
  const { teamId, channelId, maxItems = 25, since, includeReplies = false } = params;
  const base = `/teams/${teamId}/channels/${channelId}/messages`;

  // Channel messages take no $filter, so --since is applied here. A reply
  // lives under its top-level message, and a new reply can sit under an old
  // one, so replies are fetched for every top-level message read, not only
  // for new ones. --limit bounds how many that is.
  const sinceIso = since ? graphDateTime(since) : undefined;
  const roots = await graphList<ChatMessage>(`${base}?$top=${Math.min(maxItems, 50)}`, { maxItems });

  const all: ChatMessage[] = [...roots];
  if (includeReplies) {
    for (const root of roots) {
      all.push(...(await graphList<ChatMessage>(`${base}/${root.id}/replies?$top=50`, { maxItems: 200 })));
    }
  }

  return all
    .filter((m) => createdSince(m.createdDateTime, sinceIso))
    .sort((a, b) => Date.parse(b.createdDateTime) - Date.parse(a.createdDateTime))
    .map((m) => ({
      id: m.id,
      createdAt: m.createdDateTime,
      lastModified: m.lastModifiedDateTime,
      from: m.from?.user?.displayName || m.from?.application?.displayName || 'Unknown',
      // Stable Azure AD object id of the sender; match people on this, not on `from`.
      fromId: m.from?.user?.id ?? m.from?.application?.id ?? null,
      // Set on a reply: the id of the top-level message it answers.
      replyToId: m.replyToId ?? null,
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
