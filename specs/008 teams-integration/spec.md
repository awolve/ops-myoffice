# Teams Integration

## What

Add Microsoft Teams support to MyOffice MCP - read channels, post messages, and access chat history.

## Why

Teams is central to M365 collaboration. Users want to:
- Check team channels without switching apps
- Post updates from their workflow
- Search for information shared in Teams
- Get an overview of recent activity

## Requirements

- [ ] List teams the user is a member of
- [ ] List channels in a team
- [ ] Read messages from a channel
- [ ] Post messages to a channel
- [ ] List 1:1 and group chats
- [ ] Read messages from a chat
- [ ] Send messages in a chat

## API Endpoints

| Operation | Method | Endpoint |
|-----------|--------|----------|
| List joined teams | GET | `/me/joinedTeams` |
| Get team details | GET | `/teams/{team-id}` |
| List channels | GET | `/teams/{team-id}/channels` |
| Get channel | GET | `/teams/{team-id}/channels/{channel-id}` |
| List channel messages | GET | `/teams/{team-id}/channels/{channel-id}/messages` |
| Post channel message | POST | `/teams/{team-id}/channels/{channel-id}/messages` |
| List chats | GET | `/me/chats` |
| Get chat | GET | `/chats/{chat-id}` |
| List chat messages | GET | `/chats/{chat-id}/messages` |
| Send chat message | POST | `/chats/{chat-id}/messages` |

## Required Scopes

Add to `src/auth/config.ts`:

```typescript
'Team.ReadBasic.All',      // List teams
'Channel.ReadBasic.All',   // List channels
'ChannelMessage.Read.All', // Read channel messages
'ChannelMessage.Send',     // Post to channels
'Chat.Create',             // Create chats
'Chat.ReadBasic',          // List chats
'Chat.Read',               // Read chat messages
'ChatMessage.Send',        // Send chat messages
```

## Tools

### teams_list
List teams the user is a member of.

**Parameters:** none (or `maxItems`)

**Returns:** Array of `{ id, displayName, description }`

### teams_channels
List channels in a team.

**Parameters:**
- `teamId` (required) - Team ID

**Returns:** Array of `{ id, displayName, description, membershipType }`

### teams_channel_messages
Read recent messages from a channel.

**Parameters:**
- `teamId` (required) - Team ID
- `channelId` (required) - Channel ID
- `maxItems` (optional, default 25) - Max messages to return

**Returns:** Array of messages with sender, content, timestamp

### teams_channel_post
Post a message to a channel.

**Parameters:**
- `teamId` (required) - Team ID
- `channelId` (required) - Channel ID
- `content` (required) - Message content (supports HTML)

**Returns:** Created message object

### chats_list
List user's 1:1 and group chats.

**Parameters:**
- `maxItems` (optional, default 25)

**Returns:** Array of `{ id, topic, chatType, members }`

### chats_messages
Read messages from a chat.

**Parameters:**
- `chatId` (required) - Chat ID
- `maxItems` (optional, default 25)

**Returns:** Array of messages with sender, content, timestamp

### chats_send
Send a message in a chat.

**Parameters:**
- `chatId` (required) - Chat ID
- `content` (required) - Message content

**Returns:** Created message object

### chats_create
Create a new 1:1 or group chat.

**Parameters:**
- `members` (required) - Email addresses of chat members (current user added automatically)
- `topic` (optional) - Chat topic/title (for group chats)

**Returns:** `{ chatId, chatType, topic }`

## Tasks

- [x] Add Teams scopes to `src/auth/config.ts`
- [x] Create `src/tools/teams.ts` with channel operations
- [x] Create `src/tools/chats.ts` with chat operations
- [x] Add tool definitions to `src/index.ts`
- [x] Add tool routing in CallToolRequestSchema handler
- [x] Update README with Teams tools documentation
- [ ] Test with real Teams/chats

## Limitations

1. **Private channels** - User can only see channels they're a member of
2. **Message history** - Channel messages may have limited history access
3. **Rate limiting** - Microsoft throttles Teams APIs more aggressively
4. **Delegated only** - Sending messages requires user context (not app-only)

## Message Format

Messages support HTML content:

```json
{
  "body": {
    "contentType": "html",
    "content": "<p>Hello from MyOffice MCP!</p>"
  }
}
```

## Example Usage

```
"List my teams and show the recent messages in the General channel of the Engineering team"

"Post 'Build completed successfully' to the #deployments channel in DevOps team"

"Search my Teams chats for messages about the Q1 budget"
```

## Notes

- User must re-authenticate after adding new scopes (`npm run login`)
- Teams API returns less data than other Graph APIs - may need multiple calls
- Consider adding a `teams_search` tool later for cross-team message search
