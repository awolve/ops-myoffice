/**
 * Output formatter for CLI - provides human-readable output
 */

// Helper to truncate strings
function truncate(str: string, maxLen: number): string {
  if (!str) return '';
  if (str.length <= maxLen) return str;
  return str.substring(0, maxLen - 3) + '...';
}

// Helper to format date
function formatDate(dateStr: string | undefined): string {
  if (!dateStr) return '';
  try {
    const date = new Date(dateStr);
    return date.toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
    });
  } catch {
    return dateStr;
  }
}

// Helper to pad string
function pad(str: string, len: number): string {
  if (str.length >= len) return str.substring(0, len);
  return str + ' '.repeat(len - str.length);
}

// Format email list
function formatMailList(data: { value?: unknown[] }): string {
  const emails = data.value as Array<{
    id?: string;
    from?: { emailAddress?: { address?: string } };
    subject?: string;
    receivedDateTime?: string;
    isRead?: boolean;
  }>;

  if (!emails || emails.length === 0) {
    return 'No emails found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('FROM', 25)} ${pad('SUBJECT', 40)} ${pad('DATE', 16)} READ`);
  lines.push('-'.repeat(90));

  for (const email of emails) {
    const from = truncate(email.from?.emailAddress?.address || '', 24);
    const subject = truncate(email.subject || '(no subject)', 39);
    const date = formatDate(email.receivedDateTime);
    const read = email.isRead ? '✓' : '○';
    lines.push(`${pad(from, 25)} ${pad(subject, 40)} ${pad(date, 16)} ${read}`);
  }

  const unreadCount = emails.filter(e => !e.isRead).length;
  lines.push('');
  lines.push(`${emails.length} emails (${unreadCount} unread)`);

  return lines.join('\n');
}

// Format calendar events
function formatCalendarList(data: { value?: unknown[] }): string {
  const events = data.value as Array<{
    subject?: string;
    start?: { dateTime?: string };
    end?: { dateTime?: string };
    location?: { displayName?: string };
    isOnlineMeeting?: boolean;
  }>;

  if (!events || events.length === 0) {
    return 'No events found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('DATE/TIME', 20)} ${pad('SUBJECT', 35)} ${pad('LOCATION', 25)}`);
  lines.push('-'.repeat(85));

  for (const event of events) {
    const dateTime = formatDate(event.start?.dateTime);
    const subject = truncate(event.subject || '(no subject)', 34);
    let location = event.location?.displayName || '';
    if (event.isOnlineMeeting && !location) location = 'Teams Meeting';
    location = truncate(location, 24);
    lines.push(`${pad(dateTime, 20)} ${pad(subject, 35)} ${pad(location, 25)}`);
  }

  lines.push('');
  lines.push(`${events.length} events`);

  return lines.join('\n');
}

// Format task lists
function formatTaskLists(data: { value?: unknown[] }): string {
  const lists = data.value as Array<{
    id?: string;
    displayName?: string;
  }>;

  if (!lists || lists.length === 0) {
    return 'No task lists found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('NAME', 40)} ID`);
  lines.push('-'.repeat(80));

  for (const list of lists) {
    const name = truncate(list.displayName || '', 39);
    const id = truncate(list.id || '', 38);
    lines.push(`${pad(name, 40)} ${id}`);
  }

  lines.push('');
  lines.push(`${lists.length} task lists`);

  return lines.join('\n');
}

// Format tasks
function formatTasks(data: { value?: unknown[] }): string {
  const tasks = data.value as Array<{
    title?: string;
    status?: string;
    importance?: string;
    dueDateTime?: { dateTime?: string };
  }>;

  if (!tasks || tasks.length === 0) {
    return 'No tasks found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('STATUS', 8)} ${pad('TITLE', 40)} ${pad('DUE', 16)} IMP`);
  lines.push('-'.repeat(75));

  for (const task of tasks) {
    const status = task.status === 'completed' ? '✓' : '○';
    const title = truncate(task.title || '', 39);
    const due = task.dueDateTime?.dateTime ? formatDate(task.dueDateTime.dateTime) : '';
    const imp = task.importance === 'high' ? '!' : task.importance === 'low' ? '↓' : '';
    lines.push(`${pad(status, 8)} ${pad(title, 40)} ${pad(due, 16)} ${imp}`);
  }

  lines.push('');
  const completed = tasks.filter(t => t.status === 'completed').length;
  lines.push(`${tasks.length} tasks (${completed} completed)`);

  return lines.join('\n');
}

// Format files/folders
function formatFiles(data: { value?: unknown[] }): string {
  const items = data.value as Array<{
    name?: string;
    folder?: unknown;
    size?: number;
    lastModifiedDateTime?: string;
  }>;

  if (!items || items.length === 0) {
    return 'No files found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('TYPE', 6)} ${pad('NAME', 40)} ${pad('SIZE', 12)} MODIFIED`);
  lines.push('-'.repeat(85));

  for (const item of items) {
    const type = item.folder ? '📁' : '📄';
    const name = truncate(item.name || '', 39);
    const size = item.folder ? '' : formatSize(item.size || 0);
    const modified = formatDate(item.lastModifiedDateTime);
    lines.push(`${pad(type, 6)} ${pad(name, 40)} ${pad(size, 12)} ${modified}`);
  }

  lines.push('');
  lines.push(`${items.length} items`);

  return lines.join('\n');
}

function formatSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return (bytes / Math.pow(1024, i)).toFixed(i > 0 ? 1 : 0) + ' ' + units[i];
}

// Format teams list
function formatTeams(data: { value?: unknown[] }): string {
  const teams = data.value as Array<{
    id?: string;
    displayName?: string;
    description?: string;
  }>;

  if (!teams || teams.length === 0) {
    return 'No teams found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('NAME', 35)} DESCRIPTION`);
  lines.push('-'.repeat(80));

  for (const team of teams) {
    const name = truncate(team.displayName || '', 34);
    const desc = truncate(team.description || '', 43);
    lines.push(`${pad(name, 35)} ${desc}`);
  }

  lines.push('');
  lines.push(`${teams.length} teams`);

  return lines.join('\n');
}

// Format chats list
function formatChats(data: { value?: unknown[] }): string {
  const chats = data.value as Array<{
    id?: string;
    topic?: string;
    chatType?: string;
    lastUpdatedDateTime?: string;
  }>;

  if (!chats || chats.length === 0) {
    return 'No chats found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('TYPE', 10)} ${pad('TOPIC', 40)} LAST ACTIVE`);
  lines.push('-'.repeat(75));

  for (const chat of chats) {
    const type = chat.chatType || 'unknown';
    const topic = truncate(chat.topic || '(unnamed chat)', 39);
    const lastActive = formatDate(chat.lastUpdatedDateTime);
    lines.push(`${pad(type, 10)} ${pad(topic, 40)} ${lastActive}`);
  }

  lines.push('');
  lines.push(`${chats.length} chats`);

  return lines.join('\n');
}

// Format planner plans
function formatPlans(data: { value?: unknown[] }): string {
  const plans = data.value as Array<{
    id?: string;
    title?: string;
  }>;

  if (!plans || plans.length === 0) {
    return 'No plans found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('TITLE', 45)} ID`);
  lines.push('-'.repeat(90));

  for (const plan of plans) {
    const title = truncate(plan.title || '', 44);
    const id = truncate(plan.id || '', 43);
    lines.push(`${pad(title, 45)} ${id}`);
  }

  lines.push('');
  lines.push(`${plans.length} plans`);

  return lines.join('\n');
}

// Format planner tasks
function formatPlannerTasks(data: { value?: unknown[] }): string {
  const tasks = data.value as Array<{
    title?: string;
    percentComplete?: number;
    priority?: number;
    dueDateTime?: string;
    assigneePriority?: string;
  }>;

  if (!tasks || tasks.length === 0) {
    return 'No tasks found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('STATUS', 8)} ${pad('TITLE', 40)} ${pad('DUE', 16)} PRI`);
  lines.push('-'.repeat(75));

  for (const task of tasks) {
    const pct = task.percentComplete || 0;
    const status = pct === 100 ? '✓' : pct > 0 ? '◐' : '○';
    const title = truncate(task.title || '', 39);
    const due = task.dueDateTime ? formatDate(task.dueDateTime) : '';
    const pri = task.priority === 1 ? '!!!' : task.priority === 3 ? '!!' : task.priority === 5 ? '!' : '';
    lines.push(`${pad(status, 8)} ${pad(title, 40)} ${pad(due, 16)} ${pri}`);
  }

  lines.push('');
  const completed = tasks.filter(t => t.percentComplete === 100).length;
  lines.push(`${tasks.length} tasks (${completed} completed)`);

  return lines.join('\n');
}

// Format auth status
function formatAuthStatus(data: { authenticated?: boolean; user?: { displayName?: string; mail?: string } | null }): string {
  if (data.authenticated && data.user) {
    return `Authenticated as: ${data.user.displayName || data.user.mail || 'Unknown'}`;
  }
  return 'Not authenticated. Run: myoffice login';
}

// Format contacts
function formatContacts(data: { value?: unknown[] }): string {
  const contacts = data.value as Array<{
    displayName?: string;
    emailAddresses?: Array<{ address?: string }>;
    companyName?: string;
  }>;

  if (!contacts || contacts.length === 0) {
    return 'No contacts found.';
  }

  const lines: string[] = [];
  lines.push(`${pad('NAME', 30)} ${pad('EMAIL', 30)} COMPANY`);
  lines.push('-'.repeat(85));

  for (const contact of contacts) {
    const name = truncate(contact.displayName || '', 29);
    const email = truncate(contact.emailAddresses?.[0]?.address || '', 29);
    const company = truncate(contact.companyName || '', 23);
    lines.push(`${pad(name, 30)} ${pad(email, 30)} ${company}`);
  }

  lines.push('');
  lines.push(`${contacts.length} contacts`);

  return lines.join('\n');
}

// Generic object formatter
function formatObject(data: unknown): string {
  return JSON.stringify(data, null, 2);
}

// Main format function
export function formatOutput(data: unknown, toolName?: string): string {
  if (data === null || data === undefined) {
    return 'No data';
  }

  // Try to detect the data type and format appropriately
  const obj = data as Record<string, unknown>;

  // Check for specific tool outputs
  if (toolName === 'auth_status') {
    return formatAuthStatus(data as { authenticated?: boolean; user?: { displayName?: string; mail?: string } | null });
  }

  // Check for list responses (most Graph API responses have a 'value' array)
  if (Array.isArray(obj.value)) {
    // Try to determine the type of list
    const items = obj.value as Array<Record<string, unknown>>;
    if (items.length === 0) {
      return 'No items found.';
    }

    const sample = items[0];

    // Email detection
    if ('receivedDateTime' in sample && 'from' in sample) {
      return formatMailList(obj as { value: unknown[] });
    }
    // Calendar event detection
    if ('start' in sample && 'end' in sample && 'subject' in sample) {
      return formatCalendarList(obj as { value: unknown[] });
    }
    // To Do task list detection
    if ('displayName' in sample && !('folder' in sample) && !('emailAddresses' in sample) && !('members' in sample)) {
      // Could be task lists or teams - check for wellknownListName
      if ('wellknownListName' in sample || items.every(i => 'displayName' in i && !('description' in i))) {
        return formatTaskLists(obj as { value: unknown[] });
      }
    }
    // To Do tasks detection
    if ('status' in sample && 'title' in sample && !('percentComplete' in sample)) {
      return formatTasks(obj as { value: unknown[] });
    }
    // OneDrive/SharePoint files detection
    if ('name' in sample && ('folder' in sample || 'file' in sample || 'size' in sample)) {
      return formatFiles(obj as { value: unknown[] });
    }
    // Teams detection
    if ('displayName' in sample && 'description' in sample && !('emailAddresses' in sample)) {
      return formatTeams(obj as { value: unknown[] });
    }
    // Chats detection
    if ('chatType' in sample) {
      return formatChats(obj as { value: unknown[] });
    }
    // Planner plans detection
    if ('title' in sample && 'owner' in sample && !('percentComplete' in sample)) {
      return formatPlans(obj as { value: unknown[] });
    }
    // Planner tasks detection
    if ('percentComplete' in sample && 'title' in sample) {
      return formatPlannerTasks(obj as { value: unknown[] });
    }
    // Contacts detection
    if ('emailAddresses' in sample) {
      return formatContacts(obj as { value: unknown[] });
    }
  }

  // Default: JSON output
  return formatObject(data);
}
