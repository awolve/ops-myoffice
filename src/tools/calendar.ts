import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface Calendar {
  id: string;
  name: string;
  color?: string;
  isDefaultCalendar?: boolean;
  canEdit?: boolean;
  owner?: { name?: string; address?: string };
}

interface Event {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName: string };
  organizer?: { emailAddress: { name?: string; address: string } };
  attendees?: Array<{
    emailAddress: { name?: string; address: string };
    status?: { response: string };
  }>;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  body?: { contentType: string; content: string };
  isAllDay?: boolean;
}

// Schemas
export const listCalendarsSchema = z.object({});

export const listEventsSchema = z.object({
  startDate: z.string().optional().describe('Start date (ISO format). Default: today'),
  endDate: z.string().optional().describe('End date (ISO format). Default: 7 days from now'),
  maxItems: z.number().optional().describe('Maximum number of events. Default: 50'),
  calendarId: z.string().optional().describe('Calendar ID. Default: primary calendar'),
});

export const getEventSchema = z.object({
  eventId: z.string().describe('The ID of the event to retrieve'),
});

export const createEventSchema = z.object({
  subject: z.string().describe('Event subject/title'),
  start: z.string().describe('Start datetime (ISO format, e.g., 2024-01-15T09:00:00)'),
  end: z.string().describe('End datetime (ISO format)'),
  timeZone: z.string().optional().describe('Timezone (e.g., Europe/Stockholm). Default: UTC'),
  location: z.string().optional().describe('Event location'),
  body: z.string().optional().describe('Event description'),
  attendees: z.array(z.string()).optional().describe('List of attendee email addresses'),
  isOnlineMeeting: z.boolean().optional().describe('Create Teams meeting. Default: false'),
  calendarId: z.string().optional().describe('Calendar ID. Default: primary calendar'),
});

export const updateEventSchema = z.object({
  eventId: z.string().describe('The ID of the event to update'),
  subject: z.string().optional().describe('New subject'),
  start: z.string().optional().describe('New start datetime'),
  end: z.string().optional().describe('New end datetime'),
  location: z.string().optional().describe('New location'),
  body: z.string().optional().describe('New description'),
});

export const deleteEventSchema = z.object({
  eventId: z.string().describe('The ID of the event to delete'),
});

// Tool implementations
export async function listCalendars() {
  const path = `/me/calendars?$select=id,name,color,isDefaultCalendar,canEdit,owner`;

  const calendars = await graphList<Calendar>(path, { maxItems: 100 });

  return calendars.map((c) => ({
    id: c.id,
    name: c.name,
    color: c.color,
    isDefault: c.isDefaultCalendar,
    canEdit: c.canEdit,
    owner: c.owner?.address,
  }));
}

export async function listEvents(params: z.infer<typeof listEventsSchema>) {
  const {
    startDate = new Date().toISOString(),
    endDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
    maxItems = 50,
    calendarId,
  } = params;

  // Use specific calendar or default calendarView
  const basePath = calendarId
    ? `/me/calendars/${calendarId}/calendarView`
    : `/me/calendarView`;

  const path = `${basePath}?startDateTime=${startDate}&endDateTime=${endDate}&$select=id,subject,start,end,location,organizer,isOnlineMeeting,onlineMeetingUrl,isAllDay&$orderby=start/dateTime&$top=${maxItems}`;

  const events = await graphList<Event>(path, { maxItems });

  return events.map((e) => ({
    id: e.id,
    subject: e.subject,
    start: e.start.dateTime,
    end: e.end.dateTime,
    timeZone: e.start.timeZone,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress?.address,
    isOnlineMeeting: e.isOnlineMeeting,
    meetingUrl: e.onlineMeetingUrl,
    isAllDay: e.isAllDay,
  }));
}

export async function getEvent(params: z.infer<typeof getEventSchema>) {
  const { eventId } = params;

  const event = await graphRequest<Event>(
    `/me/events/${eventId}?$select=id,subject,start,end,location,organizer,attendees,body,isOnlineMeeting,onlineMeetingUrl`
  );

  return {
    id: event.id,
    subject: event.subject,
    start: event.start.dateTime,
    end: event.end.dateTime,
    timeZone: event.start.timeZone,
    location: event.location?.displayName,
    organizer: event.organizer?.emailAddress?.address,
    attendees: event.attendees?.map((a) => ({
      email: a.emailAddress.address,
      name: a.emailAddress.name,
      response: a.status?.response,
    })),
    body: event.body?.content,
    isOnlineMeeting: event.isOnlineMeeting,
    meetingUrl: event.onlineMeetingUrl,
  };
}

export async function createEvent(params: z.infer<typeof createEventSchema>) {
  const {
    subject,
    start,
    end,
    timeZone = 'UTC',
    location,
    body,
    attendees,
    isOnlineMeeting = false,
    calendarId,
  } = params;

  const eventData = {
    subject,
    start: { dateTime: start, timeZone },
    end: { dateTime: end, timeZone },
    location: location ? { displayName: location } : undefined,
    body: body ? { contentType: 'Text', content: body } : undefined,
    attendees: attendees?.map((email) => ({
      emailAddress: { address: email },
      type: 'required',
    })),
    isOnlineMeeting,
    onlineMeetingProvider: isOnlineMeeting ? 'teamsForBusiness' : undefined,
  };

  // Use specific calendar or default
  const path = calendarId ? `/me/calendars/${calendarId}/events` : '/me/events';

  const created = await graphRequest<Event>(path, {
    method: 'POST',
    body: eventData,
  });

  return {
    success: true,
    eventId: created.id,
    subject: created.subject,
    start: created.start.dateTime,
    meetingUrl: created.onlineMeetingUrl,
  };
}

export async function updateEvent(params: z.infer<typeof updateEventSchema>) {
  const { eventId, subject, start, end, location, body } = params;

  const updates: Record<string, unknown> = {};
  if (subject) updates.subject = subject;
  if (start) updates.start = { dateTime: start, timeZone: 'UTC' };
  if (end) updates.end = { dateTime: end, timeZone: 'UTC' };
  if (location) updates.location = { displayName: location };
  if (body) updates.body = { contentType: 'Text', content: body };

  const updated = await graphRequest<Event>(`/me/events/${eventId}`, {
    method: 'PATCH',
    body: updates,
  });

  return {
    success: true,
    eventId: updated.id,
    subject: updated.subject,
  };
}

export async function deleteEvent(params: z.infer<typeof deleteEventSchema>) {
  const { eventId } = params;

  await graphRequest(`/me/events/${eventId}`, {
    method: 'DELETE',
  });

  return { success: true, message: 'Event deleted' };
}
