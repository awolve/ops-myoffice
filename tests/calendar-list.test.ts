/**
 * Unit tests for `calendar list` rows (toEventRow).
 * Run with: npm test
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';

const { toEventRow } = await import('../src/tools/calendar.js');

const base = {
  id: 'e1',
  subject: 'Weekly sync',
  start: { dateTime: '2026-11-05T12:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-11-05T12:30:00.0000000', timeZone: 'UTC' },
};

test('a repeating occurrence carries its type, series and Teams join link', () => {
  const row = toEventRow({
    ...base,
    type: 'occurrence',
    seriesMasterId: 'm1',
    isOnlineMeeting: true,
    onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc' },
  });
  assert.equal(row.type, 'occurrence');
  assert.equal(row.seriesMasterId, 'm1');
  assert.equal(row.joinUrl, 'https://teams.microsoft.com/l/meetup-join/abc');
});

test('a single event has no series and no link when Graph sends null', () => {
  const row = toEventRow({ ...base, type: 'singleInstance', seriesMasterId: null, onlineMeeting: null });
  assert.equal(row.type, 'singleInstance');
  assert.equal(row.seriesMasterId, undefined);
  assert.equal(row.joinUrl, undefined);
});

test('the fields callers already read keep their names', () => {
  const row = toEventRow({ ...base, location: { displayName: 'Room 4' }, responseStatus: { response: 'accepted' } });
  assert.deepEqual(
    [row.id, row.subject, row.start, row.end, row.timeZone, row.location, row.response],
    ['e1', 'Weekly sync', base.start.dateTime, base.end.dateTime, 'UTC', 'Room 4', 'accepted'],
  );
});
