/**
 * Unit tests for the `mail list` filters (--since, --from, --conversation-id)
 * and the shared --since helpers used by chats and teams.
 * Run with: npm test
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';

const { buildMailFilter, graphDateTime, createdSince } = await import('../src/tools/mail.js');

test('no filter options means no $filter', () => {
  assert.equal(buildMailFilter({}), undefined);
});

test('--unread alone keeps working, led by receivedDateTime', () => {
  assert.equal(buildMailFilter({ unreadOnly: true }), 'receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false');
});

test('--since leads the filter as a UTC time without milliseconds', () => {
  assert.equal(
    buildMailFilter({ since: '2026-10-12T14:02:00Z' }),
    'receivedDateTime ge 2026-10-12T14:02:00Z'
  );
});

test('several senders are or-ed inside parentheses', () => {
  assert.equal(
    buildMailFilter({ since: '2026-10-12T14:02:00Z', from: ['a@x.no', 'b@x.no'] }),
    "receivedDateTime ge 2026-10-12T14:02:00Z and (from/emailAddress/address eq 'a@x.no' or from/emailAddress/address eq 'b@x.no')"
  );
});

test('one sender and a conversation, without --since', () => {
  assert.equal(
    buildMailFilter({ from: ['a@x.no'], conversationId: 'AAQk=' }),
    "receivedDateTime ge 1900-01-01T00:00:00Z and from/emailAddress/address eq 'a@x.no' and conversationId eq 'AAQk='"
  );
});

test('single quotes in values are escaped for OData', () => {
  assert.match(buildMailFilter({ from: ["o'neil@x.no"] })!, /eq 'o''neil@x\.no'/);
});

test('an invalid time is refused before any Graph call', () => {
  assert.throws(() => graphDateTime('yesterday-ish'), /Not a valid time/);
});

test('createdSince compares times, not strings', () => {
  // As strings "…00.500Z" < "…00Z"; as times it is half a second later.
  assert.equal(createdSince('2026-10-12T14:02:00.500Z', '2026-10-12T14:02:00Z'), true);
  assert.equal(createdSince('2026-10-12T14:01:59.999Z', '2026-10-12T14:02:00Z'), false);
  assert.equal(createdSince('2020-01-01T00:00:00Z', undefined), true);
});
