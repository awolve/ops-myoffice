/**
 * Unit tests for calendar forward (Graph POST /me/events/{id}/forward).
 * Run with: npm test
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';

const { buildEventForwardBody, forwardEventSchema } = await import('../src/tools/calendar.js');

test('body carries every recipient as a Graph recipient', () => {
  const body = buildEventForwardBody(['a@awolve.ai', 'b@example.com'], 'See you there');
  assert.deepEqual(body, {
    comment: 'See you there',
    toRecipients: [
      { emailAddress: { address: 'a@awolve.ai' } },
      { emailAddress: { address: 'b@example.com' } },
    ],
  });
});

test('no comment sends an empty string, which Graph accepts', () => {
  assert.equal(buildEventForwardBody(['a@awolve.ai']).comment, '');
});

test('schema needs at least one recipient', () => {
  assert.throws(() => forwardEventSchema.parse({ eventId: 'x', to: [] }));
});

test('schema rejects a malformed address before any mail goes out', () => {
  assert.throws(() => forwardEventSchema.parse({ eventId: 'x', to: ['not-an-address'] }));
});

test('schema accepts id, recipients and an optional comment', () => {
  const parsed = forwardEventSchema.parse({ eventId: 'x', to: ['a@awolve.ai'], comment: 'Hi' });
  assert.deepEqual(parsed, { eventId: 'x', to: ['a@awolve.ai'], comment: 'Hi' });
});
