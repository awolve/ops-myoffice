/**
 * Unit tests for GraphError and the Retry-After parsing a background caller
 * (the Cortex mail feed) backs off on.
 * Run with: npm test
 */

import { test } from 'node:test';
import assert from 'node:assert/strict';

const { GraphError, parseRetryAfter, toGraphError } = await import('../src/utils/graph-client.js');

test('message keeps the long-standing form callers match on', () => {
  const e = new GraphError(503, 'Command Concurrency Limit Reached');
  assert.equal(e.message, 'Graph API error (503): Command Concurrency Limit Reached');
  assert.equal(e.status, 503);
  assert.ok(e instanceof Error);
});

test('Retry-After as delay-seconds', () => {
  assert.equal(parseRetryAfter('30'), 30);
  assert.equal(parseRetryAfter(' 0 '), 0);
});

test('Retry-After as an HTTP date becomes seconds from now, never negative', () => {
  const now = Date.parse('2026-10-14T10:00:00Z');
  assert.equal(parseRetryAfter('Wed, 14 Oct 2026 10:01:30 GMT', now), 90);
  assert.equal(parseRetryAfter('Wed, 14 Oct 2026 09:59:00 GMT', now), 0);
});

test('missing or garbled Retry-After is undefined', () => {
  assert.equal(parseRetryAfter(null), undefined);
  assert.equal(parseRetryAfter('soon'), undefined);
});

test('toGraphError reads message, code and Retry-After from a Graph response', async () => {
  const res = new Response(JSON.stringify({ error: { code: 'ApplicationThrottled', message: 'Too many requests' } }), {
    status: 429,
    headers: { 'Retry-After': '12' },
  });
  const e = await toGraphError(res);
  assert.equal(e.message, 'Graph API error (429): Too many requests');
  assert.equal(e.code, 'ApplicationThrottled');
  assert.equal(e.retryAfter, 12);
});

test('toGraphError keeps a non-JSON body as the message', async () => {
  const e = await toGraphError(new Response('Bad gateway', { status: 502 }));
  assert.equal(e.message, 'Graph API error (502): Bad gateway');
  assert.equal(e.code, undefined);
  assert.equal(e.retryAfter, undefined);
});
