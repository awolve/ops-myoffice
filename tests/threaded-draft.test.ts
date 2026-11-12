// Run with: npm test
// The body of a threaded reply/forward draft (createReply + PATCH).

import { test } from 'node:test';
import assert from 'node:assert/strict';

const { threadedDraftBody } = await import('../src/tools/mail.js');

const QUOTE = '<html><head></head><body><hr><div id="divRplyFwdMsg"><b>From:</b> Toni</div></body></html>';

test('a plain-text reply over an HTML quote stays HTML, so the quote is not sent as raw tags', () => {
  const out = threadedDraftBody(QUOTE, 'html', 'Hej Toni,\n\nTack! <3', false);
  assert.equal(out.contentType, 'HTML');
  assert.equal(out.content, '<html><head></head><body>Hej Toni,<br><br>Tack! &lt;3<hr><div id="divRplyFwdMsg"><b>From:</b> Toni</div></body></html>');
});

test('an HTML reply goes inside the quote document body', () => {
  const out = threadedDraftBody(QUOTE, 'html', '<p>Hej</p>', true);
  assert.equal(out.contentType, 'HTML');
  assert.ok(out.content.startsWith('<html><head></head><body><p>Hej</p><hr>'));
});

test('a plain-text quote with a plain-text reply stays text', () => {
  const out = threadedDraftBody('> old', 'text', 'Hej', false);
  assert.deepEqual(out, { contentType: 'Text', content: 'Hej\n\n> old' });
});

test('without a content type, the quote is sniffed', () => {
  assert.equal(threadedDraftBody(QUOTE, undefined, 'Hej', false).contentType, 'HTML');
});
