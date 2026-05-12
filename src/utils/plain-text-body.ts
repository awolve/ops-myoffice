// Warn when a plain-text Graph body (calendar event, To Do task)
// looks like HTML. Outlook renders such bodies literally — recipients
// see `<p>...</p>` as text. Contrast with `mail send`, which is HTML
// by default. We warn rather than block: writer keeps control, but the
// asymmetry is made visible.

const HTML_TAG_RE = /<(p|br|div|html|body|b|i|u|strong|em|a|ul|ol|li|h[1-6])\b|<\/\w+>/i;

export function warnIfHtmlBody(body: string | undefined, surface: string): void {
  if (!body) return;
  if (!HTML_TAG_RE.test(body)) return;
  console.error(
    `[myoffice] warning: ${surface}: --body looks like HTML. Calendar and task bodies are sent as plain text — ` +
    `recipients will see literal HTML tags. Use plain text with \\n for line breaks instead. ` +
    `(mail bodies are HTML-by-default; calendar and tasks are not.)`
  );
}
