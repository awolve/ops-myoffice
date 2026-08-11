import { existsSync, readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';

const CONFIG_DIR = join(homedir(), '.config', 'myoffice-mcp');

/**
 * Awolve has two signatures: a standard one with logo and full contact details
 * for new mail, and a minimal single line for replies. `none` sends neither.
 */
export type SignatureStyle = 'standard' | 'minimal' | 'none';

const FILES: Record<Exclude<SignatureStyle, 'none'>, string> = {
  standard: join(CONFIG_DIR, 'signature.html'),
  minimal: join(CONFIG_DIR, 'signature-minimal.html'),
};

export function isSignatureStyle(value: unknown): value is SignatureStyle {
  return value === 'standard' || value === 'minimal' || value === 'none';
}

/**
 * Read a signature. Returns null when the style is `none` or the file has not
 * been installed — in Cortex, `bun scripts/cortex-signature.ts` writes both from
 * the person's own context, and /cortex-doctor reports when they are missing.
 *
 * A missing minimal signature falls back to the standard one rather than
 * sending nothing: an over-formal sign-off beats a reply that looks anonymous.
 */
export function getSignature(style: SignatureStyle = 'standard'): string | null {
  if (style === 'none') return null;
  const candidates = style === 'minimal' ? [FILES.minimal, FILES.standard] : [FILES.standard];
  for (const path of candidates) {
    try {
      if (!existsSync(path)) continue;
      const content = readFileSync(path, 'utf-8').trim();
      if (content) return content;
    } catch {
      // unreadable — try the next candidate, then give up quietly
    }
  }
  return null;
}
