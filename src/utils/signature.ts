import { existsSync, readFileSync } from 'fs';
import { homedir } from 'os';
import { join } from 'path';

const SIGNATURE_FILE = join(homedir(), '.config', 'ops-personal-m365-mcp', 'signature.html');

export function getSignature(): string | null {
  try {
    if (existsSync(SIGNATURE_FILE)) {
      const content = readFileSync(SIGNATURE_FILE, 'utf-8');
      return content.trim() || null;
    }
  } catch {
    // Ignore errors, return null
  }
  return null;
}
