import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

interface PackageJson {
  name: string;
  version: string;
}

let cachedVersion: string | null = null;

export function getVersion(): string {
  if (cachedVersion) return cachedVersion;

  try {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = dirname(__filename);
    const packagePath = join(__dirname, '..', '..', 'package.json');
    const packageJson: PackageJson = JSON.parse(readFileSync(packagePath, 'utf-8'));
    cachedVersion = packageJson.version;
    return cachedVersion;
  } catch {
    return 'unknown';
  }
}
