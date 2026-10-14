import { getAccessToken } from '../auth/index.js';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

interface GraphResponse<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

export interface GraphRequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  body?: unknown;
  headers?: Record<string, string>;
}

/**
 * A failed Graph call. The message keeps the long-standing
 * `Graph API error (<status>): <message>` form, which callers (Cortex Web's
 * error classifier among them) match on, and adds what a background caller
 * needs to back off properly: the status and Graph's `Retry-After`.
 */
export class GraphError extends Error {
  readonly status: number;
  /** Graph's `code`, e.g. `ErrorItemNotFound` or `syncStateNotFound`, when it sent one. */
  readonly code?: string;
  /** Seconds to wait before retrying, from the `Retry-After` header; undefined when absent. */
  readonly retryAfter?: number;

  constructor(status: number, message: string, code?: string, retryAfter?: number) {
    super(`Graph API error (${status}): ${message}`);
    this.name = 'GraphError';
    this.status = status;
    this.code = code;
    this.retryAfter = retryAfter;
  }
}

/** `Retry-After` is either delay-seconds or an HTTP date (RFC 9110 §10.2.3). */
export function parseRetryAfter(value: string | null, now: number = Date.now()): number | undefined {
  if (!value) return undefined;
  const trimmed = value.trim();
  if (/^\d+$/.test(trimmed)) return Number(trimmed);
  const at = Date.parse(trimmed);
  if (Number.isNaN(at)) return undefined;
  return Math.max(0, Math.ceil((at - now) / 1000));
}

/** Builds the GraphError for a non-OK response, reading its body once. */
export async function toGraphError(response: Response): Promise<GraphError> {
  const errorText = await response.text();
  let errorMessage = errorText;
  let code: string | undefined;

  try {
    const errorJson = JSON.parse(errorText);
    errorMessage = errorJson.error?.message || errorText;
    code = errorJson.error?.code;
  } catch {
    // Not JSON: keep the raw text as the message.
  }

  return new GraphError(response.status, errorMessage, code, parseRetryAfter(response.headers.get('retry-after')));
}

/**
 * `path` is relative to the v1.0 root (`/me/messages`), or an absolute Graph
 * URL as Graph hands back in `@odata.nextLink` and `@odata.deltaLink`.
 */
export async function graphRequest<T>(
  path: string,
  options: GraphRequestOptions = {}
): Promise<T> {
  const { method = 'GET', body, headers = {} } = options;

  const accessToken = await getAccessToken();

  const url = path.startsWith('https://') ? path : `${GRAPH_BASE_URL}${path}`;
  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!response.ok) {
    throw await toGraphError(response);
  }

  // Handle 204 No Content and 202 Accepted (e.g., sendMail)
  if (response.status === 204 || response.status === 202) {
    return {} as T;
  }

  return response.json() as Promise<T>;
}

export interface GraphUploadOptions {
  contentType?: string;
}

export async function graphUpload<T>(
  path: string,
  content: Buffer | Uint8Array,
  options: GraphUploadOptions = {}
): Promise<T> {
  const { contentType = 'application/octet-stream' } = options;

  const accessToken = await getAccessToken();

  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': contentType,
    },
    body: content,
  });

  if (!response.ok) {
    throw await toGraphError(response);
  }

  return response.json() as Promise<T>;
}

interface UploadSession {
  uploadUrl: string;
  expirationDateTime: string;
}

// Resumable upload for large files (up to 250GB)
// Chunk size must be multiple of 320 KiB (327,680 bytes)
const CHUNK_SIZE = 320 * 1024 * 10; // 3.2 MB chunks

export async function graphUploadLarge<T>(
  path: string,
  content: Buffer | Uint8Array,
  onProgress?: (uploaded: number, total: number) => void
): Promise<T> {
  const accessToken = await getAccessToken();
  const fileSize = content.length;

  // Step 1: Create upload session
  const sessionResponse = await fetch(`${GRAPH_BASE_URL}${path}:/createUploadSession`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      item: {
        '@microsoft.graph.conflictBehavior': 'rename',
      },
    }),
  });

  if (!sessionResponse.ok) {
    const errorText = await sessionResponse.text();
    throw new Error(`Failed to create upload session: ${errorText}`);
  }

  const session = (await sessionResponse.json()) as UploadSession;

  // Step 2: Upload in chunks
  let uploadedBytes = 0;

  while (uploadedBytes < fileSize) {
    const chunkEnd = Math.min(uploadedBytes + CHUNK_SIZE, fileSize);
    const chunk = content.slice(uploadedBytes, chunkEnd);

    const chunkResponse = await fetch(session.uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': chunk.length.toString(),
        'Content-Range': `bytes ${uploadedBytes}-${chunkEnd - 1}/${fileSize}`,
      },
      body: chunk,
    });

    if (!chunkResponse.ok) {
      const errorText = await chunkResponse.text();
      throw new Error(`Upload chunk failed: ${errorText}`);
    }

    uploadedBytes = chunkEnd;

    if (onProgress) {
      onProgress(uploadedBytes, fileSize);
    }

    // Final chunk returns the completed item
    if (uploadedBytes >= fileSize) {
      return chunkResponse.json() as Promise<T>;
    }
  }

  throw new Error('Upload completed but no response received');
}

/**
 * Download file content from Graph API.
 * Uses the /content endpoint which returns the file bytes directly.
 */
export async function graphDownload(path: string): Promise<Buffer> {
  const accessToken = await getAccessToken();

  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    redirect: 'follow',
  });

  if (!response.ok) {
    throw await toGraphError(response);
  }

  const arrayBuffer = await response.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

export async function graphList<T>(
  path: string,
  options: { maxItems?: number } = {}
): Promise<T[]> {
  const { maxItems = 100 } = options;
  const items: T[] = [];
  let nextLink: string | undefined = path;

  while (nextLink && items.length < maxItems) {
    const url: string = nextLink.startsWith('http')
      ? nextLink.replace(GRAPH_BASE_URL, '')
      : nextLink;

    const response: GraphResponse<T> = await graphRequest<GraphResponse<T>>(url);

    if (response.value) {
      items.push(...response.value);
    }

    nextLink = response['@odata.nextLink'];
  }

  return items.slice(0, maxItems);
}
