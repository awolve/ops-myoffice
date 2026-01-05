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

export async function graphRequest<T>(
  path: string,
  options: GraphRequestOptions = {}
): Promise<T> {
  const { method = 'GET', body, headers = {} } = options;

  const accessToken = await getAccessToken();

  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!response.ok) {
    const errorText = await response.text();
    let errorMessage: string;

    try {
      const errorJson = JSON.parse(errorText);
      errorMessage = errorJson.error?.message || errorText;
    } catch {
      errorMessage = errorText;
    }

    throw new Error(`Graph API error (${response.status}): ${errorMessage}`);
  }

  // Handle 204 No Content
  if (response.status === 204) {
    return {} as T;
  }

  return response.json() as Promise<T>;
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
