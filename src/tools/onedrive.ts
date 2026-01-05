import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface DriveItem {
  id: string;
  name: string;
  size?: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  '@microsoft.graph.downloadUrl'?: string;
}

// Schemas
export const listFilesSchema = z.object({
  path: z.string().optional().describe('Folder path (e.g., "Documents/Projects"). Default: root'),
  maxItems: z.number().optional().describe('Maximum number of items. Default: 50'),
});

export const getFileSchema = z.object({
  path: z.string().describe('File path (e.g., "Documents/report.pdf")'),
});

export const searchFilesSchema = z.object({
  query: z.string().describe('Search query'),
  maxItems: z.number().optional().describe('Maximum results. Default: 25'),
});

export const readFileContentSchema = z.object({
  path: z.string().describe('File path to read (works best with text files)'),
});

export const createFolderSchema = z.object({
  name: z.string().describe('Folder name'),
  parentPath: z.string().optional().describe('Parent folder path. Default: root'),
});

// Tool implementations
export async function listFiles(params: z.infer<typeof listFilesSchema>) {
  const { path = '', maxItems = 50 } = params;

  const apiPath = path
    ? `/me/drive/root:/${path}:/children`
    : '/me/drive/root/children';

  const fullPath = `${apiPath}?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,folder,file&$top=${maxItems}`;

  const items = await graphList<DriveItem>(fullPath, { maxItems });

  return items.map((item) => ({
    id: item.id,
    name: item.name,
    type: item.folder ? 'folder' : 'file',
    size: item.size,
    mimeType: item.file?.mimeType,
    childCount: item.folder?.childCount,
    created: item.createdDateTime,
    modified: item.lastModifiedDateTime,
    webUrl: item.webUrl,
  }));
}

export async function getFile(params: z.infer<typeof getFileSchema>) {
  const { path } = params;

  const item = await graphRequest<DriveItem>(
    `/me/drive/root:/${path}?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file`
  );

  return {
    id: item.id,
    name: item.name,
    size: item.size,
    mimeType: item.file?.mimeType,
    created: item.createdDateTime,
    modified: item.lastModifiedDateTime,
    webUrl: item.webUrl,
  };
}

export async function searchFiles(params: z.infer<typeof searchFilesSchema>) {
  const { query, maxItems = 25 } = params;

  const path = `/me/drive/root/search(q='${encodeURIComponent(query)}')?$select=id,name,size,webUrl,file,folder&$top=${maxItems}`;

  const items = await graphList<DriveItem>(path, { maxItems });

  return items.map((item) => ({
    id: item.id,
    name: item.name,
    type: item.folder ? 'folder' : 'file',
    size: item.size,
    mimeType: item.file?.mimeType,
    webUrl: item.webUrl,
  }));
}

export async function readFileContent(params: z.infer<typeof readFileContentSchema>) {
  const { path } = params;

  // Get the file metadata with download URL
  const item = await graphRequest<DriveItem>(
    `/me/drive/root:/${path}?$select=id,name,size,file,@microsoft.graph.downloadUrl`
  );

  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error('Could not get download URL for file');
  }

  // Check file size (limit to 1MB for text content)
  if (item.size && item.size > 1024 * 1024) {
    throw new Error('File too large to read directly. Use download URL instead.');
  }

  // Fetch the content
  const response = await fetch(downloadUrl);
  if (!response.ok) {
    throw new Error(`Failed to download file: ${response.statusText}`);
  }

  const content = await response.text();

  return {
    name: item.name,
    size: item.size,
    mimeType: item.file?.mimeType,
    content,
  };
}

export async function createFolder(params: z.infer<typeof createFolderSchema>) {
  const { name, parentPath = '' } = params;

  const apiPath = parentPath
    ? `/me/drive/root:/${parentPath}:/children`
    : '/me/drive/root/children';

  const created = await graphRequest<DriveItem>(apiPath, {
    method: 'POST',
    body: {
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename',
    },
  });

  return {
    success: true,
    folderId: created.id,
    name: created.name,
    webUrl: created.webUrl,
  };
}
