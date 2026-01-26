import { z } from 'zod';
import { readFile } from 'fs/promises';
import { basename } from 'path';
import { graphRequest, graphList, graphUpload, graphUploadLarge } from '../utils/graph-client.js';

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

export const listSharedWithMeSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of items. Default: 50'),
});

export const uploadFileSchema = z.object({
  localPath: z.string().describe('Local file path to upload'),
  remotePath: z
    .string()
    .optional()
    .describe('Destination path in OneDrive (e.g., "Documents/file.pdf"). If omitted, uploads to root with original filename'),
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

interface SharedDriveItem extends DriveItem {
  remoteItem?: {
    id: string;
    name: string;
    parentReference?: {
      driveId: string;
      driveType: string;
    };
  };
}

export async function listSharedWithMe(params: z.infer<typeof listSharedWithMeSchema>) {
  const { maxItems = 50 } = params;

  const path = `/me/drive/sharedWithMe?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,folder,file,remoteItem&$top=${maxItems}`;

  const items = await graphList<SharedDriveItem>(path, { maxItems });

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
    // Include remote item info for accessing the file via its source drive
    remoteDriveId: item.remoteItem?.parentReference?.driveId,
    remoteDriveType: item.remoteItem?.parentReference?.driveType,
    remoteItemId: item.remoteItem?.id,
  }));
}

// MIME type detection from file extension
const MIME_TYPES: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.txt': 'text/plain',
  '.csv': 'text/csv',
  '.json': 'application/json',
  '.xml': 'application/xml',
  '.html': 'text/html',
  '.htm': 'text/html',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.zip': 'application/zip',
  '.mp4': 'video/mp4',
  '.mp3': 'audio/mpeg',
};

function getMimeType(filename: string): string {
  const ext = filename.toLowerCase().match(/\.[^.]+$/)?.[0] || '';
  return MIME_TYPES[ext] || 'application/octet-stream';
}

export async function uploadFile(params: z.infer<typeof uploadFileSchema>) {
  const { localPath, remotePath } = params;

  // Read the local file
  const content = await readFile(localPath);

  // Determine the destination path
  const filename = basename(localPath);
  const destPath = remotePath || filename;

  // Get MIME type
  const contentType = getMimeType(filename);

  // Choose upload method based on file size
  // Simple upload for files <= 4MB, resumable for larger
  const MAX_SIMPLE_UPLOAD = 4 * 1024 * 1024;
  let uploaded: DriveItem;

  if (content.length <= MAX_SIMPLE_UPLOAD) {
    // Simple upload (single request)
    uploaded = await graphUpload<DriveItem>(
      `/me/drive/root:/${destPath}:/content`,
      content,
      { contentType }
    );
  } else {
    // Resumable upload for large files
    uploaded = await graphUploadLarge<DriveItem>(
      `/me/drive/root:/${destPath}`,
      content
    );
  }

  return {
    success: true,
    id: uploaded.id,
    name: uploaded.name,
    size: uploaded.size,
    webUrl: uploaded.webUrl,
    mimeType: uploaded.file?.mimeType,
  };
}

// Download schema and function

export const downloadFileSchema = z.object({
  path: z.string().describe('File path in OneDrive (e.g., "Documents/report.pdf")'),
  outputPath: z.string().describe('Local file path to save the downloaded file'),
});

/**
 * Download a file from OneDrive to a local path.
 */
export async function downloadFile(params: z.infer<typeof downloadFileSchema>) {
  const { path, outputPath } = params;
  const { writeFile } = await import('fs/promises');

  // Get the file metadata with download URL
  const item = await graphRequest<DriveItem>(
    `/me/drive/root:/${path}?$select=id,name,size,file,@microsoft.graph.downloadUrl`
  );

  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (!downloadUrl) {
    throw new Error('Could not get download URL for file');
  }

  // Download the file
  const response = await fetch(downloadUrl);
  if (!response.ok) {
    throw new Error(`Failed to download file: ${response.statusText}`);
  }

  const buffer = Buffer.from(await response.arrayBuffer());
  await writeFile(outputPath, buffer);

  return {
    success: true,
    name: item.name,
    size: item.size,
    mimeType: item.file?.mimeType,
    outputPath,
    bytesWritten: buffer.length,
  };
}
