import { z } from 'zod';
import { graphRequest, graphList } from '../utils/graph-client.js';

// Types
interface Site {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
}

interface Drive {
  id: string;
  name: string;
  driveType: string;
  webUrl: string;
  quota?: {
    total: number;
    used: number;
    remaining: number;
  };
}

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
export const listSitesSchema = z.object({
  maxItems: z.number().optional().describe('Maximum number of sites. Default: 50'),
  search: z.string().optional().describe('Search query to filter sites'),
});

export const getSiteSchema = z.object({
  siteId: z.string().describe('Site ID or hostname:site-path (e.g., "contoso.sharepoint.com:/sites/team")'),
});

export const listDrivesSchema = z.object({
  siteId: z.string().describe('Site ID'),
  maxItems: z.number().optional().describe('Maximum number of drives. Default: 50'),
});

export const listDriveFilesSchema = z.object({
  driveId: z.string().describe('Drive ID'),
  path: z.string().optional().describe('Folder path within the drive. Default: root'),
  maxItems: z.number().optional().describe('Maximum number of items. Default: 50'),
});

export const getDriveFileSchema = z.object({
  driveId: z.string().describe('Drive ID'),
  path: z.string().describe('File path within the drive'),
});

export const readDriveFileSchema = z.object({
  driveId: z.string().describe('Drive ID'),
  path: z.string().describe('File path within the drive'),
});

export const searchDriveFilesSchema = z.object({
  driveId: z.string().describe('Drive ID'),
  query: z.string().describe('Search query'),
  maxItems: z.number().optional().describe('Maximum results. Default: 25'),
});

// Tool implementations

export async function listSites(params: z.infer<typeof listSitesSchema>) {
  const { maxItems = 50, search } = params;

  let path: string;
  if (search) {
    path = `/sites?$search="${encodeURIComponent(search)}"&$select=id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime&$top=${maxItems}`;
  } else {
    // Get followed sites as default (sites the user has explicitly followed)
    path = `/me/followedSites?$select=id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime&$top=${maxItems}`;
  }

  const sites = await graphList<Site>(path, { maxItems });

  return sites.map((site) => ({
    id: site.id,
    name: site.name,
    displayName: site.displayName,
    webUrl: site.webUrl,
    description: site.description,
    created: site.createdDateTime,
    modified: site.lastModifiedDateTime,
  }));
}

export async function getSite(params: z.infer<typeof getSiteSchema>) {
  const { siteId } = params;

  const site = await graphRequest<Site>(
    `/sites/${siteId}?$select=id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime`
  );

  return {
    id: site.id,
    name: site.name,
    displayName: site.displayName,
    webUrl: site.webUrl,
    description: site.description,
    created: site.createdDateTime,
    modified: site.lastModifiedDateTime,
  };
}

export async function listDrives(params: z.infer<typeof listDrivesSchema>) {
  const { siteId, maxItems = 50 } = params;

  const path = `/sites/${siteId}/drives?$select=id,name,driveType,webUrl,quota&$top=${maxItems}`;

  const drives = await graphList<Drive>(path, { maxItems });

  return drives.map((drive) => ({
    id: drive.id,
    name: drive.name,
    type: drive.driveType,
    webUrl: drive.webUrl,
    quota: drive.quota
      ? {
          total: drive.quota.total,
          used: drive.quota.used,
          remaining: drive.quota.remaining,
        }
      : undefined,
  }));
}

export async function listDriveFiles(params: z.infer<typeof listDriveFilesSchema>) {
  const { driveId, path = '', maxItems = 50 } = params;

  const apiPath = path
    ? `/drives/${driveId}/root:/${path}:/children`
    : `/drives/${driveId}/root/children`;

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

export async function getDriveFile(params: z.infer<typeof getDriveFileSchema>) {
  const { driveId, path } = params;

  const item = await graphRequest<DriveItem>(
    `/drives/${driveId}/root:/${path}?$select=id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file`
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

export async function readDriveFile(params: z.infer<typeof readDriveFileSchema>) {
  const { driveId, path } = params;

  // Get the file metadata with download URL
  const item = await graphRequest<DriveItem>(
    `/drives/${driveId}/root:/${path}?$select=id,name,size,file,@microsoft.graph.downloadUrl`
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

export async function searchDriveFiles(params: z.infer<typeof searchDriveFilesSchema>) {
  const { driveId, query, maxItems = 25 } = params;

  const path = `/drives/${driveId}/root/search(q='${encodeURIComponent(query)}')?$select=id,name,size,webUrl,file,folder&$top=${maxItems}`;

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

// Download schemas and functions

export const downloadFromUrlSchema = z.object({
  url: z.string().describe('SharePoint URL to download (e.g., https://tenant.sharepoint.com/sites/sitename/path/to/file.jpg)'),
  outputPath: z.string().describe('Local file path to save the downloaded file'),
});

export const downloadDriveFileSchema = z.object({
  driveId: z.string().describe('Drive ID'),
  path: z.string().describe('File path within the drive'),
  outputPath: z.string().describe('Local file path to save the downloaded file'),
});

/**
 * Parse a SharePoint URL to extract site hostname, site path, and file path.
 * Supports:
 * - https://tenant.sharepoint.com/sites/sitename/Shared Documents/path/file.jpg
 * - https://tenant-my.sharepoint.com/personal/user_domain_com/Documents/file.pdf
 */
function parseSharePointUrl(url: string): { siteId: string; filePath: string } | null {
  try {
    const parsed = new URL(url);
    const hostname = parsed.hostname;
    const pathParts = decodeURIComponent(parsed.pathname).split('/').filter(Boolean);

    // Handle /sites/sitename/... pattern
    if (pathParts[0] === 'sites' && pathParts.length >= 2) {
      const siteName = pathParts[1];
      // The document library name is usually the 3rd part (e.g., "Shared Documents" or "Delade dokument")
      // Everything after that is the file path within the library
      const siteId = `${hostname}:/sites/${siteName}`;

      // The file path is everything after the site name
      // Skip the document library name for the site ID lookup, but include it for file path
      const filePath = '/' + pathParts.slice(2).join('/');

      return { siteId, filePath };
    }

    // Handle /personal/user_domain_com/... pattern (OneDrive for Business)
    if (pathParts[0] === 'personal' && pathParts.length >= 2) {
      const personalFolder = pathParts[1];
      const siteId = `${hostname}:/personal/${personalFolder}`;
      const filePath = '/' + pathParts.slice(2).join('/');

      return { siteId, filePath };
    }

    return null;
  } catch {
    return null;
  }
}

/**
 * Download a file from a SharePoint URL to a local path.
 */
export async function downloadFromUrl(params: z.infer<typeof downloadFromUrlSchema>) {
  const { url, outputPath } = params;
  const { writeFile } = await import('fs/promises');

  const parsed = parseSharePointUrl(url);
  if (!parsed) {
    throw new Error('Invalid SharePoint URL. Expected format: https://tenant.sharepoint.com/sites/sitename/path/to/file');
  }

  const { siteId, filePath } = parsed;

  // First, get the site to find its drive
  const site = await graphRequest<{ id: string }>(`/sites/${siteId}?$select=id`);

  // Get the default drive for this site
  const drive = await graphRequest<{ id: string }>(`/sites/${site.id}/drive?$select=id`);

  // Now get the file with download URL
  // The filePath includes the document library name, so we need to handle that
  const item = await graphRequest<DriveItem>(
    `/drives/${drive.id}/root:${filePath}?$select=id,name,size,file,@microsoft.graph.downloadUrl`
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

/**
 * Download a file from a SharePoint drive by drive ID and path.
 */
export async function downloadDriveFile(params: z.infer<typeof downloadDriveFileSchema>) {
  const { driveId, path, outputPath } = params;
  const { writeFile } = await import('fs/promises');

  // Get the file metadata with download URL
  const item = await graphRequest<DriveItem>(
    `/drives/${driveId}/root:/${path}?$select=id,name,size,file,@microsoft.graph.downloadUrl`
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
