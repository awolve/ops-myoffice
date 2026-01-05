# Shared OneDrive & SharePoint Access

## What

Add ability to access shared OneDrive folders (files shared with the user) and SharePoint document libraries through the MCP server.

## Why

Currently only personal OneDrive is accessible. Users need to work with shared team files and SharePoint sites for collaboration.

## Requirements

- [ ] List drives shared with the user
- [ ] List SharePoint sites the user has access to
- [ ] Browse files in shared drives and SharePoint document libraries
- [ ] Read files from shared locations
- [ ] Search across shared drives and SharePoint

## Approach

Extend the existing OneDrive tools pattern using Microsoft Graph API. Add new endpoints for `/me/drive/sharedWithMe`, `/sites`, and `/drives/{drive-id}`. Update auth scopes to include `Sites.Read.All`. Follow existing tool structure in `src/tools/`.

## Tasks

- [x] Add SharePoint scopes to auth config (`Sites.Read.All`)
- [x] Create `sharepoint.ts` tool module with site/drive listing
- [x] Add shared drive listing to OneDrive tools
- [x] Add drive-id based file operations (list, read, search)
- [x] Register new tools in index.ts
- [ ] Test with real SharePoint sites

## Notes

- Delegated permissions only (user context, not app-only)
- SharePoint sites may have different permission levels
- Consider pagination for large document libraries
