# Version Management

## What

Add proper version management to the MCP server so the version displayed in debug_info reflects the actual package version, and can be easily updated.

## Why

Currently the version is hardcoded as "0.1.0" in multiple places. Need a single source of truth for the version that's visible when calling debug_info, making it easy to verify which version is deployed.

## Requirements

- [ ] Single source of truth for version (package.json)
- [ ] debug_info shows version from package.json
- [ ] Server registration uses same version
- [ ] Version visible in startup logs

## Approach

Read version from package.json at runtime using Node's module resolution or fs. Update debug_info and server registration to use this dynamic version. Add version to startup console output.

## Tasks

- [x] Create version utility to read from package.json
- [x] Update server registration to use dynamic version
- [x] Update debug_info to use dynamic version
- [x] Add version to startup log output
- [x] Bump version to 0.2.0 for the SharePoint release

## Notes

- Could use `import pkg from '../package.json'` with assert or read file
- Keep it simple - no need for git commit hash unless requested
