/**
 * Library entry: `import { graphRequest, GraphError } from '@awolve/myoffice/graph'`.
 *
 * For long-running callers that need Graph with the signed-in user's
 * delegated login but not a process per call — the Cortex mail feed
 * (cortex-services spec 011) is the first. Keep this surface small: every
 * export here is a promise to callers outside this repo.
 */

export { graphRequest, GraphError, parseRetryAfter } from './utils/graph-client.js';
export type { GraphRequestOptions } from './utils/graph-client.js';
export { getAccessToken, isAuthenticated, getCurrentUser } from './auth/index.js';
