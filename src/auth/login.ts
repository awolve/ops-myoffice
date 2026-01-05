#!/usr/bin/env node
/**
 * Standalone login script for initial authentication
 * Run with: npm run login
 */

import { authenticateWithDeviceCode } from './device-code.js';

async function main() {
  console.log('Personal M365 MCP - Authentication');
  console.log('===================================\n');

  try {
    await authenticateWithDeviceCode();
    console.log('Authentication successful!');
    console.log('You can now use the M365 MCP with Claude Code.');
    process.exit(0);
  } catch (error) {
    console.error('Authentication failed:', error);
    process.exit(1);
  }
}

main();
