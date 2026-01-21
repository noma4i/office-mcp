#!/usr/bin/env node

import { startServer } from './lib/server.js';

async function main() {
  try {
    await startServer();
  } catch (error) {
    console.error("Server error:", error);
    process.exit(1);
  }
}

main();
