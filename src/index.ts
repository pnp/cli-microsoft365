#!/usr/bin/env node

import { cli } from './cli/cli.js';
import { app } from './utils/app.js';

// required to make console.log() in combination with piped output synchronous
// on Windows/in PowerShell so that the output is not trimmed by calling
// process.exit() after executing the command, while the output is still
// being processed; https://github.com/pnp/cli-microsoft365/issues/1266
if ((process.stdout as any)._handle) {
  (process.stdout as any)._handle.setBlocking(true);
}

if (!process.env.CLIMICROSOFT365_NOUPDATE) {
  import('update-notifier').then(updateNotifier => {
    updateNotifier.default({ pkg: app.packageJson() as any }).notify({ defer: false });
  });
}

cli.execute(process.argv.slice(2));