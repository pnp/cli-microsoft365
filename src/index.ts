#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import url from 'url';
import { Cli } from './cli/Cli.js';
import { telemetry } from './telemetry.js';
import { app } from './utils/app.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

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

fs.realpath(__dirname, async (err: NodeJS.ErrnoException | null, resolvedPath: string): Promise<void> => {
  try {
    const cli: Cli = Cli.getInstance();
    await cli.execute(path.join(resolvedPath, 'm365'), process.argv.slice(2));
  }
  catch (e: any) {
    telemetry.trackException(e);
    process.exit(1);
  }
});