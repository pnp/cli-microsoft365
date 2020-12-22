#!/usr/bin/env node

import * as fs from 'fs';
import * as path from 'path';
import * as updateNotifier from 'update-notifier';
import appInsights from './appInsights';
import { Cli } from './cli';

const packageJSON = require('../package.json');

appInsights.trackEvent({
  name: 'started'
});

// required to make console.log() in combination with piped output synchronous
// on Windows/in PowerShell so that the output is not trimmed by calling
// process.exit() after executing the command, while the output is still
// being processed; https://github.com/pnp/cli-microsoft365/issues/1266
if ((process.stdout as any)._handle) {
  (process.stdout as any)._handle.setBlocking(true);
}

updateNotifier({ pkg: packageJSON }).notify({ defer: false });

fs.realpath(__dirname, (err: NodeJS.ErrnoException | null, resolvedPath: string): void => {
  try {
    const cli: Cli = Cli.getInstance();
    cli
      .execute(path.join(resolvedPath, 'm365'), process.argv.slice(2))
  }
  catch (e) {
    appInsights.trackException({
      exception: e
    });
    appInsights.flush();
    process.exit(1);
  }
});