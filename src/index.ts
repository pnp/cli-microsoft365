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

updateNotifier({ pkg: packageJSON }).notify({ defer: false });

fs.realpath(__dirname, (err: NodeJS.ErrnoException | null, resolvedPath: string): void => {
  try {
    const cli: Cli = Cli.getInstance();
    cli.execute(path.join(resolvedPath, 'm365'), process.argv.slice(2));
  }
  catch (e) {
    appInsights.trackException({
      exception: e
    });
    appInsights.flush();
    process.exit(1);
  }
});