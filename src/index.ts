#!/usr/bin/env node

import * as fs from 'fs';
import * as path from 'path';
import config from './config';
import Command from './Command';
import appInsights from './appInsights';
import Utils from './Utils';
import { autocomplete } from './autocomplete';

const packageJSON = require('../package.json');
const vorpal: Vorpal = require('./vorpal-init'),
  chalk = vorpal.chalk;

const readdirR = (dir: string): string | string[] => {
  return fs.statSync(dir).isDirectory()
    ? Array.prototype.concat(...fs.readdirSync(dir).map(f => readdirR(path.join(dir, f))))
    : dir;
}

appInsights.trackEvent({
  name: 'started'
});

fs.realpath(__dirname, (err: NodeJS.ErrnoException, resolvedPath: string): void => {
  const commandsDir: string = path.join(resolvedPath, './o365');
  const files: string[] = readdirR(commandsDir) as string[];

  files.forEach(file => {
    if (file.indexOf(`${path.sep}commands${path.sep}`) > -1 &&
      file.indexOf('.spec.js') === -1 &&
      file.indexOf('.js.map') === -1) {
      const cmd: any = require(file);
      if (cmd instanceof Command) {
        cmd.init(vorpal);
      }
    }
  });

  if (process.argv.indexOf('--completion:clink:generate') > -1) {
    console.log(autocomplete.getClinkCompletion(vorpal));
    process.exit();
  }
  if (process.argv.indexOf('--completion:sh:generate') > -1) {
    autocomplete.generateShCompletion(vorpal);
    process.exit();
  }
  if (process.argv.indexOf('--completion:sh:setup') > -1) {
    autocomplete.generateShCompletion(vorpal);
    autocomplete.setupShCompletion();
    process.exit();
  }

  // disable linux-normalizing args to support JSON and XML values
  vorpal.isCommandArgKeyPairNormalized = false;

  vorpal
    .command('version', 'Shows the current version of the CLI')
    .action(function (this: CommandInstance, args: any, cb: () => void) {
      this.log(packageJSON.version);
      cb();
    });

  vorpal.pipe((stdout: any): any => {
    return Utils.logOutput(stdout);
  });

  let v: Vorpal | null = null;
  try {
    if (process.argv.length > 2) {
      vorpal.delimiter('');
    }
    v = vorpal.parse(process.argv);

    // if no command has been passed/match, run immersive mode
    if (!v._command) {
      vorpal
        .delimiter(chalk.red(config.delimiter + ' '))
        .show();
    }
  }
  catch (e) {
    appInsights.trackException({
      exception: e
    });
    appInsights.flush();
    process.exit(0);
  }
});