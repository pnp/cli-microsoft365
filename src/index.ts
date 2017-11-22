#!/usr/bin/env node

import * as fs from 'fs';
import * as path from 'path';
import config from './config';
import Command from './Command';
import appInsights from './appInsights';

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
      if (cmd instanceof Command ) {
        cmd.init(vorpal);
      }
    }
  });

  vorpal
    .command('version', 'Shows the current version of the CLI')
    .action(function (this: CommandInstance, args: any, cb: () => void) {
      this.log(packageJSON.version);
      cb();
    });

  vorpal.parse(process.argv);

  vorpal
    .delimiter(chalk.red(config.delimiter))
    .show();
});