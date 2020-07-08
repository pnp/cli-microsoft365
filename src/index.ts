#!/usr/bin/env node

import * as fs from 'fs';
import * as path from 'path';
import * as updateNotifier from 'update-notifier';
import Command from './Command';
import appInsights from './appInsights';
import Utils from './Utils';

const packageJSON = require('../package.json');
const vorpal: Vorpal = require('./vorpal-init');

const readdirR = (dir: string): string | string[] => {
  return fs.statSync(dir).isDirectory()
    ? Array.prototype.concat(...fs.readdirSync(dir).map(f => readdirR(path.join(dir, f))))
    : dir;
};

const loadAllCommands = (rootFolder: string): void => {
  const commandsDir: string = path.join(rootFolder, './m365');
  const files: string[] = readdirR(commandsDir) as string[];

  files.forEach(file => {
    if (file.indexOf(`${path.sep}commands${path.sep}`) > -1 &&
      file.endsWith('.js') &&
      !file.endsWith('.spec.js')) {
      try {
        const cmd: any = require(file);
        if (cmd instanceof Command) {
          cmd.init(vorpal);
        }
      }
      catch (e) {
        console.log(e);
      }
    }
  });
};

const loadCommandFromArgs = (args: string[], rootFolder: string): void => {
  if (args.length <= 3) {
    loadAllCommands(rootFolder);
    return;
  }

  const isCompletionCommand: boolean = args.indexOf('completion') > -1;
  if (isCompletionCommand) {
    loadAllCommands(rootFolder);
    return;
  }

  // get the name of the command to be executed from args
  // first two arguments are node and the name of the script
  let cliArgs: string[] = args.slice(2);
  // find the first command argument if any
  // arguments start typically with - or -- but for the 'spo login' command
  // it's the URL of the site to connect to
  const pos: number = cliArgs.findIndex(p => p.startsWith('-') || p.startsWith('https://'));
  if (pos > -1) {
    // remove command args so that what's left is only the command name
    cliArgs = cliArgs.slice(0, pos);
  }

  let commandFilePath = '';
  if (cliArgs.length === 1) {
    commandFilePath = path.join(rootFolder, 'm365', 'commands', `${cliArgs[0]}.js`);
  }
  else {
    if (cliArgs.length === 2) {
      commandFilePath = path.join(rootFolder, 'm365', cliArgs[0], 'commands', `${cliArgs.join('-')}.js`);
    }
    else {
      commandFilePath = path.join(rootFolder, 'm365', cliArgs[0], 'commands', cliArgs[1], cliArgs.slice(1).join('-') + '.js');
    }
  }

  if (!fs.existsSync(commandFilePath)) {
    loadAllCommands(rootFolder);
    return;
  }

  try {
    const cmd: any = require(commandFilePath);
    if (cmd instanceof Command) {
      cmd.init(vorpal);
    }
    else {
      loadAllCommands(rootFolder);
    }
  }
  catch {
    loadAllCommands(rootFolder);
  }
}

appInsights.trackEvent({
  name: 'started'
});

updateNotifier({ pkg: packageJSON }).notify({ defer: false });

fs.realpath(__dirname, (err: NodeJS.ErrnoException | null, resolvedPath: string): void => {
  // disable linux-normalizing args to support JSON and XML values
  vorpal.isCommandArgKeyPairNormalized = false;

  vorpal
    .title('CLI for Microsoft 365')
    .description(packageJSON.description)
    .version(packageJSON.version);

  vorpal
    .command('version', 'Shows the current version of the CLI')
    .action(function (this: CommandInstance, args: any, cb: () => void) {
      this.log(packageJSON.version);
      cb();
    });

  vorpal.pipe((stdout: any): any => {
    return Utils.logOutput(stdout);
  });

  try {
    vorpal.delimiter('');
    vorpal.on('client_command_error', (err?: any): void => {
      process.exit(1);
    });

    if (process.argv.length <= 2) {
      process.argv.push('help');
    }

    loadCommandFromArgs(process.argv, resolvedPath);
    vorpal.parse(process.argv);
  }
  catch (e) {
    appInsights.trackException({
      exception: e
    });
    appInsights.flush();
    process.exit(1);
  }
});