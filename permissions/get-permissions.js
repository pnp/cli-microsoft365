#! /usr/bin/env node

import { execSync } from 'child_process';
import createDebug from 'debug';
import fs from 'fs';
import chalk from 'chalk';

const debug = createDebug('permissions');
const commandsFile = 'cli-commands.json';
const permissionsWatchTimeoutMs = 5000;

function error(message) {
  console.error(chalk.red(message));
}

async function toggleDevProxyRecording(enabled) {
  await fetch('http://localhost:8897/proxy', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
    },
    body: JSON.stringify({ recording: enabled }),
  });
}

function runCliCommand(cmd) {
  try {
    const output = execSync(`HTTP_PROXY=http://127.0.0.1:8000 m365 ${cmd}`, { stdio: 'pipe' });
    debug(output.toString());
  }
  catch (e) {
    error(`Error running command '${cmd}': ${e.stderr.toString()}`);
    process.exit(1);
  }
}

async function detectMinimalScopes(cmd) {
  let watcher;
  try {
    let newFile = undefined;
    const startTime = Date.now();
    watcher = fs.watch(process.cwd(), { persistent: false }, (event, filename) => {
      if (event === 'rename' && filename.startsWith('MinimalPermissionsPlugin_JsonReporter')) {
        newFile = filename;
      }
    });

    await toggleDevProxyRecording(true);
    runCliCommand(cmd);
    await toggleDevProxyRecording(false);

    // Wait for the new file to be detected
    while (!newFile && (Date.now() - startTime) < permissionsWatchTimeoutMs) {
      debug('Waiting for new file to be detected...');
      await new Promise(resolve => setTimeout(resolve, 100));
    }

    if (!newFile) {
      error('    Timed out');
      return { errors: ['Timed out'] };
    }

    const report = JSON.parse(fs.readFileSync(newFile, 'utf8'));

    if (report.errors && report.errors.length > 0) {
      error(`    Errors detected while running command '${cmd}':`);
      report.errors.forEach(err => error(`      ${err}`));
      return { errors: report.errors.map(e => e.startsWith('- ') ? e.substring(2) : e) };
    }

    return { scopes: report.minimalPermissions };
  }
  catch (e) {
    error(e);
    return undefined;
  }
  finally {
    if (watcher) {
      watcher.close();
    }
  }
}

async function main() {
  console.log('Loading commands...');
  const commands = JSON.parse(fs.readFileSync(commandsFile, 'utf8'));
  console.log('Detecting minimal permissions...');

  for (const command of commands) {
    if (command.skip) {
      continue;
    }

    console.log(`- ${command.command}`);

    for (const cmd of command.examples) {
      if (cmd.skip) {
        continue;
      }

      console.log(`  - ${cmd.cmd}`);
      try {
        const { scopes, errors } = await detectMinimalScopes(cmd.cmd);
        cmd.scopes = scopes?.join(' ');
        cmd.errors = errors;
        // temp
        cmd.skip = true;
        fs.writeFileSync(commandsFile, JSON.stringify(commands, null, 2));
      }
      catch (e) {
        error(`Error detecting scopes for command '${cmd.cmd}': ${e}`);
      }
    }
  }

  console.log(chalk.green('DONE'));
}

await main();