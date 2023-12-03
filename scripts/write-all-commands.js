import fs from 'fs';
import path from 'path';
import url, { pathToFileURL } from 'url';
import Command from '../dist/Command.js';
import { Cli } from '../dist/cli/Cli.js';
import { fsUtil } from '../dist/utils/fsUtil.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
const commandsFolder = path.join(__dirname, '..', 'dist', 'm365');
const commandHelpFolder = path.join(commandsFolder, '..', '..', 'docs', 'docs', 'cmd');

async function loadAllCommands() {
  const files = fsUtil.readdirR(commandsFolder);
  const cli = Cli.getInstance();

  await Promise.all(files.map(async (filePath) => {
    const file = pathToFileURL(filePath).toString();
    if (file.indexOf(`/commands/`) > -1 &&
      file.indexOf(`/assets/`) < 0 &&
      file.endsWith('.js') &&
      !file.endsWith('.spec.js')) {

      const command = await import(file);
      if (command.default instanceof Command) {
        const helpFilePath = path.relative(commandHelpFolder, getCommandHelpFilePath(command.default.name));
        cli.commands.push(Cli.getCommandInfo(command.default, path.relative(commandsFolder, filePath), helpFilePath));
      }
    }
  }));

  cli.commands.forEach(c => {
    delete c.command;
    delete c.defaultProperties;
  });
  // this file is used by command completion
  fs.writeFileSync('allCommandsFull.json', JSON.stringify(cli.commands));

  cli.commands.forEach(c => {
    delete c.options;
  });
  // this file is use for regular command execution
  fs.writeFileSync('allCommands.json', JSON.stringify(cli.commands));
}

function getCommandHelpFilePath(commandName) {
  const commandNameWords = commandName.split(' ');
  const pathChunks = [];

  if (commandNameWords.length === 1) {
    pathChunks.push(`${commandNameWords[0]}.mdx`);
  }
  else {
    if (commandNameWords.length === 2) {
      pathChunks.push(commandNameWords[0], `${commandNameWords.join('-')}.mdx`);
    }
    else {
      pathChunks.push(commandNameWords[0], commandNameWords[1], commandNameWords.slice(1).join('-') + '.mdx');
    }
  }

  const helpFilePath = path.join(commandHelpFolder, ...pathChunks);
  return helpFilePath;
}

loadAllCommands();