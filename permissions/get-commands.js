#! /usr/bin/env node

import fs from 'fs';
import path from 'path';

const docsPath = '../docs/docs/cmd';
// commands to exclude from permissions checking because they don't
// issue authenticated requests
const commandsToExclude = [
  'adaptivecard send',
  'app open',
  'cli consent',
  'cli doctor',
  'cli issue',
  'cli reconsent',
  'cli completion clink update',
  'cli completion pwsh setup',
  'cli completion pwsh update',
  'cli completion sh setup',
  'cli completion sh update',
  'cli config get',
  'cli config list',
  'cli config reset',
  'cli config set',
  'connection list',
  'connection remove',
  'connection set',
  'connection use',
  'context init',
  'context remove',
  'context option list',
  'context option remove',
  'context option set',
  'docs',
  'login',
  'logout',
  'request',
  'setup',
  'status',
  'version'
];

const commands = [];

function getAllMdxFiles(dirPath, arrayOfFiles = []) {
  const files = fs.readdirSync(dirPath);

  files.forEach((file) => {
    const filePath = path.join(dirPath, file);
    if (fs.statSync(filePath).isDirectory()) {
      arrayOfFiles = getAllMdxFiles(filePath, arrayOfFiles);
    }
    else if (file.endsWith('.mdx')) {
      arrayOfFiles.push(filePath);
    }
  });

  return arrayOfFiles;
}

const mdxFiles = getAllMdxFiles(docsPath);
mdxFiles.forEach(file => {
  console.log(`Processing ${file}`);
  const content = fs.readFileSync(file, 'utf8');
  const commandMatch = content.match(/m365 ([a-z0-9- ]+) \[options\]/);
  if (!commandMatch) {
    return;
  }

  const command = commandMatch[1];
  if (commandsToExclude.includes(command)) {
    return;
  }

  const posExamples = content.indexOf('## Examples');
  const posEnd = content.indexOf('## ', posExamples + 1);
  const examplesString = content.substring(posExamples, posEnd > -1 ? posEnd : undefined);
  const exampleRegex = /```[^\n]*\n(.*?)```/gs;
  const examples = [];
  for (const match of examplesString.matchAll(exampleRegex)) {
    let cmd = match[1].trim();
    // if it's a remove command, and it doesn't have --force,
    // add it to avoid confirmation prompts
    if (cmd.indexOf(' remove') > -1 && cmd.indexOf('--force') < 0) {
      cmd += ' --force';
    }

    examples.push(cmd);
  }

  commands.push({
    command,
    examples: examples.map(example => {
      return { cmd: example.replace(/m365 /g, '') }
    })
  });
});

fs.writeFileSync('cli-commands.json', JSON.stringify(commands, null, 2));