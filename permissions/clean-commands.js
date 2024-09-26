#! /usr/bin/env node

import fs from 'fs';

const commands = JSON.parse(fs.readFileSync('cli-commands.json', 'utf8'));

commands.forEach(c => {
  delete c.skip;
  delete c.errors;
  delete c.comment;
  c.examples.forEach(e => {
    delete e.skip;
    delete e.errors;
    delete e.comment;
  });
});

fs.writeFileSync('cli-commands.json', JSON.stringify(commands, null, 2));