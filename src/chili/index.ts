#!/usr/bin/env node

import { chili } from './chili';

try {
  (async () => await chili.startConversation(process.argv.slice(2)))();
}
catch (err) {
  console.error(`🛑 An error has occurred while searching documentation: ${err}`);
}