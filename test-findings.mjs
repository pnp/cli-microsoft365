import path from 'path';
import { fileURLToPath } from 'url';
import command from './dist/m365/spfx/commands/project/project-upgrade.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Stub getProjectRoot
command.getProjectRoot = () => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-fieldcustomizer-react');

const logger = {
  log: async (msg) => {console.log('Number of findings:', msg.length); msg.forEach((f, i) => console.log(`${i+1}. ${f.id}`));},
  logRaw: async (msg) => {},
  logToStderr: async (msg) => {}
};

command.action(logger, {options: {toVersion: '1.9.1', output: 'json'}})
  .then(() => console.log('Done'))
  .catch(err => console.error('Error:', err));
