import fs from 'fs';
import path from 'path';
import url from 'url';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));
const repoRoot = path.join(__dirname, '..', '..', '..', '..');
const allCommandsPath = path.join(repoRoot, 'allCommandsFull.json');
const outputPath = path.join(repoRoot, 'skills', 'clim365', 'references', 'commands.txt');

const commands = JSON.parse(fs.readFileSync(allCommandsPath, 'utf8'));

const lines = commands
  .map(cmd => `${cmd.name}|${cmd.description}`)
  .sort();

fs.writeFileSync(outputPath, lines.join('\n') + '\n', 'utf8');

console.log(`Generated ${lines.length} commands in skills/clim365/references/commands.txt`);
