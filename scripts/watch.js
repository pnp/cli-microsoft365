import { spawn } from 'node:child_process';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const tsgoPath = resolve(__dirname, '..', 'node_modules', '.bin', 'tsgo');

const tsgo = spawn(tsgoPath, ['--watch'], { stdio: 'pipe' });

tsgo.stdout.on('data', (data) => {
  const output = data.toString();
  process.stdout.write(output);

  if (output.includes('Found 0 errors.')) {
    const cmd = spawn(process.execPath, [resolve(__dirname, 'write-all-commands.js')], { stdio: 'inherit' });
    cmd.on('error', (err) => console.error('Failed to run write-all-commands:', err));
  }
});

tsgo.stderr.on('data', (data) => {
  process.stderr.write(data);
});

tsgo.on('close', (code) => {
  process.exit(code);
});
