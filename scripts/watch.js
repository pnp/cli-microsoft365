import { spawn } from 'node:child_process';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const tscPath = resolve(__dirname, '..', 'node_modules', 'typescript-native', 'bin', 'tsc');

const tsc = spawn(process.execPath, [tscPath, '--watch'], { stdio: 'pipe' });

tsc.stdout.on('data', (data) => {
  const output = data.toString();
  process.stdout.write(output);

  if (output.includes('Found 0 errors.')) {
    const cmd = spawn(process.execPath, [resolve(__dirname, 'write-all-commands.js')], { stdio: 'inherit' });
    cmd.on('error', (err) => console.error('Failed to run write-all-commands:', err));
  }
});

tsc.stderr.on('data', (data) => {
  process.stderr.write(data);
});

tsc.on('close', (code) => {
  process.exit(code);
});
