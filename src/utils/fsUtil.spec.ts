import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { fsUtil } from './fsUtil.js';

describe('utils/fsUtil', () => {
  it('should get safe filename when file\'name.txt', () => {
    const result = fsUtil.getSafeFileName('file\'name.txt');
    assert.strictEqual(result, 'file\'\'name.txt');
  });

  it('copies a directory recursively when destination does not exist', () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'fsutil-'));
    const src = path.join(root, 'src');
    const dest = path.join(root, 'dest');

    fs.mkdirSync(src);
    fs.writeFileSync(path.join(src, 'file.txt'), 'content');

    fsUtil.copyRecursiveSync(src, dest);

    assert.strictEqual(fs.existsSync(dest), true);
    assert.strictEqual(fs.readFileSync(path.join(dest, 'file.txt'), 'utf8'), 'content');

    fs.rmSync(root, { recursive: true, force: true });
  });
});