import assert from 'assert';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fsUtil } from './fsUtil.js';

describe('utils/fsUtil', () => {
  it('should get safe filename when file\'name.txt', () => {
    const result = fsUtil.getSafeFileName('file\'name.txt');
    assert.strictEqual(result, 'file\'\'name.txt');
  });

  describe('copyRecursiveSync', () => {
    let tmpDir: string;

    beforeEach(() => {
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'fsUtil-test-'));
    });

    afterEach(() => {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    });

    it('copies a directory recursively', () => {
      const srcDir = path.join(tmpDir, 'src');
      const destDir = path.join(tmpDir, 'dest');
      fs.mkdirSync(srcDir);
      fs.writeFileSync(path.join(srcDir, 'file.txt'), 'hello');

      fsUtil.copyRecursiveSync(srcDir, destDir);

      assert.strictEqual(fs.existsSync(destDir), true);
      assert.strictEqual(fs.readFileSync(path.join(destDir, 'file.txt'), 'utf8'), 'hello');
    });

    it('copies a directory recursively when destination already exists', () => {
      const srcDir = path.join(tmpDir, 'src');
      const destDir = path.join(tmpDir, 'dest');
      fs.mkdirSync(srcDir);
      fs.mkdirSync(destDir);
      fs.writeFileSync(path.join(srcDir, 'file.txt'), 'hello');

      fsUtil.copyRecursiveSync(srcDir, destDir);

      assert.strictEqual(fs.readFileSync(path.join(destDir, 'file.txt'), 'utf8'), 'hello');
    });
  });
});