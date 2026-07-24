import assert from 'assert';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fsUtil } from './fsUtil.js';

// Resolve the OS temp directory once, at module load time, before any test runs.
// On Windows, os.tmpdir() reads process.env directly, so tests that replace
// process.env (e.g. Auth.spec.ts) can make a later os.tmpdir() call resolve to an
// invalid 'undefined\temp' path. Capturing it here keeps this suite immune to that.
const tmpBaseDir = os.tmpdir();

describe('utils/fsUtil', () => {
  it('should get safe filename when file\'name.txt', () => {
    const result = fsUtil.getSafeFileName('file\'name.txt');
    assert.strictEqual(result, 'file\'\'name.txt');
  });

  describe('copyRecursiveSync', () => {
    let tmpDir: string;

    beforeEach(() => {
      tmpDir = fs.mkdtempSync(path.join(tmpBaseDir, 'fsUtil-test-'));
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