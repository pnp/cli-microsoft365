import assert from 'assert';
import fs from 'fs';
import path from 'path';
import os from 'os';
import { fsUtil } from './fsUtil.js';

describe('utils/fsUtil', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'fsutil-test-'));
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  it('copies a file', () => {
    const srcFile = path.join(tmpDir, 'source.txt');
    const destFile = path.join(tmpDir, 'dest.txt');
    fs.writeFileSync(srcFile, 'hello');

    fsUtil.copyRecursiveSync(srcFile, destFile);

    assert.strictEqual(fs.readFileSync(destFile, 'utf8'), 'hello');
  });

  it('copies a directory recursively', () => {
    const srcDir = path.join(tmpDir, 'src');
    const destDir = path.join(tmpDir, 'dest');
    fs.mkdirSync(srcDir);
    fs.writeFileSync(path.join(srcDir, 'file.txt'), 'content');

    fsUtil.copyRecursiveSync(srcDir, destDir);

    assert.strictEqual(fs.readFileSync(path.join(destDir, 'file.txt'), 'utf8'), 'content');
  });

  it('copies a directory with replaceTokens', () => {
    const srcDir = path.join(tmpDir, 'src');
    const destDir = path.join(tmpDir, 'dest');
    fs.mkdirSync(srcDir);
    fs.writeFileSync(path.join(srcDir, 'file.txt'), 'content');

    fsUtil.copyRecursiveSync(srcDir, destDir, (s: string) => s);

    assert.strictEqual(fs.readFileSync(path.join(destDir, 'file.txt'), 'utf8'), 'content');
  });

  it('copies into existing destination directory', () => {
    const srcDir = path.join(tmpDir, 'src');
    const destDir = path.join(tmpDir, 'dest');
    fs.mkdirSync(srcDir);
    fs.mkdirSync(destDir);
    fs.writeFileSync(path.join(srcDir, 'file.txt'), 'content');

    fsUtil.copyRecursiveSync(srcDir, destDir);

    assert.strictEqual(fs.readFileSync(path.join(destDir, 'file.txt'), 'utf8'), 'content');
  });
});
