import assert from 'assert';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
import { fsUtil } from './fsUtil.js';
import { sinonUtil } from './sinonUtil.js';

describe('utils/fsUtil', () => {
  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.statSync,
      fs.readdirSync,
      fs.copyFileSync,
      fs.mkdirSync,
      (fsUtil as any).copyRecursiveSync
    ]);
  });

  describe('copyRecursiveSync', () => {
    it('copies a file when src is not a directory', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => false } as fs.Stats);
      const copyFileStub = sinon.stub(fs, 'copyFileSync');

      fsUtil.copyRecursiveSync('/src/file.txt', '/dest/file.txt');

      assert(copyFileStub.calledOnceWith('/src/file.txt', '/dest/file.txt'));
    });

    it('creates destination directory if it does not exist', () => {
      const existsStub = sinon.stub(fs, 'existsSync');
      existsStub.withArgs('/src').returns(true);
      existsStub.withArgs('/dest').returns(false);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => true } as fs.Stats);
      sinon.stub(fs, 'readdirSync').returns([]);
      const mkdirStub = sinon.stub(fs, 'mkdirSync');

      fsUtil.copyRecursiveSync('/src', '/dest');

      assert(mkdirStub.calledOnceWith('/dest'));
    });

    it('copies directory contents recursively', () => {
      const srcDir = path.join('/src');
      const destDir = path.join('/dest');
      const srcChild = path.join('/src', 'child.txt');
      const destChild = path.join('/dest', 'child.txt');

      const existsStub = sinon.stub(fs, 'existsSync');
      existsStub.returns(true);
      const statStub = sinon.stub(fs, 'statSync');
      statStub.withArgs(srcDir).returns({ isDirectory: () => true } as fs.Stats);
      statStub.withArgs(srcChild).returns({ isDirectory: () => false } as fs.Stats);
      sinon.stub(fs, 'readdirSync').returns(['child.txt'] as any);
      sinon.stub(fs, 'mkdirSync');
      const copyFileStub = sinon.stub(fs, 'copyFileSync');

      fsUtil.copyRecursiveSync(srcDir, destDir);

      assert(copyFileStub.calledOnceWith(srcChild, destChild));
    });

    it('applies replaceTokens to destination path', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => false } as fs.Stats);
      const copyFileStub = sinon.stub(fs, 'copyFileSync');
      const replaceTokens = (s: string): string => s.replace('TOKEN', 'value');

      fsUtil.copyRecursiveSync('/src/file.txt', '/dest/TOKEN/file.txt', replaceTokens);

      assert(copyFileStub.calledOnceWith('/src/file.txt', '/dest/value/file.txt'));
    });
  });
});
