import assert from 'assert';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
import { fsUtil } from './fsUtil.js';

describe('utils/fsUtil', () => {
  afterEach(() => {
    sinon.restore();
  });

  describe('copyRecursiveSync', () => {
    it('copies a directory recursively creating dest if it does not exist', () => {
      sinon.stub(fs, 'existsSync')
        .withArgs('src').returns(true)
        .withArgs('dest').returns(false);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => true } as fs.Stats);
      const mkdirStub = sinon.stub(fs, 'mkdirSync');
      sinon.stub(fs, 'readdirSync').returns(['file1.txt'] as any);
      const copyFileStub = sinon.stub(fs, 'copyFileSync');
      (fs.existsSync as sinon.SinonStub)
        .withArgs(path.join('src', 'file1.txt')).returns(true);
      (fs.statSync as sinon.SinonStub)
        .withArgs(path.join('src', 'file1.txt')).returns({ isDirectory: () => false } as fs.Stats);

      fsUtil.copyRecursiveSync('src', 'dest');

      assert(mkdirStub.calledWith('dest'));
      assert(copyFileStub.calledWith(path.join('src', 'file1.txt'), path.join('dest', 'file1.txt')));
    });

    it('copies a directory recursively when dest already exists', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'statSync')
        .withArgs('src').returns({ isDirectory: () => true } as fs.Stats)
        .withArgs(path.join('src', 'child.txt')).returns({ isDirectory: () => false } as fs.Stats);
      const mkdirStub = sinon.stub(fs, 'mkdirSync');
      sinon.stub(fs, 'readdirSync').returns(['child.txt'] as any);
      const copyFileStub = sinon.stub(fs, 'copyFileSync');

      fsUtil.copyRecursiveSync('src', 'dest');

      assert(mkdirStub.notCalled);
      assert(copyFileStub.calledWith(path.join('src', 'child.txt'), path.join('dest', 'child.txt')));
    });

    it('applies replaceTokens to destination path', () => {
      sinon.stub(fs, 'existsSync')
        .withArgs('src').returns(true)
        .withArgs('replaced-dest').returns(false);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => true } as fs.Stats);
      const mkdirStub = sinon.stub(fs, 'mkdirSync');
      sinon.stub(fs, 'readdirSync').returns([] as any);

      fsUtil.copyRecursiveSync('src', 'dest', (s: string) => s === 'dest' ? 'replaced-dest' : s);

      assert(mkdirStub.calledWith('replaced-dest'));
    });

    it('copies a single file', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'statSync').returns({ isDirectory: () => false } as fs.Stats);
      const copyFileStub = sinon.stub(fs, 'copyFileSync');

      fsUtil.copyRecursiveSync('src/file.txt', 'dest/file.txt');

      assert(copyFileStub.calledWith('src/file.txt', 'dest/file.txt'));
    });
  });
});
