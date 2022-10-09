import * as assert from 'assert';
import * as fs from 'fs';
import path = require('path');
import * as sinon from 'sinon';
import { cache } from './cache';
import { sinonUtil } from './sinonUtil';

describe('utils/cache', () => {
  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.mkdirSync,
      fs.writeFile,
      fs.readdir,
      fs.stat,
      fs.unlink,
      cache.clearExpired
    ]);
  });

  describe('getValue', () => {
    it(`returns undefined when the specified value doesn't exist in cache`, () => {
      sinon.stub(fs, 'existsSync').returns(false);
      assert.strictEqual(cache.getValue('key'), undefined);
    });

    it('returns the specified value from cache', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'readFileSync').returns('value');
      assert.strictEqual(cache.getValue('key'), 'value');
    });

    it('returns undefined if an error occurs while reading cache', () => {
      sinon.stub(fs, 'existsSync').returns(true);
      sinon.stub(fs, 'readFileSync').throws();
      assert.strictEqual(cache.getValue('key'), undefined);
    });

    it('clears expired values', () => {
      const clearExpiredSpy = sinon.spy(cache, 'clearExpired');
      sinon.stub(fs, 'existsSync').returns(false);
      cache.getValue('key');

      assert(clearExpiredSpy.called);
    });
  });

  describe('setValue', () => {
    it('clears expired values', () => {
      const clearExpiredSpy = sinon.spy(cache, 'clearExpired');
      sinon.stub(fs, 'mkdirSync').throws();
      cache.setValue('key', 'value');

      assert(clearExpiredSpy.called);
    });

    it(`doesn't fail when creating the cache folder fails`, () => {
      sinon.stub(fs, 'mkdirSync').throws();
      const writeFilesSpy = sinon.spy(fs, 'writeFile');
      cache.setValue('key', 'value');

      assert(writeFilesSpy.notCalled);
    });

    it(`doesn't fail when writing value to cache file fails`, (done) => {
      sinon.stub(fs, 'mkdirSync').callsFake(() => undefined);
      sinon
        .stub(fs, 'writeFile')
        .callsFake(() => {
          done();
        })
        .callsArgWith(2, 'error');
      try {
        cache.setValue('key', 'value');
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`writes value to cache in a cache file`, (done) => {
      sinon.stub(fs, 'mkdirSync').callsFake(() => undefined);
      sinon.stub(fs, 'writeFile').callsFake(() => {
        done();
      });
      try {
        cache.setValue('key', 'value');
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  describe('clearExpired', () => {
    it(`doesn't fail when reading the cache folder fails`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, 'error');
      try {
        cache.clearExpired(() => {
          done();
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail when the cache folder is empty`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, []);
      try {
        cache.clearExpired(() => {
          done();
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`skips directories while clearing expired entries (dir + file)`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['directory', 'file']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, undefined, { isDirectory: () => true })
        .onSecondCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledWith(path.join(cache.cacheFolderPath,  'file')));
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`skips directories while clearing expired entries (dir only)`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['directory']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, undefined, { isDirectory: () => true });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.notCalled);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail while reading file information fails`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['file1', 'file2']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, 'error')
        .onSecondCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't fail while removing expired cache entry fails`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['file']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArgWith(1, 'error');
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't remove cache entries that have been accessed in the last 24 hours`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['file1', 'file2']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: new Date() })
        .onSecondCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`doesn't remove cache entries that have been accessed in the last 24 hours (last file recently accessed)`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['file1', 'file2']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo })
        .onSecondCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: new Date() });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledOnce);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });

    it(`removes cache entries that haven't been accessed in the last 24 hours`, (done) => {
      sinon.stub(fs, 'readdir').callsArgWith(1, undefined, ['file1', 'file2']);
      const twoDaysAgo = new Date();
      twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
      sinon
        .stub(fs, 'stat')
        .onFirstCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo })
        .onSecondCall()
        .callsArgWith(1, undefined, { isDirectory: () => false, atime: twoDaysAgo });
      const unlinkStub = sinon.stub(fs, 'unlink')
        .callsFake(() => { })
        .callsArg(1);
      try {
        cache.clearExpired(() => {
          try {
            assert(unlinkStub.calledTwice);
            done();
          }
          catch (ex) {
            done(ex);
          }
        });
      }
      catch (ex) {
        done(ex);
      }
    });
  });
});