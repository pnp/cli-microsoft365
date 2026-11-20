import assert from 'assert';
import child_process from 'child_process';
import fs from 'fs';
import os from 'os';
import path from 'path';
import sinon from 'sinon';
import { cache } from './cache.js';
import { pid } from './pid.js';
import { sinonUtil } from './sinonUtil.js';

describe('utils/pid', () => {
  let cacheSetValueStub: sinon.SinonStub;

  before(() => {
    sinon.stub(cache, 'getValue').returns(undefined);
    cacheSetValueStub = sinon.stub(cache, 'setValue').returns(undefined);
  });

  afterEach(() => {
    sinonUtil.restore([
      os.platform,
      child_process.spawnSync,
      fs.existsSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('retrieves process name on Windows', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","123","Console","1","98,765 K"\r\n' } as any);

    assert.strictEqual(pid.getProcessName(123), 'pwsh.exe');
  });

  it('runs tasklist from the Windows system directory specified in SystemRoot', () => {
    const systemRoot = process.env.SystemRoot;
    process.env.SystemRoot = 'D:\\Windows';
    sinon.stub(os, 'platform').returns('win32');
    const spawnSyncStub = sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","123","Console","1","98,765 K"\r\n' } as any);

    try {
      pid.getProcessName(123);
      assert.strictEqual(spawnSyncStub.firstCall.args[0], path.join('D:\\Windows', 'System32', 'tasklist.exe'));
      assert.deepStrictEqual(spawnSyncStub.firstCall.args[1], ['/FI', 'PID eq 123', '/FO', 'CSV', '/NH']);
    }
    finally {
      if (systemRoot === undefined) {
        delete process.env.SystemRoot;
      }
      else {
        process.env.SystemRoot = systemRoot;
      }
    }
  });

  it('runs tasklist from the default Windows system directory when SystemRoot is not set', () => {
    const systemRoot = process.env.SystemRoot;
    delete process.env.SystemRoot;
    sinon.stub(os, 'platform').returns('win32');
    const spawnSyncStub = sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","123","Console","1","98,765 K"\r\n' } as any);

    try {
      pid.getProcessName(123);
      assert.strictEqual(spawnSyncStub.firstCall.args[0], path.join('C:\\Windows', 'System32', 'tasklist.exe'));
    }
    finally {
      if (systemRoot !== undefined) {
        process.env.SystemRoot = systemRoot;
      }
    }
  });

  it('retrieves process name on macOS', () => {
    sinon.stub(os, 'platform').returns('darwin');
    sinon.stub(child_process, 'spawnSync').returns({ stdout: '/bin/bash' } as any);

    assert.strictEqual(pid.getProcessName(123), '/bin/bash');
  });

  it('retrieves undefined on macOS when retrieving the process name failed', () => {
    sinon.stub(os, 'platform').returns('darwin');
    sinon.stub(child_process, 'spawnSync').returns({ error: 'An error has occurred' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('retrieves process name on Linux', () => {
    sinon.stub(os, 'platform').returns('linux');
    sinon.stub(fs, 'existsSync').returns(true);
    sinon.stub(fs, 'readFileSync').returns('(pwsh)');

    assert.strictEqual(pid.getProcessName(123), 'pwsh');
  });

  it(`returns undefined on Linux if the process is not found`, () => {
    sinon.stub(os, 'platform').returns('linux');
    sinon.stub(fs, 'existsSync').returns(false);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined name on other platforms', () => {
    sinon.stub(os, 'platform').returns('android');

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined when retrieving process name on Windows fails', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ error: 'An error has occurred' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('runs tasklist from the default Windows system directory when SystemRoot is empty', () => {
    const systemRoot = process.env.SystemRoot;
    process.env.SystemRoot = '';
    sinon.stub(os, 'platform').returns('win32');
    const spawnSyncStub = sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","123","Console","1","98,765 K"\r\n' } as any);

    try {
      pid.getProcessName(123);
      assert.strictEqual(spawnSyncStub.firstCall.args[0], path.join('C:\\Windows', 'System32', 'tasklist.exe'));
    }
    finally {
      if (systemRoot === undefined) {
        delete process.env.SystemRoot;
      }
      else {
        process.env.SystemRoot = systemRoot;
      }
    }
  });

  it('returns undefined when tasklist exits with a non-zero code on Windows', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ status: 1, stdout: '' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined when the process is not found on Windows', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: 'INFO: No tasks are running which match the specified criteria.\r\n' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('returns undefined when tasklist returns a different process on Windows', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","1234","Console","1","98,765 K"\r\n' } as any);

    assert.strictEqual(pid.getProcessName(123), undefined);
  });

  it('stores retrieved process name in cache', () => {
    sinon.stub(os, 'platform').returns('win32');
    sinon.stub(child_process, 'spawnSync').returns({ status: 0, stdout: '"pwsh.exe","123","Console","1","98,765 K"\r\n' } as any);

    pid.getProcessName(123);

    assert(cacheSetValueStub.called);
  });

  it('retrieves process name from cache when available', () => {
    sinonUtil.restore(cache.getValue);
    sinon.stub(cache, 'getValue').returns('pwsh');
    const osPlatformSpy = sinon.spy(os, 'platform');

    assert.strictEqual(pid.getProcessName(123), 'pwsh');
    assert(osPlatformSpy.notCalled);
  });
});