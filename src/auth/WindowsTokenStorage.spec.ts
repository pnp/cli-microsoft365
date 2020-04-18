import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from '../Utils';
import { WindowsTokenStorage } from './WindowsTokenStorage';
import * as childProcess from 'child_process';
import * as path from 'path';
import * as os from 'os';

describe('WindowsTokenStorage', () => {
  const windowsCredsManager = new WindowsTokenStorage();
  const prefix: string = 'Office365Cli:target=';
  const prefixShort: string = 'Office365Cli';

  it('executes right command to get password from Windows Credential Manager', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a as any; return {} as childProcess.ChildProcess; });
    windowsCredsManager.get();
    try {
      assert.equal(file, path.join(__dirname, '../../bin/windows/creds.exe'));
      assert.deepEqual(args, [
        '-s',
        '-g',
        '-t', `${prefix}${prefixShort}*`
      ]);
      done();
    }
    catch (e) {
      done(e);
    }
    finally {
      Utils.restore(childProcess.execFile);
    }
  });

  it('correctly handles error when getting password from Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, { message: 'An error has occurred' });
    windowsCredsManager
      .get()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('correctly handles no credentials found in Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, '', null);
    windowsCredsManager
      .get()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'Credential not found');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('returns from Windows Credential Manager credential consisting of a single chunk', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, [
      'Target Name: SPO',
      'Credential: ' + Buffer.from('ABC', 'utf8').toString('hex')
    ].join(os.EOL), null);
    windowsCredsManager
      .get()
      .then((password) => {
        try {
          assert.equal(password, 'ABC');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      }, (error: any) => {
        Utils.restore(childProcess.execFile);
        done('Expected pass but failed instead');
      });
  });

  it('returns from Windows Credential Manager credential consisting of multiple chunks', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, [
      'Target Name: SPO--1-2',
      'Credential: ' + Buffer.from('ABC', 'utf8').toString('hex'),
      '',
      'Target Name: SPO--2-2',
      'Credential: ' + Buffer.from('DEF', 'utf8').toString('hex')
    ].join(os.EOL), null);
    windowsCredsManager
      .get()
      .then((password) => {
        try {
          assert.equal(password, 'ABCDEF');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      }, (error: any) => {
        Utils.restore(childProcess.execFile);
        done('Expected pass but failed instead');
      });
  });

  it('correctly handles error when incorrect number password chunks retrieved from Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, [
      'Target Name: SPO--1-2',
      'Credential: ' + Buffer.from('ABC', 'utf8').toString('hex')
    ].join(os.EOL), null);
    windowsCredsManager
      .get()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected fail but passed instead');
      }, (error: any) => {
        try {
          assert.equal(error, `Couldn't load all credential chunks. Expected 2, found 1`);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('correctly handles invalid chunk number retrieved from Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, [
      'Target Name: SPO--1-2',
      'Credential: ABC',
      '',
      'Target Name: SPO--a-2',
      'Credential: DEF',
    ].join(os.EOL), null);
    windowsCredsManager
      .get()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected fail but passed instead');
      }, (error: any) => {
        try {
          assert.equal(error, `Couldn't load all credential chunks. Expected 2, found 1`);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('correctly handles password chunk missing in Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, [
      'Target Name: SPO--1-3',
      'Credential: ' + Buffer.from('ABC', 'utf8').toString('hex'),
      '',
      'Target Name: SPO--3-3',
      'Credential: ' + Buffer.from('GHI', 'utf8').toString('hex')
    ].join(os.EOL), null);
    windowsCredsManager
      .get()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected fail but passed instead');
      }, (error: any) => {
        try {
          assert.equal(error, `Missing chunk 2/3`);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('clears existing passwords before setting new one in Windows Credential Manager', (done) => {
    const removeStub = sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.reject('An error has occurred when removing existing passwords'));
    windowsCredsManager
      .set('ABC')
      .then(() => {
        try {
          assert(removeStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(windowsCredsManager.remove);
        }
      }, (error) => {
        try {
          assert(removeStub.called);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(windowsCredsManager.remove);
        }
      });
  });

  it('correctly handles error when clearing existing passwords in Windows Credential Manager', (done) => {
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.reject('An error has occurred when removing existing passwords'));
    windowsCredsManager
      .set('ABC')
      .then(() => {
        Utils.restore(windowsCredsManager.remove);
        done('Fail expected but passed instead');
      }, (error) => {
        try {
          assert.equal(error, 'An error has occurred when removing existing passwords');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(windowsCredsManager.remove);
        }
      });
  });

  it('doesn\'t set new password if clearing existing passwords in Windows Credential Manager failed', (done) => {
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.reject('An error has occurred when removing existing passwords'));
    const execFileStub = sinon.stub(childProcess, 'execFile').callsFake(() => { return {} as childProcess.ChildProcess; });
    windowsCredsManager
      .set('ABC')
      .then(() => {
        try {
          assert(execFileStub.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            windowsCredsManager.remove,
            childProcess.execFile
          ]);
        }
      }, (error) => {
        try {
          assert(execFileStub.notCalled);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            windowsCredsManager.remove,
            childProcess.execFile
          ]);
        }
      });
  });

  it('executes right command to set password in Windows Credential Manager', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a as any; return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .set('ABC')
      .then(() => {
        try {
          assert.equal(file, path.join(__dirname, '../../bin/windows/creds.exe'));
          assert.deepEqual(args, [
            '-a',
            '-t', `${prefix}${prefixShort}`,
            '-p', Buffer.from('ABC', 'utf8').toString('hex')
          ]);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            childProcess.execFile,
            windowsCredsManager.remove
          ]);
        }
      });
  });

  it('continues when adding short password to Windows Credential Manager succeeded', (done) => {
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(childProcess, 'execFile').callsFake(() => { return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .set('ABC')
      .then(() => {
        Utils.restore([
          childProcess.execFile,
          windowsCredsManager.remove
        ]);
        done();
      }, () => {
        Utils.restore([
          childProcess.execFile,
          windowsCredsManager.remove
        ]);
        done('Passed expected but failed instead');
      });
  });

  it('continues when adding long password to Windows Credential Manager succeeded', (done) => {
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(childProcess, 'execFile').callsFake(() => { return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .set(new Array(3000).fill('x').join(''))
      .then(() => {
        Utils.restore([
          childProcess.execFile,
          windowsCredsManager.remove
        ]);
        done();
      }, () => {
        Utils.restore([
          childProcess.execFile,
          windowsCredsManager.remove
        ]);
        done('Passed expected but failed instead');
      });
  });

  it('creates correct number of chunks for long password to be stored in Windows Credential Manager', (done) => {
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    const execFileStub = sinon.stub(childProcess, 'execFile').callsFake(() => { return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .set(new Array(3000).fill('x').join(''))
      .then(() => {
        try {
          assert(execFileStub.calledTwice);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            childProcess.execFile,
            windowsCredsManager.remove
          ]);
        }
      });
  });

  it('correctly names chunks for long password to be stored in Windows Credential Manager', (done) => {
    const names: string[] = [];
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { names.push((a as string[])[2]); return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .set(new Array(3000).fill('x').join(''))
      .then(() => {
        try {
          assert.equal(names[0], 'Office365Cli:target=Office365Cli--1-2');
          assert.equal(names[1], 'Office365Cli:target=Office365Cli--2-2');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            childProcess.execFile,
            windowsCredsManager.remove
          ]);
        }
      });
  });

  it('correctly handles error when storing password in Windows Credential Manager', (done) => {
    const names: string[] = [];
    sinon.stub(windowsCredsManager, 'remove').callsFake(() => Promise.resolve());
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { names.push((a as string[])[2]); return {} as childProcess.ChildProcess; }).callsArgWith(2, { message: 'An error has occurred' }, null, null);
    windowsCredsManager
      .set('ABC')
      .then(() => {
        Utils.restore([
          childProcess.execFile,
          windowsCredsManager.remove
        ]);
        done('Fail expected but passed instead');
      }, (err) => {
        try {
          assert.equal(err, 'Could not add password to credential store: An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            childProcess.execFile,
            windowsCredsManager.remove
          ]);
        }
      });
  });

  it('executes right command to remove password from Windows Credential Manager', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a as any; return {} as childProcess.ChildProcess; }).callsArgWith(2, null, null, null);
    windowsCredsManager
      .remove()
      .then(() => {
        try {
          assert.equal(file, path.join(__dirname, '../../bin/windows/creds.exe'));
          assert.deepEqual(args, [
            '-d',
            '-g',
            '-t', `${prefix}${prefixShort}*`
          ]);
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore([
            childProcess.execFile,
            windowsCredsManager.remove
          ]);
        }
      });
  });

  it('correctly handles error when removing password from Windows Credential Manager', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, { message: 'An error has occurred' });
    windowsCredsManager
      .remove()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'Could not remove password from credential store: An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
        finally {
          Utils.restore(childProcess.execFile);
        }
      });
  });

  it('completes when removing password from Windows Credential Manager succeeded', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, null, null);
    windowsCredsManager
      .remove()
      .then(() => {
        Utils.restore(childProcess.execFile);
        done();
      }, (error: any) => {
        Utils.restore(childProcess.execFile);
        done('Expected pass but failed');
      });
  });
});