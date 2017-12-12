import * as sinon from 'sinon';
import * as assert from 'assert';
import Utils from '../Utils';
import { KeychainTokenStorage } from './KeychainTokenStorage';
import * as childProcess from 'child_process';

describe('KeychainTokenStorage', () => {
  const keychain = new KeychainTokenStorage();

  it('executes right command to get password from Keychain', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a });
    const service = 'mock';
    keychain.get(service);
    try {
      assert.equal(file, '/usr/bin/security');
      assert.deepEqual(args, [
          'find-generic-password',
          '-a', service,
          '-s', service,
          '-D', 'Office 365 CLI',
          '-g'
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

  it('correctly handles error when getting password from Keychain', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, { message: 'An error has occurred' });
    keychain
      .get('mock')
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

  it('correctly handles error when something else than password returned from Keychain', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, '', 'random output');
    keychain
      .get('mock')
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'Password in invalid format');
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

  it('returns password retrieved from Keychain', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, '', 'password: "ABC"');
    keychain
      .get('mock')
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
        done('Expected pass but failed');        
      });
  });

  it('executes right command to set password in Keychain', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a });
    const service = 'mock';
    const token = 'ABC';
    keychain.set(service, token);
    try {
      assert.equal(file, '/usr/bin/security');
      assert.deepEqual(args, [
          'add-generic-password',
          '-a', service,
          '-s', service,
          '-D', 'Office 365 CLI',
          '-w', token,
          '-U'
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

  it('correctly handles error when setting password in Keychain', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, { message: 'An error has occurred' });
    keychain
      .set('mock', 'ABC')
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'Could not add password to keychain: An error has occurred');
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

  it('completes when setting password in Keychain succeeded', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, null, null);
    keychain
      .set('mock', 'ABC')
      .then(() => {
        Utils.restore(childProcess.execFile);
        done();
      }, (error: any) => {
        Utils.restore(childProcess.execFile);
        done('Expected pass but failed');
      });
  });

  it('executes right command to remove password from Keychain', (done) => {
    let file = '';
    let args: string[] = [];
    sinon.stub(childProcess, 'execFile').callsFake((f, a) => { file = f; args = a });
    const service = 'mock';
    keychain.remove(service);
    try {
      assert.equal(file, '/usr/bin/security');
      assert.deepEqual(args, [
          'delete-generic-password',
          '-a', service,
          '-s', service,
          '-D', 'Office 365 CLI'
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

  it('correctly handles error when removing password from Keychain', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, { message: 'An error has occurred' });
    keychain
      .remove('mock')
      .then(() => {
        Utils.restore(childProcess.execFile);
        done('Expected failure but passed');
      }, (error: any) => {
        try {
          assert.equal(error, 'Could not remove account from keychain: An error has occurred');
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

  it('completes when removing password from Keychain succeeded', (done) => {
    sinon.stub(childProcess, 'execFile').callsArgWith(2, null, null, null);
    keychain
      .remove('mock')
      .then(() => {
        Utils.restore(childProcess.execFile);
        done();
      }, (error: any) => {
        Utils.restore(childProcess.execFile);
        done('Expected pass but failed');
      });
  });
});