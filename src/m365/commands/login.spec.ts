import commands from './commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../Command';
import * as sinon from 'sinon';
import appInsights from '../../appInsights';
import auth from '../../Auth';
const command: Command = require('./login');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../Utils';
import { AuthType } from '../../Auth';
import * as fs from 'fs';

describe(commands.LOGIN, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      action: command.action(),
      commandWrapper: {
        command: 'login'
      },
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    sinon.stub(auth.service, 'logout').callsFake(() => { });
  });

  afterEach(() => {
    Utils.restore([
      auth.cancel,
      fs.existsSync,
      fs.readFileSync,
      auth.service.logout,
      auth.ensureAccessToken
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      auth.clearConnectionInfo,
      auth.storeConnectionInfo,
      request.post,
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('logs in to Microsoft 365', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(auth.service.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 (debug)', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(auth.service.connected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using username and password when authType password set', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    cmdInstance.action({ options: { debug: false, authType: 'password', userName: 'user', password: 'password' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Password, 'Incorrect authType set');
        assert.strictEqual(auth.service.userName, 'user', 'Incorrect user name set');
        assert.strictEqual(auth.service.password, 'password', 'Incorrect password set');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    cmdInstance.action({ options: { debug: false, authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
        assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
        assert.strictEqual(auth.service.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using system managed identity when authType identity set', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    cmdInstance.action({ options: { debug: false, authType: 'identity', userName:  'ac9fbed5-804c-4362-a369-21a4ec51109e' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
        assert.strictEqual(auth.service.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    cmdInstance.action({ options: { debug: false, authType: 'identity' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
        assert.strictEqual(auth.service.userName, undefined, 'Incorrect userName set');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying authType', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--authType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying userName', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying password', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--password') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if authType is set to password and userName and password not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'password' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and userName not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'password', password: 'password' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and password not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'password', userName: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and certificateFile not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'certificate', thumbprint: 'thumbprint' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and certificateFile does not exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and thumbprint not specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'certificate', certificateFile: 'certificate' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to password and userName and password specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'password', userName: 'user', password: 'password' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to deviceCode and userName and password not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { authType: 'deviceCode' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is not set and userName and password not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('ignores the error raised by cancelling device code auth flow', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Polling_Request_Cancelled'); });
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores the error raised by cancelling device code auth flow (debug)', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Polling_Request_Cancelled'); });
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('Polling_Request_Cancelled'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error in device code auth flow', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Error'); });
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when clearing persisted auth information', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    Utils.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo
        ]);
      }
    });
  });

  it('correctly handles error when clearing persisted auth information (debug)', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    Utils.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo
        ]);
      }
    });
  });

  it('correctly handles error when restoring auth information', (done) => {
    Utils.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          auth.clearConnectionInfo
        ]);
      }
    });
  });
});