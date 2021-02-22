import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import Axios from 'axios';
import appInsights from '../../appInsights';
import auth, { AuthType } from '../../Auth';
import { Logger } from '../../cli';
import Command, { CommandError } from '../../Command';
import Utils from '../../Utils';
import commands from './commands';
const command: Command = require('./login');

describe(commands.LOGIN, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
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
      Axios.post,
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
    command.action(logger, { options: { debug: false } }, () => {
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
    command.action(logger, { options: { debug: true } }, () => {
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
    command.action(logger, { options: { debug: false, authType: 'password', userName: 'user', password: 'password' } }, () => {
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

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    command.action(logger, { options: { debug: false, authType: 'certificate', certificateFile: 'certificate' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
        assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    command.action(logger, { options: { debug: false, authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }, () => {
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

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided', (done) => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    command.action(logger, { options: { debug: false, authType: 'certificate', certificateBase64Encoded: 'certificate', thumbprint: 'thumbprint' } }, () => {
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

    command.action(logger, { options: { debug: false, authType: 'identity', userName:  'ac9fbed5-804c-4362-a369-21a4ec51109e' } }, () => {
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

    command.action(logger, { options: { debug: false, authType: 'identity' } }, () => {
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
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--authType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying userName', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying password', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--password') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if authType is set to password and userName and password not specified', () => {
    const actual = command.validate({ options: { authType: 'password' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and userName not specified', () => {
    const actual = command.validate({ options: { authType: 'password', password: 'password' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and password not specified', () => {
    const actual = command.validate({ options: { authType: 'password', userName: 'user' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and both certificateFile and certificateBase64Encoded are specified', () => {
    const actual = command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', certificateBase64Encoded: 'certificateB64', thumbprint: 'thumbprint' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and neither certificateFile nor certificateBase64Encoded are specified', () => {
    const actual = command.validate({ options: { authType: 'certificate' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and certificateFile does not exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile are specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to password and userName and password specified', () => {
    const actual = command.validate({ options: { authType: 'password', userName: 'user', password: 'password' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to deviceCode and userName and password not specified', () => {
    const actual = command.validate({ options: { authType: 'deviceCode' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is not set and userName and password not specified', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('ignores the error raised by cancelling device code auth flow', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Polling_Request_Cancelled'); });
    command.action(logger, { options: {} }, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores the error raised by cancelling device code auth flow (debug)', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Polling_Request_Cancelled'); });
    command.action(logger, { options: { debug: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('Polling_Request_Cancelled'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error in device code auth flow', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject('Error'); });
    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs in to Microsoft 365 using browser authentication', (done) => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    command.action(logger, { options: { debug: false, authType: 'browser' } }, () => {
      try {
        assert.strictEqual(auth.service.authType, AuthType.Browser, 'Incorrect authType set');
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
    command.action(logger, { options: {} }, () => {
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
    command.action(logger, { options: { debug: true } }, () => {
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
    command.action(logger, { options: { debug: true } } as any, (err?: any) => {
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