import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import Axios from 'axios';
import appInsights from '../../appInsights';
import auth, { AuthType } from '../../Auth';
import { Cli } from '../../cli/Cli';
import { CommandInfo } from '../../cli/CommandInfo';
import { Logger } from '../../cli/Logger';
import Command, { CommandError } from '../../Command';
import { sinonUtil } from '../../utils/sinonUtil';
import commands from './commands';
const command: Command = require('./login');

describe(commands.LOGIN, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(auth.service, 'logout').callsFake(() => { });
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      auth.service.logout,
      auth.ensureAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
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

  it('logs in to Microsoft 365', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    await command.action(logger, { options: { debug: false } });
    assert(auth.service.connected);
  });

  it('logs in to Microsoft 365 (debug)', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    await command.action(logger, { options: { debug: true } });
    assert(auth.service.connected);
  });

  it('logs in to Microsoft 365 using username and password when authType password set', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    await command.action(logger, { options: { debug: false, authType: 'password', userName: 'user', password: 'password' } });
    assert.strictEqual(auth.service.authType, AuthType.Password, 'Incorrect authType set');
    assert.strictEqual(auth.service.userName, 'user', 'Incorrect user name set');
    assert.strictEqual(auth.service.password, 'password', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await assert.rejects(command.action(logger, { options: { debug: false, authType: 'certificate', certificateFile: 'certificate' } }));
    assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await assert.rejects(command.action(logger, { options: { debug: false, authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }));
    assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.service.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await assert.rejects(command.action(logger, { options: { debug: false, authType: 'certificate', certificateBase64Encoded: 'certificate', thumbprint: 'thumbprint' } }));
    assert.strictEqual(auth.service.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.service.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.service.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using system managed identity when authType identity set', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { debug: false, authType: 'identity', userName:  'ac9fbed5-804c-4362-a369-21a4ec51109e' } });
    assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.service.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
  });

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { debug: false, authType: 'identity' } });
    assert.strictEqual(auth.service.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.service.userName, undefined, 'Incorrect userName set');
  });

  
  it('logs in to Microsoft 365 using client secret authType "secret" set', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    await command.action(logger, { options: { debug: false, authType: 'secret', secret: 'unBrEakaBle@123' } });
    assert.strictEqual(auth.service.authType, AuthType.Secret, 'Incorrect authType set');
    assert.strictEqual(auth.service.secret, 'unBrEakaBle@123', 'Incorrect secret set');
  });


  it('supports specifying authType', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--authType') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying userName', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--userName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying password', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--password') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if authType is set to password and userName and password not specified', async () => {
    const actual = await command.validate({ options: { authType: 'password' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and userName not specified', async () => {
    const actual = await command.validate({ options: { authType: 'password', password: 'password' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to password and password not specified', async () => {
    const actual = await command.validate({ options: { authType: 'password', userName: 'user' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and both certificateFile and certificateBase64Encoded are specified', async () => {
    const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', certificateBase64Encoded: 'certificateB64', thumbprint: 'thumbprint' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and neither certificateFile nor certificateBase64Encoded are specified', async () => {
    const actual = await command.validate({ options: { authType: 'certificate' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if authType is set to certificate and certificateFile does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile and thumbprint are specified', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to certificate and certificateFile are specified', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = await command.validate({ options: { authType: 'certificate', certificateFile: 'certificate' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to password and userName and password specified', async () => {
    const actual = await command.validate({ options: { authType: 'password', userName: 'user', password: 'password' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is set to deviceCode and userName and password not specified', async () => {
    const actual = await command.validate({ options: { authType: 'deviceCode' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if authType is not set and userName and password not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles error in device code auth flow', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Error'));
  });

  it('correctly handles error in device code auth flow (debug)', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('Error'));
  });

  it('logs in to Microsoft 365 using browser authentication', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { debug: false, authType: 'browser' } });
    assert.strictEqual(auth.service.authType, AuthType.Browser, 'Incorrect authType set');
  });

  it('correctly handles error when clearing persisted auth information', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    
    try {
      await command.action(logger, { options: {} });
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('correctly handles error when clearing persisted auth information (debug)', async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
    sinonUtil.restore(auth.clearConnectionInfo);
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    
    try {
      await command.action(logger, { options: { debug: true } });
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('correctly handles error when restoring auth information', async () => {
    sinonUtil.restore(auth.restoreAuth);
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.reject('An error has occurred'));
    try {
      await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('An error has occurred'));
    }
    finally {
      sinonUtil.restore([
        auth.clearConnectionInfo
      ]);
    }
  });

  it('fails validation if authType is set to secret and secret option is not specified', async () => {
    const actual = await command.validate({ options: { authType: 'secret' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});