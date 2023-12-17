import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth, { AuthType, CloudType } from '../../Auth.js';
import { CommandArgs, CommandError } from '../../Command.js';
import { cli } from '../../cli/cli.js';
import { CommandInfo } from '../../cli/CommandInfo.js';
import { Logger } from '../../cli/Logger.js';
import { telemetry } from '../../telemetry.js';
import { pid } from '../../utils/pid.js';
import { session } from '../../utils/session.js';
import { sinonUtil } from '../../utils/sinonUtil.js';
import commands from './commands.js';
import command from './login.js';

describe(commands.LOGIN, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'clearConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(auth.connection, 'deactivate').callsFake(() => { });
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => {
      auth.connection.name = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityName = 'alexw@contoso.com';
      auth.connection.identityId = '028de82d-7fd9-476e-a9fd-be9714280ff3';
      auth.connection.identityTenantId = 'db308122-52f3-4241-af92-1734aa6e2e50';
      return Promise.resolve('');
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      auth.connection.deactivate,
      auth.ensureAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LOGIN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('in telemetry defaults to Public cloud when no cloud has been specified', () => {
    const args: CommandArgs = { options: {} };
    command.telemetry.forEach(fn => fn(args));
    assert.strictEqual((command as any).telemetryProperties.cloud, CloudType.Public);
  });

  it('in telemetry tracks the specified cloud', () => {
    const args: CommandArgs = { options: { cloud: 'USGov' } };
    command.telemetry.forEach(fn => fn(args));
    assert.strictEqual((command as any).telemetryProperties.cloud, 'USGov');
  });

  it('logs in to Microsoft 365', async () => {
    await command.action(logger, { options: {} });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 (debug)', async () => {
    await command.action(logger, { options: { debug: true } });
    assert(auth.connection.active);
  });

  it('logs in to Microsoft 365 using username and password when authType password set', async () => {
    await command.action(logger, { options: { authType: 'password', userName: 'user', password: 'password' } });
    assert.strictEqual(auth.connection.authType, AuthType.Password, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'user', 'Incorrect user name set');
    assert.strictEqual(auth.connection.password, 'password', 'Incorrect password set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateFile is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await command.action(logger, { options: { authType: 'certificate', certificateFile: 'certificate' } });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set with thumbprint', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await command.action(logger, { options: { authType: 'certificate', certificateFile: 'certificate', thumbprint: 'thumbprint' } });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using certificate when authType certificate set and certificateBase64Encoded is provided', async () => {
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await command.action(logger, { options: { authType: 'certificate', certificateBase64Encoded: 'certificate', thumbprint: 'thumbprint' } });
    assert.strictEqual(auth.connection.authType, AuthType.Certificate, 'Incorrect authType set');
    assert.strictEqual(auth.connection.certificate, 'certificate', 'Incorrect certificate set');
    assert.strictEqual(auth.connection.thumbprint, 'thumbprint', 'Incorrect thumbprint set');
  });

  it('logs in to Microsoft 365 using system managed identity when authType identity set', async () => {
    await command.action(logger, { options: { authType: 'identity', userName: 'ac9fbed5-804c-4362-a369-21a4ec51109e' } });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, 'ac9fbed5-804c-4362-a369-21a4ec51109e', 'Incorrect userName set');
  });

  it('logs in to Microsoft 365 using user-assigned managed identity when authType identity set', async () => {
    await command.action(logger, { options: { authType: 'identity' } });
    assert.strictEqual(auth.connection.authType, AuthType.Identity, 'Incorrect authType set');
    assert.strictEqual(auth.connection.userName, undefined, 'Incorrect userName set');
  });


  it('logs in to Microsoft 365 using client secret authType "secret" set', async () => {
    await command.action(logger, { options: { authType: 'secret', secret: 'unBrEakaBle@123' } });
    assert.strictEqual(auth.connection.authType, AuthType.Secret, 'Incorrect authType set');
    assert.strictEqual(auth.connection.secret, 'unBrEakaBle@123', 'Incorrect secret set');
  });

  it('logs in to Microsoft 365 using the specified cloud', async () => {
    await command.action(logger, { options: { cloud: 'USGov' } });
    assert.strictEqual(auth.connection.cloudType, CloudType.USGov);
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

  it('fails validation if invalid authType specified', async () => {
    const actual = await command.validate({ options: { authType: 'invalid authType' } }, commandInfo);
    assert.notStrictEqual(actual, true);
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

  it('fails validation cloud is set to an invalid value', async () => {
    const actual = await command.validate({ options: { cloud: 'invalid' } }, commandInfo);
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

  it('passes validation when cloud is set to Public', async () => {
    const actual = await command.validate({ options: { cloud: 'Public' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGov', async () => {
    const actual = await command.validate({ options: { cloud: 'USGov' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGovHigh', async () => {
    const actual = await command.validate({ options: { cloud: 'USGovHigh' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to USGovDoD', async () => {
    const actual = await command.validate({ options: { cloud: 'USGovDoD' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cloud is set to China', async () => {
    const actual = await command.validate({ options: { cloud: 'China' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly handles error in device code auth flow', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Error'));
  });

  it('correctly handles error in device code auth flow (debug)', async () => {
    sinonUtil.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error')); });
    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError('Error'));
  });

  it('logs in to Microsoft 365 using browser authentication', async () => {
    await command.action(logger, { options: { authType: 'browser' } });
    assert.strictEqual(auth.connection.authType, AuthType.Browser, 'Incorrect authType set');
  });

  it(`Don't start login flow when the CLI is signed in`, async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { ensure: true } });
    assert.strictEqual(auth.service.connected, true);
  });

  it(`Don't start login flow if the CLI is signed in as a user`, async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { ensure: true, authType: 'password', userName: 'john.doe@contoso.com', password: 'password' } });
    assert.strictEqual(auth.service.connected, true);
  });

  it(`Don't start login flow, if the CLI is signed in using a certificate`, async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));
    sinon.stub(fs, 'readFileSync').callsFake(() => 'certificate');

    await command.action(logger, { options: { ensure: true, authType: 'certificate ', certificateFile: 'certificate' } });
    assert.strictEqual(auth.service.connected, true);
  });

  it(`Don't start login flow if the CLI is signed in using Managed Identity`, async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { ensure: true, authType: 'identity' } });
    assert.strictEqual(auth.service.connected, true);
  });

  it(`Don't start login flow if the CLI is signed in using the specified app and to the specified tenant`, async () => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve(''));

    await command.action(logger, { options: { ensure: true, appId: '1cf21ca6-c8f0-4a21-839d-68a09d3a0f55', tenant: '973fce64-6409-4843-9328-c2cef0427f4e' } });
    assert.strictEqual(auth.service.connected, true);
  });

  it('correctly handles error when clearing persisted auth information', async () => {
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
